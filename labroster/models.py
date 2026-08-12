"""The things a laboratory roster is made of.

Deliberately free of spreadsheet and scheduling code: these objects describe the
workforce and the requirements, and answer questions about themselves.  Reading
them from a workbook lives in ``workbook.py``, and deciding who works when lives
in ``scheduler.py``.

Three distinctions matter here and are kept strictly separate, because conflating
them is what makes a rota misleading:

* **grade** (NHS band or local grade) is about pay and responsibility
* **seniority / coordination** is about running a shift
* **competency** is about being safe to work in a particular discipline

A Band 7 is not automatically competent in morphology, and someone competent in
morphology is not automatically a shift coordinator.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta

from .timeutils import (
    TimeError, duration_hours, pattern_week, shift_window,
)

WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
                 "Saturday", "Sunday"]
WEEKDAY_SHORT = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


# --------------------------------------------------------------------------
# competency
# --------------------------------------------------------------------------

class CompetencyStatus:
    """Where somebody has got to with a particular competency.

    Only ``COMPETENT``, ``TRAINER`` and ``ASSESSOR`` mean a person can satisfy a
    requirement on their own.  ``IN_TRAINING`` and ``SUPERVISED`` explicitly do
    not: someone still being supervised needs somebody else to be there, so
    counting them as cover would overstate what the laboratory can actually do.
    """

    NOT_TRAINED = "Not Trained"
    IN_TRAINING = "In Training"
    SUPERVISED = "Supervised"
    COMPETENT = "Competent"
    TRAINER = "Trainer"
    ASSESSOR = "Assessor"
    EXPIRED = "Expired"

    ALL = [NOT_TRAINED, IN_TRAINING, SUPERVISED, COMPETENT, TRAINER,
           ASSESSOR, EXPIRED]

    #: Statuses where the person can work in the discipline unsupervised.
    INDEPENDENT = {COMPETENT, TRAINER, ASSESSOR}

    #: Statuses that mean the person is still learning and needs support.
    LEARNING = {IN_TRAINING, SUPERVISED}

    @classmethod
    def normalise(cls, text) -> str | None:
        """Match a spreadsheet value to a known status, tolerating variations."""
        target = re.sub(r"[^a-z]", "", str(text or "").lower())
        for status in cls.ALL:
            if re.sub(r"[^a-z]", "", status.lower()) == target:
                return status
        aliases = {
            "nottrainednotapplicable": cls.NOT_TRAINED,
            "none": cls.NOT_TRAINED,
            "na": cls.NOT_TRAINED,
            "training": cls.IN_TRAINING,
            "trainee": cls.IN_TRAINING,
            "undersupervision": cls.SUPERVISED,
            "signedoff": cls.COMPETENT,
            "independent": cls.COMPETENT,
            "lapsed": cls.EXPIRED,
        }
        return aliases.get(target)


@dataclass
class Competency:
    """One person's standing in one discipline."""

    staff_id: str
    discipline: str                       # BT, HAEM, COAG, MORPH …
    name: str = ""                        # e.g. "Electronic issue"
    status: str = CompetencyStatus.NOT_TRAINED
    date_achieved: date | None = None
    review_date: date | None = None
    expiry_date: date | None = None
    trainer: bool = False
    assessor: bool = False
    authoriser: bool = False              # may authorise results in this discipline
    notes: str = ""

    def has_expired(self, as_of: date) -> bool:
        """Expiry is a fact about a date, not a status somebody remembered to set."""
        if self.status == CompetencyStatus.EXPIRED:
            return True
        return self.expiry_date is not None and self.expiry_date < as_of

    def effective_status(self, as_of: date) -> str:
        return CompetencyStatus.EXPIRED if self.has_expired(as_of) else self.status

    def is_independent(self, as_of: date) -> bool:
        """Can satisfy a requirement for a competent member of staff, alone."""
        return self.effective_status(as_of) in CompetencyStatus.INDEPENDENT

    def is_learning(self, as_of: date) -> bool:
        return self.effective_status(as_of) in CompetencyStatus.LEARNING

    def can_authorise(self, as_of: date) -> bool:
        """Authorisation needs the flag *and* current independent competence."""
        return self.authoriser and self.is_independent(as_of)

    def can_train(self, as_of: date) -> bool:
        return ((self.trainer or self.status == CompetencyStatus.TRAINER)
                and self.is_independent(as_of))

    def can_assess(self, as_of: date) -> bool:
        return ((self.assessor or self.status == CompetencyStatus.ASSESSOR)
                and self.is_independent(as_of))

    def days_until_expiry(self, as_of: date) -> int | None:
        if self.expiry_date is None:
            return None
        return (self.expiry_date - as_of).days


# --------------------------------------------------------------------------
# staff
# --------------------------------------------------------------------------

@dataclass
class Availability:
    """When somebody is contracted to be available.

    ``weekdays`` maps a week in the repeating cycle (1-based) to the weekday
    numbers that person works, using Python's Monday=0 convention.  A full-time
    member of staff has a one-week cycle listing Monday to Friday; somebody on an
    alternating pattern has a two-week cycle with different days in each.
    """

    cycle_weeks: int = 1
    weekdays: dict[int, set[int]] = field(default_factory=dict)
    earliest_start: time | None = None
    latest_finish: time | None = None
    max_days_per_week: int = 0            # 0 = no personal limit
    excluded_weekdays: set[int] = field(default_factory=set)

    def works_weekday(self, day: date, anchor: date) -> bool:
        """Is this person contracted to work on this calendar day at all?"""
        if day.weekday() in self.excluded_weekdays:
            return False
        if not self.weekdays:
            return True                    # nothing specified: treat as flexible
        week = pattern_week(day, anchor, self.cycle_weeks)
        allowed = self.weekdays.get(week)
        if allowed is None:
            # Cycle declared but this week left blank: fall back to week 1 so a
            # half-filled pattern does not silently make somebody unavailable.
            allowed = self.weekdays.get(1, set())
        return day.weekday() in allowed

    def permits_times(self, start: time, end: time) -> bool:
        """Does a shift fit inside this person's available hours?"""
        if self.earliest_start is not None and start < self.earliest_start:
            return False
        if self.latest_finish is not None:
            # A shift running past midnight can never fit inside a same-day
            # latest-finish limit.
            if end <= start or end > self.latest_finish:
                return False
        return True


@dataclass
class Staff:
    """A member of the laboratory workforce."""

    staff_id: str
    name: str
    job_title: str = ""
    band: str = ""
    registered: bool = True               # e.g. HCPC-registered BMS
    is_senior: bool = False               # explicit, not inferred from band
    shift_coordinator: bool = False
    trainee: bool = False

    contracted_weekly_hours: float = 0.0
    fte: float = 1.0
    working_pattern: str = ""
    max_period_hours: float = 0.0         # 0 = no hard ceiling, fairness governs
    min_period_hours: float = 0.0         # 0 = derive from contracted hours
    max_weekly_hours: float = 0.0         # 0 = no weekly ceiling

    availability: Availability = field(default_factory=Availability)
    nights_ok: bool = True
    weekends_ok: bool = True
    max_nights: int = 0                   # 0 = no personal limit
    max_weekends: int = 0
    max_consecutive_days: int = 0         # 0 = use the organisational rule
    restrictions: str = ""

    group: str = "Main"
    notes: str = ""

    # Filled in by the scheduler / reporting layer.
    target_period_hours: float = 0.0
    allocated_hours: float = 0.0          # hours actually worked on shifts
    credited_absence_hours: float = 0.0   # hours credited for absence

    @property
    def band_value(self) -> float:
        """Numeric band for comparisons: '8a' sorts above '8', which sorts above '7'."""
        text = str(self.band).strip().lower().replace("band", "").strip()
        match = re.match(r"(\d+)\s*([a-d])?", text)
        if not match:
            return 0.0
        value = float(match.group(1))
        if match.group(2):
            value += (ord(match.group(2)) - ord("a") + 1) / 10.0
        return value

    @property
    def total_accounted_hours(self) -> float:
        """Worked hours plus hours credited for absence.

        This is what workload balancing compares, so a week of annual leave
        counts towards somebody's contracted hours instead of leaving a deficit
        the roster would try to fill with extra shifts.
        """
        return round(self.allocated_hours + self.credited_absence_hours, 2)

    @property
    def hours_variance(self) -> float:
        """Total accounted hours minus target.  Negative means under-used."""
        return round(self.total_accounted_hours - self.target_period_hours, 2)

    @property
    def percent_of_target(self) -> float | None:
        if not self.target_period_hours:
            return None
        return round(100.0 * self.total_accounted_hours
                     / self.target_period_hours, 1)

    @property
    def working_days_per_week(self) -> int:
        """How many days a week this person normally works.

        Taken from the working pattern where one is set, averaged across the
        cycle; otherwise inferred from contracted hours against a standard day.
        """
        pattern = self.availability.weekdays
        if pattern:
            counts = [len(days) for days in pattern.values() if days]
            if counts:
                return max(1, round(sum(counts) / len(counts)))
        if self.contracted_weekly_hours:
            return max(1, min(7, round(self.contracted_weekly_hours / 7.5)))
        return 5

    @property
    def normal_daily_hours(self) -> float:
        """Hours in one of this person's normal working days."""
        weekly = self.contracted_weekly_hours or (37.5 * (self.fte or 1.0))
        return round(weekly / self.working_days_per_week, 4)

    def meets_band(self, threshold: float) -> bool:
        return self.band_value >= threshold


# --------------------------------------------------------------------------
# shifts, requirements, benches
# --------------------------------------------------------------------------

def _applies_to_day(scope: str, day: date, weekend_days: set[int]) -> bool:
    """Shared day-scope test used by shifts, requirements and benches."""
    key = re.sub(r"[^a-z]", "", str(scope or "all").lower())
    weekend = day.weekday() in weekend_days
    if key in ("", "all", "everyday", "daily", "everyday"):
        return True
    if key.startswith("weekend"):
        return weekend
    if key.startswith("weekday"):
        return not weekend
    for index, name in enumerate(WEEKDAY_NAMES):
        if key == name.lower() or key == WEEKDAY_SHORT[index].lower():
            return day.weekday() == index
    return True


@dataclass
class ShiftType:
    """A kind of shift, defined by real clock times."""

    code: str
    name: str
    start: time
    end: time
    days: str = "All"
    is_night: bool = False
    colour: str = "D9E1F2"
    font_colour: str = "000000"
    counts_as_weekend: bool = False

    @property
    def hours(self) -> float:
        """Length in hours, derived from the times rather than typed in."""
        return duration_hours(self.start, self.end, f"{self.name} shift")

    @property
    def crosses_midnight(self) -> bool:
        return self.end <= self.start

    def window(self, day: date) -> tuple[datetime, datetime]:
        return shift_window(day, self.start, self.end)

    def applies_on(self, day: date, weekend_days: set[int]) -> bool:
        return _applies_to_day(self.days, day, weekend_days)

    @property
    def times_label(self) -> str:
        return f"{self.start.strftime('%H:%M')}–{self.end.strftime('%H:%M')}"

    @property
    def label(self) -> str:
        return f"{self.name} ({self.times_label})"


@dataclass
class ShiftRequirement:
    """What a shift needs in order to be safe to run.

    Every condition is checked independently.  A senior member of staff does not
    satisfy a transfusion competency requirement, and a transfusion-competent
    scientist does not satisfy a requirement for a shift coordinator.
    """

    shift_code: str
    days: str = "All"
    min_staff: int = 1
    min_registered: int = 0
    min_senior: int = 0
    min_band: float = 0.0
    min_at_band: int = 0
    min_coordinators: int = 0
    min_trainers: int = 0
    max_trainees: int = 0                 # 0 = no limit
    required_competencies: dict[str, int] = field(default_factory=dict)
    required_authorisers: dict[str, int] = field(default_factory=dict)
    notes: str = ""

    def applies_on(self, day: date, weekend_days: set[int]) -> bool:
        return _applies_to_day(self.days, day, weekend_days)


@dataclass
class Bench:
    """A laboratory section that needs somebody actually standing at it.

    ``start``/``end`` are optional and unused for whole-shift allocation, but
    they are here so that part-shift cover (morphology 09:00–13:00, coagulation
    13:00–17:00) can be added later without reshaping the data model.
    """

    name: str
    discipline: str
    days: str = "All"
    min_staff: int = 1
    min_weekend: int | None = None
    shift_codes: list[str] = field(default_factory=list)   # empty = all day shifts
    start: time | None = None
    end: time | None = None
    requires_authoriser: bool = False
    target_rotation_days: int = 0         # 0 = no rotation expectation

    def applies_on(self, day: date, weekend_days: set[int]) -> bool:
        return _applies_to_day(self.days, day, weekend_days)

    def required_on(self, day: date, weekend_days: set[int]) -> int:
        """Weekends usually run a reduced service, so they get their own minimum."""
        if day.weekday() in weekend_days and self.min_weekend is not None:
            return self.min_weekend
        return self.min_staff

    def covers_shift(self, shift: ShiftType) -> bool:
        if self.shift_codes:
            return shift.code in self.shift_codes
        return not shift.is_night


@dataclass
class LeaveEntry:
    staff_id: str
    start: date
    end: date
    code: str
    reason: str = ""
    credited_hours: float | None = None
    """Hours credited for the whole absence, when a manager wants to state it
    explicitly rather than have it derived from the working pattern."""


#: How the hours credited for a day of absence are worked out.
CREDIT_FROM_PATTERN = "Working pattern"
CREDIT_FIXED = "Fixed hours per day"
CREDIT_NONE = "Not credited"
CREDIT_METHODS = [CREDIT_FROM_PATTERN, CREDIT_FIXED, CREDIT_NONE]


@dataclass
class LeaveType:
    """A kind of absence, and how it is treated for contracted hours.

    ``credits_hours`` is what stops the roster trying to claw back hours somebody
    was away for.  If a week of annual leave is credited, that person's accounted
    hours already reach their contracted target, so the scheduler has no deficit
    to close and will not hand them extra shifts to make up the difference.

    This is a planning setting, not a statement of HR or pay policy.  Different
    organisations treat sickness and long-term absence differently, so the
    behaviour is configured per leave type rather than assumed.
    """

    code: str
    label: str
    colour: str = "FFFF00"
    font_colour: str = "000000"
    counts_as_absence: bool = True
    credits_hours: bool = True
    credited_method: str = CREDIT_FROM_PATTERN
    fixed_daily_hours: float = 0.0


# --------------------------------------------------------------------------
# settings
# --------------------------------------------------------------------------

@dataclass
class RosterDetails:
    """Who the roster is for.  Organisation-neutral by default."""

    rota_name: str = "Laboratory Staff Rota"
    organisation: str = ""
    department: str = ""
    site: str = ""
    prepared_by: str = ""

    @property
    def heading(self) -> str:
        parts = [part for part in (self.organisation, self.department, self.site)
                 if part]
        return f"{self.rota_name}" + (f" — {' · '.join(parts)}" if parts else "")


@dataclass
class Rules:
    """Organisational rules, all configurable.

    These are the laboratory's own operating rules.  They are not a
    determination of legal or regulatory compliance, and the wording used
    throughout the application reflects that.
    """

    # Seniority is recorded explicitly against each member of staff, on the Staff
    # sheet.  There is deliberately no band threshold here: a band is about pay
    # and responsibility, and inferring "senior" from it silently overrides what a
    # manager actually recorded.  Shifts that need a particular grade use the
    # separate Min Band requirement instead.
    minimum_rest_hours: float = 11.0
    max_consecutive_days: int = 6
    max_consecutive_nights: int = 4
    night_block_length: int = 3
    recovery_days_after_nights: int = 2
    hours_tolerance_percent: float = 10.0
    expiry_warning_days: list[int] = field(default_factory=lambda: [30, 60, 90])
    max_simultaneous_bench_assignments: int = 1
    cross_cover_allowed: bool = False
    weekend_days: set[int] = field(default_factory=lambda: {5, 6})
    seed: int = 42
    rotation_warning_days: int = 56

    @property
    def hours_tolerance_fraction(self) -> float:
        return max(0.0, self.hours_tolerance_percent) / 100.0


@dataclass
class Period:
    start: date
    end: date

    @property
    def days(self) -> list[date]:
        span = (self.end - self.start).days
        return [self.start + timedelta(days=offset) for offset in range(span + 1)]

    @property
    def day_count(self) -> int:
        return (self.end - self.start).days + 1

    @property
    def weeks(self) -> float:
        return self.day_count / 7.0

    def contains(self, day: date) -> bool:
        return self.start <= day <= self.end


@dataclass
class Config:
    """Everything read from one workbook."""

    details: RosterDetails
    period: Period
    rules: Rules
    staff: list[Staff]
    competencies: list[Competency]
    shifts: list[ShiftType]
    requirements: list[ShiftRequirement]
    benches: list[Bench]
    leave: list[LeaveEntry]
    leave_types: dict[str, LeaveType]
    schema_version: int = 2

    # -- lookups --------------------------------------------------------

    def staff_by_id(self) -> dict[str, Staff]:
        return {person.staff_id: person for person in self.staff}

    def shift_by_code(self) -> dict[str, ShiftType]:
        return {shift.code: shift for shift in self.shifts}

    def competencies_for(self, staff_id: str) -> list[Competency]:
        return [record for record in self.competencies
                if record.staff_id == staff_id]

    def competency(self, staff_id: str, discipline: str) -> Competency | None:
        """The best record a person holds in a discipline.

        Somebody may have several records in one discipline; the strongest one
        governs, so a current 'Competent' is not hidden by an old 'In Training'.
        """
        ranking = {status: index for index, status
                   in enumerate([CompetencyStatus.NOT_TRAINED,
                                 CompetencyStatus.EXPIRED,
                                 CompetencyStatus.IN_TRAINING,
                                 CompetencyStatus.SUPERVISED,
                                 CompetencyStatus.COMPETENT,
                                 CompetencyStatus.TRAINER,
                                 CompetencyStatus.ASSESSOR])}
        matches = [record for record in self.competencies
                   if record.staff_id == staff_id
                   and record.discipline.upper() == discipline.upper()]
        if not matches:
            return None
        return max(matches, key=lambda r: (
            ranking.get(r.effective_status(self.period.start), 0),
            r.expiry_date or date.max,
        ))

    def is_independently_competent(self, staff_id: str, discipline: str,
                                  as_of: date) -> bool:
        record = self.competency(staff_id, discipline)
        return bool(record and record.is_independent(as_of))

    def can_authorise(self, staff_id: str, discipline: str, as_of: date) -> bool:
        record = self.competency(staff_id, discipline)
        return bool(record and record.can_authorise(as_of))

    def disciplines(self) -> list[str]:
        """Every discipline mentioned by a bench or a requirement."""
        found = {bench.discipline.upper() for bench in self.benches
                 if bench.discipline}
        for requirement in self.requirements:
            found.update(key.upper() for key in requirement.required_competencies)
            found.update(key.upper() for key in requirement.required_authorisers)
        found.update(record.discipline.upper() for record in self.competencies
                     if record.discipline)
        return sorted(found)

    def requirement_for(self, shift: ShiftType, day: date) -> ShiftRequirement:
        """The requirement that applies to a shift on a day.

        A day-specific requirement wins over a general one, so a manager can set
        a different Saturday minimum without restating everything.
        """
        matches = [requirement for requirement in self.requirements
                   if requirement.shift_code == shift.code
                   and requirement.applies_on(day, self.rules.weekend_days)]
        if not matches:
            return ShiftRequirement(shift_code=shift.code, min_staff=0)
        specific = [r for r in matches
                    if re.sub(r"[^a-z]", "", r.days.lower()) not in ("", "all")]
        return (specific or matches)[0]

    def leave_type(self, code: str) -> LeaveType:
        key = re.sub(r"[^a-z0-9]", "", str(code or "").lower())
        return self.leave_types.get(key) or LeaveType(code=code, label=code)

    def credited_hours_for(self, entry: LeaveEntry, person: Staff) -> float:
        """Hours to credit for one absence, clipped to the roster period.

        An explicit figure on the leave row wins.  Otherwise the hours come from
        the person's own working pattern, so a part-timer absent for a week is
        credited a part-time week rather than a full one.
        """
        kind = self.leave_type(entry.code)
        if not kind.credits_hours or kind.credited_method == CREDIT_NONE:
            return 0.0

        first = max(entry.start, self.period.start)
        last = min(entry.end, self.period.end)
        if last < first:
            return 0.0

        if entry.credited_hours is not None:
            # Pro-rata if only part of the absence falls inside the period.
            whole = (entry.end - entry.start).days + 1
            inside = (last - first).days + 1
            return round(entry.credited_hours * inside / whole, 4)

        per_day = (kind.fixed_daily_hours if kind.credited_method == CREDIT_FIXED
                   and kind.fixed_daily_hours else person.normal_daily_hours)

        total = 0.0
        day = first
        while day <= last:
            # Only days the person would otherwise have worked are credited.
            if person.availability.works_weekday(day, self.period.start) \
                    and day.weekday() not in self.rules.weekend_days:
                total += per_day
            day += timedelta(days=1)
        return round(total, 4)

    def target_hours(self, person: Staff) -> float:
        """Contracted hours for the roster period.

        Weekly contracted hours scaled by the length of the period.  If only an
        FTE is given, it is applied to the standard full-time week.
        """
        weekly = person.contracted_weekly_hours
        if not weekly:
            weekly = 37.5 * (person.fte or 1.0)
        return round(weekly * self.period.weeks, 2)
