"""Turning a finished roster into things a laboratory manager can act on.

The question this module answers is the one a manager actually asks:

    "Do I have the right people, with the right competencies, in the right
    laboratory areas for this shift?"

Everything is expressed as an *issue* with a severity, so the output is a review
list rather than a set of scheduler statistics.  Nothing here decides anything;
it reports what the draft roster does and does not achieve.
"""

from __future__ import annotations

import statistics
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta

from .models import CompetencyStatus, Config
from .scheduler import Scheduler, Shortfall

CRITICAL = "CRITICAL"
REVIEW = "REVIEW"
PASSED = "PASSED"

#: Fairness verdicts.  Deliberately coarse: there is no single correct fairness
#: formula, so the tool reports a judgement to review rather than a score.
GOOD = "Good"
UNEVEN = "Uneven"


@dataclass
class Issue:
    severity: str
    category: str
    title: str
    explanation: str
    review_point: str = ""
    day: date | None = None
    shift_code: str = ""
    bench_name: str = ""
    staff: list[str] = field(default_factory=list)
    required: str = ""
    available: str = ""
    impact: str = ""
    causes: list[str] = field(default_factory=list)
    """The underlying checks that produced this issue.  Managers see one
    actionable problem; traceability back to the individual rules is kept here."""

    @property
    def when(self) -> str:
        return self.day.strftime("%a %d %b %Y") if self.day else ""


@dataclass
class HoursRow:
    staff_id: str
    name: str
    band: str
    contracted_weekly_hours: float
    fte: float
    target_hours: float
    worked_hours: float
    credited_absence_hours: float
    total_accounted_hours: float
    variance: float
    percent_of_target: float | None
    shifts: int
    nights: int
    saturdays: int
    sundays: int
    full_weekends: int
    lates: int
    earlies: int
    leave_days: int
    status: str                            # Within tolerance / Under / Over


@dataclass
class Resilience:
    """How exposed the laboratory is if one person is unavailable."""

    discipline: str
    competent_count: int
    authoriser_count: int
    severity: str
    message: str


class Analysis:
    """Reads a built :class:`~labroster.scheduler.Scheduler` and reports on it."""

    def __init__(self, scheduler: Scheduler):
        self.scheduler = scheduler
        self.config: Config = scheduler.config
        self.rules = scheduler.rules
        self.period = scheduler.period
        self.by_id = scheduler.by_id

        self.gaps: list[Shortfall] = []
        self.issues: list[Issue] = []
        self.hours_rows: list[HoursRow] = []
        self.workforce_resilience: list[Resilience] = []
        self.expiring: list[dict] = []
        self.recency: list[dict] = []
        self.metrics: dict = {}

        self.run()

    # ------------------------------------------------------------------

    def run(self) -> None:
        self.gaps = self.final_gaps()
        self.check_shift_coverage()
        self.check_bench_coverage()
        self.check_double_bench_allocation()
        self.check_unavailable_assignments()
        self.check_rest()
        self.build_hours_rows()
        self.check_hours_variance()
        self.check_competency_expiry()
        self.check_workforce_resilience()
        self.check_shift_resilience()
        self.build_recency()
        self.build_metrics()

    def add(self, severity, category, title, explanation, **kwargs) -> None:
        self.issues.append(Issue(severity=severity, category=category,
                                 title=title, explanation=explanation, **kwargs))

    def name_of(self, staff_id: str) -> str:
        person = self.by_id.get(staff_id)
        return person.name if person else staff_id

    # -- coverage --------------------------------------------------------

    def _expected_slots(self) -> tuple[int, int]:
        """Required and filled staffing slots across the whole period.

        This is a headcount measure only: it says whether positions are occupied,
        not whether the shift is adequately staffed.  See
        :meth:`shift_requirement_compliance` for that.
        """
        required = filled = 0
        for day in self.period.days:
            for shift in self.config.shifts:
                if not shift.applies_on(day, self.rules.weekend_days):
                    continue
                requirement = self.config.requirement_for(shift, day)
                required += requirement.min_staff
                filled += min(len(self.scheduler.assigned_to(day, shift)),
                              requirement.min_staff)
        return required, filled

    def shift_requirement_compliance(self) -> tuple[int, int, Counter]:
        """How many shift instances satisfy *every* configured condition.

        A shift with all its positions occupied is not necessarily a shift that can
        safely run, so this checks the whole set: minimum staffing, registration,
        seniority, band, coordinator, competencies, authorisers, the trainee limit,
        section coverage, availability and the configured rest interval.

        Returns the number of shift instances, how many met everything, and a count
        of why the others did not.
        """
        scheduler = self.scheduler
        as_of = self.period.start
        total = met = 0
        reasons: Counter = Counter()

        rest_by_day: dict[date, set[str]] = defaultdict(set)
        for conflict in scheduler.rest_conflicts:
            rest_by_day[conflict.day].add(conflict.staff_id)

        for day in self.period.days:
            for shift in self.config.shifts:
                if not shift.applies_on(day, self.rules.weekend_days):
                    continue
                requirement = self.config.requirement_for(shift, day)
                if not requirement.min_staff and not \
                        scheduler.competency_demand(day, shift, requirement):
                    continue                      # nothing configured for it
                total += 1
                failed: set[str] = set()

                assigned = scheduler.assigned_to(day, shift)
                if len(assigned) < requirement.min_staff:
                    failed.add("Minimum staffing")

                for need in scheduler.outstanding_needs(day, shift, requirement):
                    category, _ = self.KIND_LABELS.get(
                        need["kind"], ("Other requirement", ""))
                    failed.add(category)

                if requirement.max_trainees:
                    trainees = sum(1 for sid in assigned
                                   if self.by_id[sid].trainee)
                    if trainees > requirement.max_trainees:
                        failed.add("Trainee limit")

                for bench in scheduler.benches_for(day, shift):
                    wanted = bench.required_on(day, self.rules.weekend_days)
                    if wanted and len(scheduler.bench_staff(
                            day, bench.name, shift.code)) < wanted:
                        failed.add("Section coverage")

                for staff_id in assigned:
                    person = self.by_id.get(staff_id)
                    if person is None:
                        continue
                    if scheduler.on_leave(staff_id, day) or not \
                            person.availability.works_weekday(day, as_of):
                        failed.add("Availability")
                    if staff_id in rest_by_day.get(day, ()):
                        failed.add("Rest rules")

                if failed:
                    for reason in failed:
                        reasons[reason] += 1
                else:
                    met += 1

        return total, met, reasons

    #: How a technical shortfall is described to a manager.
    KIND_LABELS = {
        "staffing": ("Staffing level", "staff"),
        "senior": ("Senior cover", "senior member of staff"),
        "registered": ("Registration", "registered biomedical scientist"),
        "coordinator": ("Shift coordination", "shift coordinator"),
        "competency": ("Competency coverage", "independently competent scientist"),
        "authoriser": ("Authorisation", "result authoriser"),
        "trainer": ("Training", "trainer or supervisor"),
        "bench": ("Section coverage", "competent scientist for the section"),
    }

    def discipline_name(self, code: str) -> str:
        """How a discipline is written for a manager.

        Disciplines are shown by their short code — BT, HAEM, COAG, MORPH — because
        that is the everyday shorthand in a diagnostic laboratory, and it keeps the
        workbook and the interface using the same terms. Section names remain in
        words, so an issue reads "Morphology cannot be covered" while the
        requirement behind it reads "1 independently competent MORPH scientist".

        This is the single place that decision is made, so it can be changed to
        expand codes into full names without touching anything else.
        """
        return (code or "").upper()

    def final_gaps(self) -> list[Shortfall]:
        """Coverage gaps in the *finished* roster.

        Recomputed rather than taken from the shortfalls the scheduler recorded
        while building, because those are a construction log: a gap noted part-way
        through filling a shift may be resolved by the time the shift is complete.
        Reporting the log would show a manager a critical problem that no longer
        exists, and would disagree with the compliance figure, which is measured
        from the finished roster.
        """
        scheduler = self.scheduler
        gaps: list[Shortfall] = []

        for day in self.period.days:
            for shift in self.config.shifts:
                if not shift.applies_on(day, self.rules.weekend_days):
                    continue
                requirement = self.config.requirement_for(shift, day)
                assigned = scheduler.assigned_to(day, shift)

                if requirement.min_staff and len(assigned) < requirement.min_staff:
                    gaps.append(Shortfall(
                        day=day, shift_code=shift.code, kind="staffing",
                        detail=(f"{len(assigned)} of {requirement.min_staff} "
                                f"staff available"),
                        needed=requirement.min_staff, found=len(assigned)))

                for need in scheduler.outstanding_needs(day, shift, requirement):
                    gaps.append(Shortfall(
                        day=day, shift_code=shift.code, kind=need["kind"],
                        detail=f"no available {need['label']}",
                        needed=need["missing"], found=0,
                        discipline=need.get("discipline", "")))

                for bench in scheduler.benches_for(day, shift):
                    wanted = bench.required_on(day, self.rules.weekend_days)
                    allocated = scheduler.bench_staff(day, bench.name, shift.code)
                    if wanted and len(allocated) < wanted:
                        gaps.append(Shortfall(
                            day=day, shift_code=shift.code, kind="bench",
                            detail=(f"{bench.name}: {len(allocated)} of {wanted} "
                                    f"independently competent staff allocated"),
                            needed=wanted, found=len(allocated),
                            discipline=bench.discipline, bench_name=bench.name))
        return gaps

    def check_shift_coverage(self) -> None:
        """Report coverage gaps as one problem per shift, per cause.

        A missing morphology scientist previously produced two critical warnings —
        one for the competency and one for the bench that consequently could not be
        staffed. They are the same problem, so they are now consolidated into a
        single actionable issue, with the underlying checks kept for traceability.
        """
        grouped: dict[tuple, list] = defaultdict(list)
        for shortfall in self.gaps:
            key = (shortfall.day, shortfall.shift_code,
                   (shortfall.discipline or "").upper())
            grouped[key].append(shortfall)

        shift_by_code = self.config.shift_by_code()
        seen_kinds: set[str] = set()

        for (day, shift_code, discipline), group in sorted(
                grouped.items(), key=lambda item: (item[0][0] or date.min,
                                                   item[0][1], item[0][2])):
            shift = shift_by_code.get(shift_code)
            shift_label = shift.name if shift else shift_code
            kinds = {item.kind for item in group}
            seen_kinds |= kinds
            benches = sorted({item.bench_name for item in group if item.bench_name})

            # The most specific cause leads: a section that cannot be staffed is
            # usually the consequence of a competency that nobody available holds.
            lead = next((kind for kind in
                         ("competency", "bench", "authoriser", "senior",
                          "registered", "coordinator", "trainer", "staffing")
                         if kind in kinds), "staffing")
            category, role = self.KIND_LABELS.get(lead, ("Coverage", "staff"))

            needed = max((item.needed for item in group), default=1)
            found = min((item.found for item in group), default=0)
            readable = self.discipline_name(discipline)

            if lead in ("competency", "bench"):
                # Name the section in the title where the workbook defines one,
                # even if only the competency check failed: "Morphology cannot be
                # covered" reads better than "MORPH cannot be covered", while the
                # requirement line keeps the code the laboratory actually uses.
                section = next((bench.name for bench in self.config.benches
                                if bench.discipline.upper() == discipline), "")
                subject = benches[0] if benches else (section or readable
                                                      or "The section")
                title = f"{subject} cannot be covered"
                required = (f"{needed} independently competent "
                            f"{readable} scientist"
                            f"{'s' if needed > 1 else ''}")
                impact = (f"{subject} cannot be staffed on this shift."
                          if benches else
                          f"The {readable} requirement for this "
                          f"shift is not met.")
                review = ("Consider moving a competent scientist onto this shift, "
                          "changing the deployment plan, or arranging additional "
                          "cover.")
            elif lead == "staffing":
                title = f"{shift_label} is short of staff"
                required = f"{needed} staff"
                impact = "The shift is running below its minimum staffing level."
                review = ("Find cover, use bank or agency, or reduce the service "
                          "planned for this shift.")
            else:
                title = f"No {role} on {shift_label}"
                required = (f"{needed} {role}"
                            + (f" ({readable})" if readable else ""))
                impact = f"This shift's {category.lower()} requirement is not met."
                review = ("Check whether somebody suitable can be moved onto this "
                          "shift, or whether the requirement should be met "
                          "differently.")

            detail = "; ".join(sorted(item.detail for item in group))
            # Required and Available are surfaced as their own fields, so the
            # explanation states the consequence rather than repeating them.
            self.add(CRITICAL, category, title,
                     f"{shift_label}. {impact}",
                     review_point=review, day=day, shift_code=shift_code,
                     bench_name=benches[0] if benches else "",
                     required=required, available=str(found), impact=impact,
                     causes=sorted(kinds) + ([detail] if detail else []))

        if "staffing" not in seen_kinds:
            self.add(PASSED, "Staffing level", "Minimum staffing met",
                     "Every shift in the period reached its minimum staffing level.")
        if self.config.benches and not ({"bench", "competency"} & seen_kinds):
            self.add(PASSED, "Section coverage",
                     "Required section competencies covered",
                     "Every section had the required number of independently "
                     "competent staff allocated to it.")

    def check_bench_coverage(self) -> None:
        """Section gaps are reported by :meth:`check_shift_coverage`.

        Kept as a separate step so the ordering of the issue list stays stable and
        so a future part-shift allocation model has somewhere obvious to report.
        """
        return

    def check_double_bench_allocation(self) -> None:
        """Belt and braces: prove nobody holds two benches at once.

        The allocator prevents this by construction, but a manual override could
        reintroduce it, so it is verified against the finished allocations.
        """
        limit = max(1, self.rules.max_simultaneous_bench_assignments)
        offenders: list[Issue] = []
        grouped: dict[tuple[date, str], Counter] = defaultdict(Counter)
        for allocation in self.scheduler.bench_allocations:
            grouped[(allocation.day, allocation.shift_code)][allocation.staff_id] += 1

        for (day, shift_code), counts in sorted(grouped.items()):
            for staff_id, count in counts.items():
                if count > limit:
                    offenders.append(Issue(
                        severity=CRITICAL, category="Bench coverage",
                        title="One person counted on more than one bench",
                        explanation=(
                            f"{self.name_of(staff_id)} is allocated to {count} "
                            f"benches at the same time. One person cannot provide "
                            f"simultaneous cover for more than {limit}."),
                        review_point="Allocate separate staff to each bench.",
                        day=day, shift_code=shift_code, staff=[staff_id]))
        self.issues.extend(offenders)
        if not offenders:
            self.add(PASSED, "Bench coverage",
                     "No bench counted twice",
                     f"No member of staff is relied on for more than {limit} "
                     f"bench at the same time.")

    def check_unavailable_assignments(self) -> None:
        """Nobody should be rostered while on leave or outside their pattern."""
        problems = 0
        for (day, staff_id), assignment in sorted(self.scheduler.assignments.items()):
            person = self.by_id.get(staff_id)
            if person is None:
                continue
            if self.scheduler.on_leave(staff_id, day):
                problems += 1
                self.add(CRITICAL, "Availability",
                         f"{person.name} is rostered while absent",
                         f"{person.name} is recorded as "
                         f"{self.scheduler.leave[(day, staff_id)]} on this date but "
                         f"appears on the roster.",
                         review_point="Remove the assignment or correct the leave "
                                      "record.",
                         day=day, shift_code=assignment.shift_code, staff=[staff_id])
            elif not person.availability.works_weekday(day, self.period.start):
                problems += 1
                self.add(CRITICAL, "Availability",
                         f"{person.name} is rostered outside their working pattern",
                         f"{person.name} does not normally work "
                         f"{day.strftime('%A')}s.",
                         review_point="Confirm the change with the member of staff "
                                      "or move the shift.",
                         day=day, shift_code=assignment.shift_code, staff=[staff_id])
        if not problems:
            self.add(PASSED, "Availability", "Working patterns respected",
                     "Nobody is rostered while on leave or outside their agreed "
                     "working pattern.")

    def check_rest(self) -> None:
        for conflict in self.scheduler.rest_conflicts:
            self.add(CRITICAL, "Rest", "Configured rest rule conflict",
                     f"{self.name_of(conflict.staff_id)} has "
                     f"{conflict.rest_interval_hours:g} hours between finishing "
                     f"{conflict.previous_shift} and starting {conflict.next_shift}. "
                     f"The rule configured for this laboratory is "
                     f"{conflict.minimum_rest_hours:g} hours.",
                     review_point="Adjust one of the two shifts. This is the "
                                  "laboratory's own configured rule, not a "
                                  "determination of legal compliance.",
                     day=conflict.day, staff=[conflict.staff_id])
        if not self.scheduler.rest_conflicts:
            self.add(PASSED, "Rest", "Rest rules satisfied",
                     f"Every consecutive pair of shifts leaves at least "
                     f"{self.rules.minimum_rest_hours:g} hours, the interval "
                     f"configured for this laboratory.")

    # -- hours -----------------------------------------------------------

    def build_hours_rows(self) -> None:
        scheduler = self.scheduler
        tolerance = self.rules.hours_tolerance_fraction
        night_codes = {s.code for s in self.config.shifts if s.is_night}
        late_codes = {s.code for s in self.config.shifts
                      if "late" in s.name.lower()}
        early_codes = {s.code for s in self.config.shifts
                       if "early" in s.name.lower()}

        for person in self.config.staff:
            shifts = sum(1 for (_, sid) in scheduler.assignments
                         if sid == person.staff_id)
            leave_days = sum(1 for (_, sid) in scheduler.leave
                             if sid == person.staff_id)
            target = person.target_period_hours
            accounted = person.total_accounted_hours
            status = "Within tolerance"
            if target:
                if accounted < target * (1 - tolerance):
                    status = "Under target"
                elif accounted > target * (1 + tolerance):
                    status = "Over target"
            elif accounted:
                status = "No target set"

            self.hours_rows.append(HoursRow(
                staff_id=person.staff_id, name=person.name, band=person.band,
                contracted_weekly_hours=person.contracted_weekly_hours,
                fte=person.fte, target_hours=round(target, 2),
                worked_hours=round(person.allocated_hours, 2),
                credited_absence_hours=round(person.credited_absence_hours, 2),
                total_accounted_hours=round(accounted, 2),
                variance=person.hours_variance,
                percent_of_target=person.percent_of_target,
                shifts=shifts,
                nights=sum(scheduler.count_shift_code(person.staff_id, c)
                           for c in night_codes),
                saturdays=scheduler.count_saturdays(person.staff_id),
                sundays=scheduler.count_sundays(person.staff_id),
                full_weekends=scheduler.count_full_weekends(person.staff_id),
                lates=sum(scheduler.count_shift_code(person.staff_id, c)
                          for c in late_codes),
                earlies=sum(scheduler.count_shift_code(person.staff_id, c)
                            for c in early_codes),
                leave_days=leave_days, status=status))

    def check_hours_variance(self) -> None:
        off_target = [row for row in self.hours_rows
                      if row.status in ("Under target", "Over target")
                      and row.target_hours]
        for row in off_target:
            direction = "below" if row.variance < 0 else "above"
            credited = (f", plus {row.credited_absence_hours:g} hours credited "
                        f"for absence" if row.credited_absence_hours else "")
            self.add(REVIEW, "Contracted hours",
                     f"{row.name} is {direction} contracted hours",
                     f"{row.name} works {row.worked_hours:g} hours{credited}, "
                     f"giving {row.total_accounted_hours:g} accounted against a "
                     f"target of {row.target_hours:g} "
                     f"({row.percent_of_target:g}% of target, "
                     f"{row.variance:+g} hours).",
                     review_point="Check whether this reflects part-time hours, "
                                  "absence in the period, or a roster imbalance.",
                     staff=[row.staff_id])
        if not off_target:
            self.add(PASSED, "Contracted hours", "Hours within tolerance",
                     f"Everybody is within "
                     f"{self.rules.hours_tolerance_percent:g}% of their "
                     f"contracted hours for the period.")

    # -- competency ------------------------------------------------------

    def check_competency_expiry(self) -> None:
        thresholds = sorted(self.rules.expiry_warning_days) or [30, 60, 90]
        widest = thresholds[-1]
        as_of = self.period.start

        for record in self.config.competencies:
            person = self.by_id.get(record.staff_id)
            if person is None:
                continue
            if record.has_expired(as_of):
                self.expiring.append({
                    "staff_id": record.staff_id, "name": person.name,
                    "discipline": record.discipline, "name_of_competency": record.name,
                    "expiry": record.expiry_date, "days": record.days_until_expiry(as_of),
                    "state": "Expired"})
                self.add(CRITICAL, "Competency",
                         f"{person.name}'s {record.discipline} competency has expired",
                         f"The record expired on "
                         f"{record.expiry_date.strftime('%d %b %Y') if record.expiry_date else 'an unrecorded date'}. "
                         f"Expired competencies are not counted as coverage.",
                         review_point="Reassess and record the outcome, or remove "
                                      "the person from work requiring it.",
                         staff=[record.staff_id])
                continue

            days = record.days_until_expiry(as_of)
            if days is not None and 0 <= days <= widest:
                band = next(t for t in thresholds if days <= t)
                self.expiring.append({
                    "staff_id": record.staff_id, "name": person.name,
                    "discipline": record.discipline, "name_of_competency": record.name,
                    "expiry": record.expiry_date, "days": days,
                    "state": f"Within {band} days"})
                self.add(REVIEW, "Competency",
                         f"{person.name}'s {record.discipline} competency expires soon",
                         f"It expires on {record.expiry_date.strftime('%d %b %Y')}, "
                         f"in {days} days.",
                         review_point="Schedule reassessment before the expiry date.",
                         staff=[record.staff_id])

        if not self.expiring:
            self.add(PASSED, "Competency", "No competencies expiring soon",
                     f"No competency expires within {widest} days of the start of "
                     f"the period.")

    def check_workforce_resilience(self) -> None:
        """How many people in the whole workforce can do each thing.

        This is a workforce risk, distinct from whether a particular shift is
        thin: one morphology-competent scientist in the department is a different
        problem from one rostered on Tuesday.
        """
        as_of = self.period.start
        for discipline in self.config.disciplines():
            competent = [person.staff_id for person in self.config.staff
                         if self.config.is_independently_competent(
                             person.staff_id, discipline, as_of)]
            authorisers = [person.staff_id for person in self.config.staff
                           if self.config.can_authorise(
                               person.staff_id, discipline, as_of)]
            count = len(competent)
            if count <= 1:
                severity, message = CRITICAL, (
                    f"Only {count} member of staff in the workforce is "
                    f"independently competent in {discipline}. If they are absent, "
                    f"the laboratory has no cover at all.")
            elif count == 2:
                severity, message = REVIEW, (
                    f"Two members of staff are independently competent in "
                    f"{discipline}. Cover depends on both remaining available.")
            else:
                severity, message = PASSED, (
                    f"{count} members of staff are independently competent in "
                    f"{discipline}.")

            self.workforce_resilience.append(Resilience(
                discipline=discipline, competent_count=count,
                authoriser_count=len(authorisers), severity=severity,
                message=message))
            self.add(severity, "Workforce resilience",
                     f"{discipline}: {count} independently competent",
                     message,
                     review_point=("Consider training additional staff in this "
                                   "discipline." if severity != PASSED else ""),
                     staff=competent)

    def check_shift_resilience(self) -> None:
        """Whether individual shifts depend on a single person.

        Reported one line per shift and discipline, with a count of how many
        dates are affected, rather than one line per day.  A thinly covered
        section is one problem to think about, not thirty.
        """
        as_of = self.period.start
        thin: dict[tuple[str, str, str], list] = defaultdict(list)

        for day in self.period.days:
            for shift in self.config.shifts:
                if not shift.applies_on(day, self.rules.weekend_days):
                    continue
                on_duty = self.scheduler.assigned_to(day, shift)
                if not on_duty:
                    continue
                requirement = self.config.requirement_for(shift, day)

                for discipline in self.scheduler.competency_demand(
                        day, shift, requirement):
                    competent = [sid for sid in on_duty
                                 if self.config.is_independently_competent(
                                     sid, discipline, as_of)]
                    if len(competent) == 1:
                        thin[(shift.code, discipline, "competent")].append(
                            (day, competent[0]))

                for discipline in requirement.required_authorisers:
                    authorisers = [sid for sid in on_duty
                                   if self.config.can_authorise(
                                       sid, discipline.upper(), as_of)]
                    if len(authorisers) == 1:
                        thin[(shift.code, discipline.upper(),
                              "authoriser")].append((day, authorisers[0]))

        shift_names = {shift.code: shift.name for shift in self.config.shifts}
        for (shift_code, discipline, kind), occurrences in sorted(thin.items()):
            people = sorted({staff_id for _, staff_id in occurrences})
            dates = [day for day, _ in occurrences]
            role = ("independently competent member of staff"
                    if kind == "competent" else "result authoriser")
            named = ", ".join(self.name_of(sid) for sid in people[:4])
            if len(people) > 4:
                named += f" and {len(people) - 4} others"
            self.add(REVIEW, "Shift resilience",
                     f"{shift_names.get(shift_code, shift_code)}: single "
                     f"{discipline} {role.split()[0]} cover on "
                     f"{len(occurrences)} date"
                     f"{'s' if len(occurrences) > 1 else ''}",
                     f"On {len(occurrences)} of these shifts there is only one "
                     f"{discipline} {role} on duty ({named}). Short-notice absence "
                     f"on those days would leave the section uncovered. First "
                     f"affected: {min(dates):%d %b}, last: {max(dates):%d %b}.",
                     review_point="Consider rostering a second competent person, "
                                  "or accept the risk knowingly.",
                     shift_code=shift_code, staff=people)

    # -- recency ---------------------------------------------------------

    def build_recency(self) -> None:
        """Groundwork for rotation warnings: when did somebody last work a section?

        No rotation rule is invented here.  The interval is configurable and the
        report simply states how long it has been.
        """
        as_of = self.period.end
        last_seen: dict[tuple[str, str], date] = {}
        for allocation in self.scheduler.bench_allocations:
            key = (allocation.staff_id, allocation.discipline.upper())
            if key not in last_seen or allocation.day > last_seen[key]:
                last_seen[key] = allocation.day

        threshold = self.rules.rotation_warning_days
        for person in self.config.staff:
            for discipline in self.config.disciplines():
                if not self.config.is_independently_competent(
                        person.staff_id, discipline, self.period.start):
                    continue
                worked = last_seen.get((person.staff_id, discipline))
                days_since = (as_of - worked).days if worked else None
                self.recency.append({
                    "staff_id": person.staff_id, "name": person.name,
                    "discipline": discipline,
                    "last_worked_date": worked,
                    "days_since_last_assignment": days_since,
                    "target_rotation_interval": threshold,
                })
                if worked is None and threshold:
                    self.add(REVIEW, "Section rotation",
                             f"{person.name} did not work in {discipline}",
                             f"{person.name} is competent in {discipline} but was "
                             f"not allocated to it during this period.",
                             review_point=f"Rotation interval is configured at "
                                          f"{threshold} days. Adjust if this is "
                                          f"not a concern.",
                             staff=[person.staff_id])

    # -- fairness --------------------------------------------------------

    def fairness(self, values: list[int]) -> str:
        """A coarse verdict on how evenly something is shared.

        Compares only the group handed in, which is how peer comparison is kept
        honest: night counts are compared between people who can work nights.
        There is no single correct fairness formula, so this reports a judgement
        for a manager to review rather than a score to optimise.
        """
        useful = [value for value in values]
        if len(useful) < 2:
            return GOOD
        spread = max(useful) - min(useful)
        if spread <= 1:
            return GOOD
        mean = statistics.mean(useful)
        if mean and spread <= max(2.0, mean * 0.5):
            return REVIEW.title()
        return UNEVEN

    def weekend_fairness(self) -> str:
        eligible = [person for person in self.config.staff if person.weekends_ok]
        return self.fairness([self.scheduler.count_weekend_days(p.staff_id)
                              for p in eligible])

    def night_fairness(self) -> str:
        """Judged on blocks of nights, which is the unit nights are rostered in.

        Counting individual nights would report a level share as uneven whenever
        the number of blocks does not divide by the number of eligible staff.
        """
        eligible = [person for person in self.config.staff if person.nights_ok]
        return self.fairness([self.scheduler.count_night_blocks(p.staff_id)
                              for p in eligible])

    # -- dashboard -------------------------------------------------------

    def build_metrics(self) -> None:
        required, filled = self._expected_slots()
        coverage = 100.0 if not required else round(100.0 * filled / required, 1)

        shifts_total, shifts_met, shift_reasons = \
            self.shift_requirement_compliance()
        compliance = (100.0 if not shifts_total
                      else round(100.0 * shifts_met / shifts_total, 1))

        critical = [issue for issue in self.issues if issue.severity == CRITICAL]
        review = [issue for issue in self.issues if issue.severity == REVIEW]

        # What is actually driving the problems, so a manager can start somewhere.
        causes = Counter(issue.category for issue in critical + review)

        if critical:
            status = "ATTENTION REQUIRED"
        elif review:
            status = "REVIEW SUGGESTED"
        else:
            status = "NO ISSUES FOUND"

        weekend_verdict = self.weekend_fairness()
        night_verdict = self.night_fairness()

        self.metrics = {
            "roster_status": status,
            # Headcount only: are the positions occupied?
            "staffing_slot_coverage_percent": coverage,
            # The meaningful one: does the shift satisfy everything configured?
            "shifts_meeting_all_requirements_percent": compliance,
            "shift_instances": shifts_total,
            "shift_instances_met": shifts_met,
            "requirement_failure_causes": dict(shift_reasons.most_common()),
            "main_causes": dict(causes.most_common()),
            "required_slots": required,
            "filled_slots": filled,
            "unfilled_shifts": len([g for g in self.gaps
                                    if g.kind == "staffing"]),
            "uncovered_benches": len([g for g in self.gaps
                                      if g.kind == "bench"]),
            "senior_cover_gaps": len([g for g in self.gaps
                                      if g.kind == "senior"]),
            "competency_gaps": len([g for g in self.gaps
                                    if g.kind in ("competency", "authoriser")]),
            "rest_conflicts": len(self.scheduler.rest_conflicts),
            "staff_outside_target_hours": len(
                [row for row in self.hours_rows
                 if row.status in ("Under target", "Over target")
                 and row.target_hours]),
            "competencies_expiring_soon": len(
                [item for item in self.expiring if item["state"] != "Expired"]),
            "competencies_expired": len(
                [item for item in self.expiring if item["state"] == "Expired"]),
            "single_points_of_failure": len(
                [item for item in self.workforce_resilience
                 if item.severity == CRITICAL]),
            "weekend_fairness": weekend_verdict,
            "night_fairness": night_verdict,
            "critical_count": len(critical),
            "review_count": len(review),
            "passed_count": len([i for i in self.issues if i.severity == PASSED]),
            "total_assignments": len(self.scheduler.assignments),
            "staff_count": len(self.config.staff),
            "day_count": self.period.day_count,
        }

    # -- serialisation for the browser -----------------------------------

    def issues_payload(self) -> list[dict]:
        return [{
            "severity": issue.severity,
            "category": issue.category,
            "title": issue.title,
            "explanation": issue.explanation,
            "review_point": issue.review_point,
            "date": issue.day.isoformat() if issue.day else None,
            "when": issue.when,
            "shift": issue.shift_code,
            "bench": issue.bench_name,
            "staff": [self.name_of(sid) for sid in issue.staff],
            "required": issue.required,
            "available": issue.available,
            "impact": issue.impact,
            "causes": issue.causes,
        } for issue in self.issues]

    def hours_payload(self) -> list[dict]:
        return [{
            "name": row.name, "band": row.band,
            "contracted_weekly_hours": row.contracted_weekly_hours,
            "fte": row.fte, "target": row.target_hours,
            "worked": row.worked_hours,
            "credited": row.credited_absence_hours,
            "accounted": row.total_accounted_hours,
            "variance": row.variance,
            "percent": row.percent_of_target, "status": row.status,
            "shifts": row.shifts, "nights": row.nights,
            "saturdays": row.saturdays, "sundays": row.sundays,
            "full_weekends": row.full_weekends, "leave_days": row.leave_days,
        } for row in self.hours_rows]
