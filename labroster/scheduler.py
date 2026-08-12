"""Deciding who works when.

The engine is rules-based, not artificial intelligence: it fills the shifts a
manager has defined, in an order chosen to satisfy the hardest requirements
first, and it refuses to break a hard rule.  Where it cannot fill a shift
without breaking one, it leaves the shift short and says so.  Everything it
produces is a draft for a manager to review.

Two changes matter most relative to the first version:

* fairness is measured in **hours against contracted target**, not shift count,
  so twenty ten-hour nights are not mistaken for twenty seven-and-a-half hour days
* bench allocation is **exclusive**: one person occupies one bench, so coverage
  reflects who is actually standing where rather than who happens to be on duty
  holding the right competency
"""

from __future__ import annotations

import random
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

from .models import (
    Bench, Config, ShiftRequirement, ShiftType, Staff,
)
from .timeutils import rest_hours

AUTO = "auto"
MANUAL = "manual"


@dataclass
class Assignment:
    """One person on one shift on one day."""

    day: date
    staff_id: str
    shift_code: str
    source: str = AUTO
    reason: str = ""                      # why the scheduler chose them

    @property
    def is_manual(self) -> bool:
        return self.source == MANUAL


@dataclass
class BenchAllocation:
    """One person standing at one bench for one shift."""

    day: date
    bench_name: str
    discipline: str
    shift_code: str
    staff_id: str
    source: str = AUTO


@dataclass
class Shortfall:
    """A requirement a shift could not meet.  Consumed by the issue engine."""

    day: date
    shift_code: str
    kind: str                             # staffing | competency | authoriser | …
    detail: str
    needed: int = 0
    found: int = 0
    discipline: str = ""
    bench_name: str = ""


@dataclass
class RestConflict:
    day: date
    staff_id: str
    previous_shift: str
    next_shift: str
    rest_interval_hours: float
    minimum_rest_hours: float


class Scheduler:
    """Builds a draft roster from a :class:`~labroster.models.Config`."""

    def __init__(self, config: Config,
                 manual_assignments: list[Assignment] | None = None):
        self.config = config
        self.rules = config.rules
        self.period = config.period
        self.days = config.period.days
        self.staff = config.staff
        self.by_id = config.staff_by_id()
        self.shifts = config.shifts
        self.shift_by_code = config.shift_by_code()
        self.rng = random.Random(self.rules.seed)

        # Contracted target for the period, per person.
        for person in self.staff:
            person.target_period_hours = (
                person.min_period_hours or config.target_hours(person))
            person.allocated_hours = 0.0
            person.credited_absence_hours = 0.0

        # Leave lookup, clipped to the roster period, plus the hours credited for
        # it.  Crediting absence is what stops the roster handing somebody extra
        # shifts to make up hours they were away for.
        self.leave: dict[tuple[date, str], str] = {}
        for entry in config.leave:
            first = max(entry.start, self.period.start)
            last = min(entry.end, self.period.end)
            day = first
            while day <= last:
                self.leave[(day, entry.staff_id)] = entry.code
                day += timedelta(days=1)
            person = self.by_id.get(entry.staff_id)
            if person is not None:
                person.credited_absence_hours = round(
                    person.credited_absence_hours
                    + config.credited_hours_for(entry, person), 4)

        self.assignments: dict[tuple[date, str], Assignment] = {}
        self.bench_allocations: list[BenchAllocation] = []
        self.recovery_days: dict[str, set[date]] = defaultdict(set)
        self.shortfalls: list[Shortfall] = []
        self.rest_conflicts: list[RestConflict] = []

        # Manual overrides are placed first and never displaced.
        self.manual_keys: set[tuple[date, str]] = set()
        for assignment in manual_assignments or []:
            key = (assignment.day, assignment.staff_id)
            assignment.source = MANUAL
            self.assignments[key] = assignment
            self.manual_keys.add(key)
            self._add_hours(assignment)

    # ------------------------------------------------------------------
    # bookkeeping
    # ------------------------------------------------------------------

    def _add_hours(self, assignment: Assignment) -> None:
        shift = self.shift_by_code.get(assignment.shift_code)
        person = self.by_id.get(assignment.staff_id)
        if shift and person:
            person.allocated_hours = round(person.allocated_hours + shift.hours, 4)

    def _remove_hours(self, assignment: Assignment) -> None:
        shift = self.shift_by_code.get(assignment.shift_code)
        person = self.by_id.get(assignment.staff_id)
        if shift and person:
            person.allocated_hours = round(person.allocated_hours - shift.hours, 4)

    def place(self, day: date, person: Staff, shift: ShiftType,
              reason: str = "") -> Assignment:
        assignment = Assignment(day=day, staff_id=person.staff_id,
                                shift_code=shift.code, reason=reason)
        self.assignments[(day, person.staff_id)] = assignment
        self._add_hours(assignment)
        return assignment

    def unplace(self, day: date, staff_id: str) -> Assignment | None:
        """Remove an assignment.  Manual overrides are never removed."""
        key = (day, staff_id)
        if key in self.manual_keys:
            return None
        assignment = self.assignments.pop(key, None)
        if assignment:
            self._remove_hours(assignment)
        return assignment

    def shift_on(self, staff_id: str, day: date) -> ShiftType | None:
        assignment = self.assignments.get((day, staff_id))
        return self.shift_by_code.get(assignment.shift_code) if assignment else None

    def assigned_to(self, day: date, shift: ShiftType) -> list[str]:
        return [staff_id for (d, staff_id), assignment in self.assignments.items()
                if d == day and assignment.shift_code == shift.code]

    def on_duty(self, day: date, shift: ShiftType | None = None) -> list[str]:
        if shift is not None:
            return self.assigned_to(day, shift)
        return [staff_id for (d, staff_id) in self.assignments if d == day]

    def is_weekend(self, day: date) -> bool:
        return day.weekday() in self.rules.weekend_days

    def on_leave(self, staff_id: str, day: date) -> bool:
        return (day, staff_id) in self.leave

    # -- counters used for fairness -------------------------------------

    def hours_of(self, staff_id: str) -> float:
        person = self.by_id.get(staff_id)
        return person.allocated_hours if person else 0.0

    def hours_deficit(self, staff_id: str) -> float:
        """How far below contracted target somebody is.  Higher = more owed work.

        Measured against *total accounted* hours, so credited absence counts.
        Somebody who has been on leave for a week is not treated as owing that
        week back.
        """
        person = self.by_id.get(staff_id)
        if not person:
            return 0.0
        return round(person.target_period_hours - person.total_accounted_hours, 4)

    def hours_in_week(self, staff_id: str, day: date) -> float:
        """Worked hours in the Monday-to-Sunday week containing ``day``."""
        monday = day - timedelta(days=day.weekday())
        total = 0.0
        for offset in range(7):
            assignment = self.assignments.get(
                (monday + timedelta(days=offset), staff_id))
            if assignment:
                shift = self.shift_by_code.get(assignment.shift_code)
                if shift:
                    total += shift.hours
        return round(total, 4)

    def count_nights(self, staff_id: str) -> int:
        return sum(1 for (_, sid), a in self.assignments.items()
                   if sid == staff_id
                   and self.shift_by_code[a.shift_code].is_night)

    def count_weekend_days(self, staff_id: str) -> int:
        return sum(1 for (day, sid) in self.assignments
                   if sid == staff_id and self.is_weekend(day))

    def count_saturdays(self, staff_id: str) -> int:
        return sum(1 for (day, sid) in self.assignments
                   if sid == staff_id and day.weekday() == 5)

    def count_sundays(self, staff_id: str) -> int:
        return sum(1 for (day, sid) in self.assignments
                   if sid == staff_id and day.weekday() == 6)

    def count_full_weekends(self, staff_id: str) -> int:
        """Saturday and the following Sunday both worked."""
        total = 0
        for day in self.days:
            if day.weekday() == 5 and (day, staff_id) in self.assignments:
                if (day + timedelta(days=1), staff_id) in self.assignments:
                    total += 1
        return total

    def count_shift_code(self, staff_id: str, code: str) -> int:
        return sum(1 for (_, sid), a in self.assignments.items()
                   if sid == staff_id and a.shift_code == code)

    def consecutive_run(self, staff_id: str, day: date) -> int:
        """Length of the unbroken run of worked days that would contain ``day``."""
        length = 1
        cursor = day - timedelta(days=1)
        while (cursor, staff_id) in self.assignments:
            length += 1
            cursor -= timedelta(days=1)
        cursor = day + timedelta(days=1)
        while (cursor, staff_id) in self.assignments:
            length += 1
            cursor += timedelta(days=1)
        return length

    def consecutive_nights(self, staff_id: str, day: date) -> int:
        length = 1
        cursor = day - timedelta(days=1)
        while True:
            shift = self.shift_on(staff_id, cursor)
            if shift is None or not shift.is_night:
                break
            length += 1
            cursor -= timedelta(days=1)
        cursor = day + timedelta(days=1)
        while True:
            shift = self.shift_on(staff_id, cursor)
            if shift is None or not shift.is_night:
                break
            length += 1
            cursor += timedelta(days=1)
        return length

    def days_worked_in_week(self, staff_id: str, day: date) -> int:
        monday = day - timedelta(days=day.weekday())
        return sum(1 for offset in range(7)
                   if (monday + timedelta(days=offset), staff_id) in self.assignments)

    # ------------------------------------------------------------------
    # rest, measured from real shift times
    # ------------------------------------------------------------------

    def rest_before(self, staff_id: str, day: date,
                    shift: ShiftType) -> float | None:
        """Hours between the previous shift finishing and this one starting."""
        previous_day = day - timedelta(days=1)
        previous = self.shift_on(staff_id, previous_day)
        if previous is None:
            return None
        _, previous_end = previous.window(previous_day)
        next_start, _ = shift.window(day)
        return rest_hours(previous_end, next_start)

    def rest_after(self, staff_id: str, day: date,
                   shift: ShiftType) -> float | None:
        following_day = day + timedelta(days=1)
        following = self.shift_on(staff_id, following_day)
        if following is None:
            return None
        _, this_end = shift.window(day)
        next_start, _ = following.window(following_day)
        return rest_hours(this_end, next_start)

    # ------------------------------------------------------------------
    # hard constraints
    # ------------------------------------------------------------------

    def can_assign(self, person: Staff, day: date, shift: ShiftType) -> bool:
        """Every check here is a hard rule the scheduler will not break."""
        staff_id = person.staff_id

        if (day, staff_id) in self.assignments:
            return False                                  # one shift per day
        if self.on_leave(staff_id, day):
            return False                                  # leave is absolute
        if day in self.recovery_days[staff_id]:
            return False                                  # recovery after nights

        # Contracted working pattern.
        if not person.availability.works_weekday(day, self.period.start):
            return False
        if not person.availability.permits_times(shift.start, shift.end):
            return False

        # Night and weekend eligibility.
        if shift.is_night and not person.nights_ok:
            return False
        if self.is_weekend(day) and not person.weekends_ok:
            return False
        if shift.is_night and person.max_nights:
            if self.count_nights(staff_id) >= person.max_nights:
                return False
        if self.is_weekend(day) and person.max_weekends:
            if self.count_full_weekends(staff_id) >= person.max_weekends \
                    and self.count_weekend_days(staff_id) >= person.max_weekends:
                return False

        # Days per week.
        if person.availability.max_days_per_week:
            if self.days_worked_in_week(staff_id, day) >= \
                    person.availability.max_days_per_week:
                return False

        # Consecutive days, personal limit overriding the organisational one.
        limit = person.max_consecutive_days or self.rules.max_consecutive_days
        if limit and self.consecutive_run(staff_id, day) > limit:
            return False

        # Consecutive nights.
        if shift.is_night and self.rules.max_consecutive_nights:
            if self.consecutive_nights(staff_id, day) > \
                    self.rules.max_consecutive_nights:
                return False

        # Rest, from real clock times, in both directions.
        minimum = self.rules.minimum_rest_hours
        if minimum > 0:
            before = self.rest_before(staff_id, day, shift)
            if before is not None and before < minimum:
                return False
            after = self.rest_after(staff_id, day, shift)
            if after is not None and after < minimum:
                return False

        # Hard hours ceilings, only where a manager has configured one.
        if person.max_period_hours:
            if person.allocated_hours + shift.hours > person.max_period_hours:
                return False
        if person.max_weekly_hours:
            if self.hours_in_week(staff_id, day) + shift.hours \
                    > person.max_weekly_hours:
                return False

        return True

    # ------------------------------------------------------------------
    # who should work, in what order
    # ------------------------------------------------------------------

    def candidates(self, day: date, shift: ShiftType,
                   predicate=None) -> list[Staff]:
        """Eligible staff, most owed work first.

        The pool is already restricted to people who *can* work this shift, which
        is what makes the fairness comparison fair: night counts are only ever
        compared between people eligible to work nights, and weekend counts
        between people eligible to work weekends.
        """
        pool = [person for person in self.staff
                if self.can_assign(person, day, shift)
                and (predicate is None or predicate(person))]

        weekend = self.is_weekend(day)

        def sort_key(person: Staff):
            staff_id = person.staff_id

            if shift.is_night and self.rules.share_nights_evenly:
                # The manager has asked for nights to be shared out, accepting
                # that it may cost some day-service cover.
                return (
                    self.count_nights(staff_id),
                    -self.hours_deficit(staff_id),
                    self.count_shift_code(staff_id, shift.code),
                    self.consecutive_run(staff_id, day),
                    self.rng.random(),
                )

            if weekend and not shift.is_night:
                # Weekend days lead on who has done fewest of them.
                #
                # Ordering by hours deficit first concentrates weekends on
                # full-time staff: a larger target means a larger absolute
                # deficit, so they win every early comparison and absorb the
                # weekends while part-time colleagues do none. A weekend day
                # costs nothing but itself, so sharing them out is close to free.
                return (
                    self.count_weekend_days(staff_id),  # least burdened first
                    -self.hours_deficit(staff_id),      # then most owed work
                    self.count_shift_code(staff_id, shift.code),
                    self.consecutive_run(staff_id, day),
                    self.rng.random(),
                )

            # Nights keep hours owed first, with night count as the tie-break.
            # Spreading night *blocks* more widely also spreads the recovery days
            # that follow them, and each block takes five days out of one person's
            # availability. Pushing them onto more people fragments cover for the
            # day service, which measurably costs shifts their requirements. Cover
            # is the point of the roster; fairness is a soft target, so it yields
            # here rather than the other way round.
            return (
                -self.hours_deficit(staff_id),      # most below target first
                self.count_nights(staff_id) if shift.is_night else 0,
                self.count_shift_code(staff_id, shift.code),
                self.consecutive_run(staff_id, day),
                self.rng.random(),
            )

        return sorted(pool, key=sort_key)

    # ------------------------------------------------------------------
    # requirement satisfaction
    # ------------------------------------------------------------------

    def competency_demand(self, day: date, shift: ShiftType,
                          requirement: ShiftRequirement) -> dict[str, int]:
        """How many independently competent people each discipline needs.

        Bench minimums are folded in here, which is the root fix for double
        counting: if three benches each need one competent person, the shift needs
        three *distinct* competent people, so the scheduler must roster three.
        """
        demand = {key.upper(): value
                  for key, value in requirement.required_competencies.items()}
        bench_need: Counter[str] = Counter()
        for bench in self.benches_for(day, shift):
            wanted = bench.required_on(day, self.rules.weekend_days)
            if wanted:
                bench_need[bench.discipline.upper()] += wanted
        for discipline, count in bench_need.items():
            demand[discipline] = max(demand.get(discipline, 0), count)
        return {k: v for k, v in demand.items() if v > 0}

    def benches_for(self, day: date, shift: ShiftType) -> list[Bench]:
        return [bench for bench in self.config.benches
                if bench.applies_on(day, self.rules.weekend_days)
                and bench.covers_shift(shift)]

    def _competent(self, staff_id: str, discipline: str, day: date) -> bool:
        return self.config.is_independently_competent(staff_id, discipline, day)

    def match_distinct(self, day: date, people: list[str], demand: dict[str, int],
                       capable=None) -> tuple[int, list[str]]:
        """Fill each demanded slot with a *different* person, as far as possible.

        This is what stops one multi-skilled scientist being counted as covering
        several sections at once.  Somebody competent in transfusion, haematology
        and coagulation can only stand at one bench, so three benches need three
        people.  Returns how many slots could be filled and which disciplines are
        left over.

        A straightforward augmenting-path matching: small numbers here, so
        clarity matters more than asymptotics.
        """
        capable = capable or (lambda staff_id, discipline:
                              self._competent(staff_id, discipline, day))
        slots: list[str] = []
        for discipline, count in sorted(demand.items()):
            slots.extend([discipline] * count)
        if not slots:
            return 0, []

        eligible = {index: [staff_id for staff_id in people
                            if capable(staff_id, discipline)]
                    for index, discipline in enumerate(slots)}
        taken_by: dict[str, int] = {}

        def assign(slot: int, seen: set[str]) -> bool:
            for staff_id in eligible[slot]:
                if staff_id in seen:
                    continue
                seen.add(staff_id)
                if staff_id not in taken_by or assign(taken_by[staff_id], seen):
                    taken_by[staff_id] = slot
                    return True
            return False

        filled = {slot for slot in eligible if assign(slot, set())}
        unmatched = [slots[slot] for slot in eligible if slot not in filled]
        return len(filled), unmatched

    def _authoriser(self, staff_id: str, discipline: str, day: date) -> bool:
        return self.config.can_authorise(staff_id, discipline, day)

    def outstanding_needs(self, day: date, shift: ShiftType,
                          requirement: ShiftRequirement) -> list[dict]:
        """What this shift still lacks, as independent conditions.

        Each need is checked on its own.  A senior member of staff does not
        satisfy a transfusion competency, and a coordinator does not satisfy an
        authoriser requirement.
        """
        present = self.assigned_to(day, shift)
        needs: list[dict] = []

        def shortfall(count_needed, have, kind, label, **extra):
            if count_needed and have < count_needed:
                needs.append({"kind": kind, "label": label,
                              "missing": count_needed - have, **extra})

        shortfall(requirement.min_registered,
                  sum(1 for sid in present if self.by_id[sid].registered),
                  "registered", "registered biomedical scientist")

        shortfall(requirement.min_senior,
                  sum(1 for sid in present if self.by_id[sid].is_senior),
                  "senior", "senior member of staff")

        if requirement.min_band and requirement.min_at_band:
            shortfall(requirement.min_at_band,
                      sum(1 for sid in present
                          if self.by_id[sid].meets_band(requirement.min_band)),
                      "band", f"Band {requirement.min_band:g} or above")

        shortfall(requirement.min_coordinators,
                  sum(1 for sid in present if self.by_id[sid].shift_coordinator),
                  "coordinator", "shift coordinator")

        # Two different kinds of competency requirement, checked differently.
        #
        # A *section* needs somebody standing at it, so three sections need three
        # distinct people even if one person holds all three competencies.
        #
        # A shift's own competency list is not about stations: "BT:1, HAEM:1" on a
        # single-handed night shift means one person competent in both, and asking
        # for two distinct people there could never be satisfied.
        bench_demand: Counter[str] = Counter()
        for bench in self.benches_for(day, shift):
            wanted = bench.required_on(day, self.rules.weekend_days)
            if wanted:
                bench_demand[bench.discipline.upper()] += wanted
        if bench_demand:
            _, unmatched = self.match_distinct(day, present, dict(bench_demand))
            for discipline in dict.fromkeys(unmatched):
                needs.append({
                    "kind": "competency",
                    "label": f"{discipline}-competent staff",
                    "missing": unmatched.count(discipline),
                    "discipline": discipline})

        for discipline, wanted in requirement.required_competencies.items():
            discipline = discipline.upper()
            if bench_demand.get(discipline, 0) >= wanted:
                continue                      # already covered by the section check
            have = sum(1 for sid in present
                       if self._competent(sid, discipline, day))
            if have < wanted:
                needs.append({
                    "kind": "competency",
                    "label": f"{discipline}-competent staff",
                    "missing": wanted - have,
                    "discipline": discipline})

        # Authorising results does not occupy a bench, so authorisers are matched
        # among themselves but may also be standing at a section.
        authoriser_demand = {key.upper(): value for key, value
                             in requirement.required_authorisers.items()}
        if authoriser_demand:
            _, unmatched = self.match_distinct(
                day, present, authoriser_demand,
                capable=lambda staff_id, discipline:
                    self._authoriser(staff_id, discipline, day))
            for discipline in dict.fromkeys(unmatched):
                needs.append({
                    "kind": "authoriser",
                    "label": f"{discipline} result authoriser",
                    "missing": unmatched.count(discipline),
                    "discipline": discipline})

        if requirement.min_trainers:
            have = 0
            for sid in present:
                records = self.config.competencies_for(sid)
                if any(record.can_train(day) for record in records):
                    have += 1
            shortfall(requirement.min_trainers, have, "trainer",
                      "trainer or supervisor")

        return needs

    def _need_predicate(self, need: dict, day: date):
        kind = need["kind"]
        if kind == "registered":
            return lambda person: person.registered
        if kind == "senior":
            return lambda person: person.is_senior
        if kind == "band":
            threshold = need.get("min_band", 0.0)
            return lambda person: person.meets_band(threshold)
        if kind == "coordinator":
            return lambda person: person.shift_coordinator
        if kind == "competency":
            discipline = need["discipline"]
            return lambda person: self._competent(person.staff_id, discipline, day)
        if kind == "authoriser":
            discipline = need["discipline"]
            return lambda person: self._authoriser(person.staff_id, discipline, day)
        if kind == "trainer":
            return lambda person: any(
                record.can_train(day)
                for record in self.config.competencies_for(person.staff_id))
        return lambda person: True

    def _trainee_cap_reached(self, day: date, shift: ShiftType,
                             requirement: ShiftRequirement) -> bool:
        if not requirement.max_trainees:
            return False
        trainees = sum(1 for sid in self.assigned_to(day, shift)
                       if self.by_id[sid].trainee)
        return trainees >= requirement.max_trainees

    # ------------------------------------------------------------------
    # building
    # ------------------------------------------------------------------

    def build(self) -> None:
        self.place_night_blocks()
        self.place_remaining_shifts()
        self.allocate_benches()
        self.record_rest_conflicts()

    def _applicable_shifts(self, day: date) -> list[ShiftType]:
        return [shift for shift in self.shifts
                if shift.applies_on(day, self.rules.weekend_days)]

    def place_night_blocks(self) -> None:
        """Nights first: most constrained, and they generate the recovery days."""
        for shift in [s for s in self.shifts if s.is_night]:
            for day in self.days:
                if not shift.applies_on(day, self.rules.weekend_days):
                    continue
                requirement = self.config.requirement_for(shift, day)
                target = max(requirement.min_staff, 0)
                while len(self.assigned_to(day, shift)) < target:
                    if not self._start_night_block(day, shift, requirement):
                        break

    def _start_night_block(self, day: date, shift: ShiftType,
                           requirement: ShiftRequirement) -> bool:
        block_days = [day + timedelta(days=offset)
                      for offset in range(self.rules.night_block_length)]
        block_days = [d for d in block_days
                      if d <= self.period.end
                      and shift.applies_on(d, self.rules.weekend_days)]
        if not block_days:
            return False

        needs = self.outstanding_needs(day, shift, requirement)
        predicate = self._need_predicate(needs[0], day) if needs else None

        for pool in ([self.candidates(day, shift, predicate)] if predicate else []) \
                + [self.candidates(day, shift)]:
            for person in pool:
                placed: list[date] = []
                for target_day in block_days:
                    if len(self.assigned_to(target_day, shift)) >= requirement.min_staff:
                        break
                    if not self.can_assign(person, target_day, shift):
                        break
                    self.place(target_day, person, shift, "night block")
                    placed.append(target_day)
                if not placed:
                    continue
                last = placed[-1]
                for offset in range(1, self.rules.recovery_days_after_nights + 1):
                    self.recovery_days[person.staff_id].add(last + timedelta(days=offset))
                return True
        return False

    def place_remaining_shifts(self) -> None:
        for day in self.days:
            shifts = [s for s in self._applicable_shifts(day) if not s.is_night]
            # Hardest first: the shift with the fewest eligible people.
            shifts.sort(key=lambda s: len(self.candidates(day, s)))
            for shift in shifts:
                self._fill_shift(day, shift)

    def _fill_shift(self, day: date, shift: ShiftType) -> None:
        requirement = self.config.requirement_for(shift, day)

        # 1. Satisfy specific conditions first, rarest capability first.
        guard = 0
        while guard < 60:
            guard += 1
            needs = self.outstanding_needs(day, shift, requirement)
            if not needs:
                break
            scored = []
            for need in needs:
                predicate = self._need_predicate(need, day)
                pool = self.candidates(day, shift, predicate)
                scored.append((len(pool), need, pool))
            scored.sort(key=lambda item: item[0])
            _, need, pool = scored[0]
            if not pool:
                self.shortfalls.append(Shortfall(
                    day=day, shift_code=shift.code,
                    kind=need["kind"],
                    detail=f"no available {need['label']}",
                    needed=need["missing"], found=0,
                    discipline=need.get("discipline", "")))
                break
            person = self._pick(pool, day, shift, requirement)
            if person is None:
                break
            self.place(day, person, shift, f"required: {need['label']}")

        # 2. Top up to the minimum headcount.
        while len(self.assigned_to(day, shift)) < requirement.min_staff:
            pool = self.candidates(day, shift)
            person = self._pick(pool, day, shift, requirement)
            if person is None:
                found = len(self.assigned_to(day, shift))
                self.shortfalls.append(Shortfall(
                    day=day, shift_code=shift.code, kind="staffing",
                    detail=(f"{found} of {requirement.min_staff} staff available"),
                    needed=requirement.min_staff, found=found))
                break
            self.place(day, person, shift, "staffing level")

    def _pick(self, pool: list[Staff], day: date, shift: ShiftType,
              requirement: ShiftRequirement) -> Staff | None:
        """First eligible candidate that does not breach the trainee cap."""
        cap_reached = self._trainee_cap_reached(day, shift, requirement)
        for person in pool:
            if cap_reached and person.trainee:
                continue
            return person
        return None

    # ------------------------------------------------------------------
    # bench allocation — one person, one bench
    # ------------------------------------------------------------------

    def allocate_benches(self) -> None:
        """Allocate people to benches exclusively.

        A person occupies at most ``max_simultaneous_bench_assignments`` benches
        (one by default).  Only independently competent staff are eligible, so
        somebody still in training does not appear to be providing cover.  The
        recorded allocations are what coverage is then judged on.
        """
        limit = max(1, self.rules.max_simultaneous_bench_assignments)
        rotation_load: Counter[str] = Counter()

        for day in self.days:
            for shift in self._applicable_shifts(day):
                benches = self.benches_for(day, shift)
                if not benches:
                    continue
                on_duty = self.assigned_to(day, shift)
                held: Counter[str] = Counter()

                # Scarcest discipline first, so a rare competency is not spent on
                # a bench that somebody else could have covered.
                def scarcity(bench: Bench) -> int:
                    return sum(1 for sid in on_duty
                               if self._competent(sid, bench.discipline, day))

                for bench in sorted(benches, key=scarcity):
                    wanted = bench.required_on(day, self.rules.weekend_days)
                    if not wanted:
                        continue
                    eligible = [
                        sid for sid in on_duty
                        if self._competent(sid, bench.discipline, day)
                        and held[sid] < limit
                        and (not bench.requires_authoriser
                             or self._authoriser(sid, bench.discipline, day))
                    ]
                    # Spread section work around for recency, then keep it stable.
                    eligible.sort(key=lambda sid: (rotation_load[sid], sid))
                    picked = eligible[:wanted]

                    for staff_id in picked:
                        held[staff_id] += 1
                        rotation_load[staff_id] += 1
                        self.bench_allocations.append(BenchAllocation(
                            day=day, bench_name=bench.name,
                            discipline=bench.discipline, shift_code=shift.code,
                            staff_id=staff_id))

                    if len(picked) < wanted:
                        self.shortfalls.append(Shortfall(
                            day=day, shift_code=shift.code, kind="bench",
                            detail=(f"{bench.name}: {len(picked)} of {wanted} "
                                    f"independently competent staff allocated"),
                            needed=wanted, found=len(picked),
                            discipline=bench.discipline, bench_name=bench.name))

    def bench_staff(self, day: date, bench_name: str,
                    shift_code: str | None = None) -> list[str]:
        return [allocation.staff_id for allocation in self.bench_allocations
                if allocation.day == day and allocation.bench_name == bench_name
                and (shift_code is None or allocation.shift_code == shift_code)]

    def simultaneous_bench_counts(self, day: date,
                                  shift_code: str) -> Counter[str]:
        """How many benches each person holds at once.  Used by the tests."""
        counts: Counter[str] = Counter()
        for allocation in self.bench_allocations:
            if allocation.day == day and allocation.shift_code == shift_code:
                counts[allocation.staff_id] += 1
        return counts

    # ------------------------------------------------------------------
    # rest reporting
    # ------------------------------------------------------------------

    def record_rest_conflicts(self) -> None:
        """Report every gap shorter than the configured rest rule.

        The scheduler will not create one, but a manual override can, so this is
        recalculated from the finished roster rather than assumed.
        """
        minimum = self.rules.minimum_rest_hours
        if minimum <= 0:
            return
        for (day, staff_id), assignment in sorted(self.assignments.items()):
            shift = self.shift_by_code.get(assignment.shift_code)
            if shift is None:
                continue
            previous_day = day - timedelta(days=1)
            previous = self.shift_on(staff_id, previous_day)
            if previous is None:
                continue
            _, previous_end = previous.window(previous_day)
            next_start, _ = shift.window(day)
            interval = rest_hours(previous_end, next_start)
            if interval < minimum:
                self.rest_conflicts.append(RestConflict(
                    day=day, staff_id=staff_id,
                    previous_shift=previous.code, next_shift=shift.code,
                    rest_interval_hours=interval,
                    minimum_rest_hours=minimum))
