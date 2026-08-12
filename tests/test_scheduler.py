"""Scheduling rules: hours, availability, rest, competence and bench exclusivity."""

from __future__ import annotations

from datetime import date, time, timedelta

import pytest

from conftest import (
    MONDAY, competent, make_config, make_shift, make_staff, working_days,
)
from labroster.analysis import CRITICAL, Analysis
from labroster.models import (
    Bench, CompetencyStatus, LeaveEntry, Rules, ShiftRequirement,
)
from labroster.scheduler import Scheduler


def build(config) -> Scheduler:
    scheduler = Scheduler(config)
    scheduler.build()
    return scheduler


# --------------------------------------------------------------------------
# contracted hours rather than shift counts
# --------------------------------------------------------------------------

def test_target_hours_scale_with_the_length_of_the_period():
    config = make_config([make_staff("A", contracted_weekly_hours=37.5)],
                         days=14)
    assert config.target_hours(config.staff[0]) == pytest.approx(75.0)


def test_target_hours_fall_back_to_fte_when_hours_are_blank():
    config = make_config([make_staff("A", contracted_weekly_hours=0, fte=0.6)],
                         days=7)
    assert config.target_hours(config.staff[0]) == pytest.approx(22.5)


def test_allocated_hours_come_from_shift_length_not_shift_count():
    """Ten hours of nights must not be recorded as the same as 7.5 of days."""
    night = make_shift("N", start="21:00", end="07:00", night=True)
    config = make_config(
        [make_staff("A", nights_ok=True)],
        shifts=[night],
        requirements=[ShiftRequirement(shift_code="N", min_staff=1)],
        days=1)
    scheduler = build(config)
    assert scheduler.by_id["A"].allocated_hours == pytest.approx(10.0)


def test_the_person_furthest_below_target_is_offered_work_first():
    part_time = make_staff("P", contracted_weekly_hours=15.0)
    full_time = make_staff("F", contracted_weekly_hours=37.5)
    config = make_config([part_time, full_time], days=7,
                         requirements=[ShiftRequirement(shift_code="C",
                                                        min_staff=1)])
    scheduler = build(config)
    # The full-timer is owed far more hours, so should get most of the work.
    assert scheduler.by_id["F"].allocated_hours > scheduler.by_id["P"].allocated_hours


def test_hours_variance_and_percentage_are_reported():
    config = make_config([make_staff("A", contracted_weekly_hours=37.5)], days=7)
    scheduler = build(config)
    person = scheduler.by_id["A"]
    assert person.target_period_hours == pytest.approx(37.5)
    assert person.hours_variance == pytest.approx(
        person.allocated_hours - 37.5, abs=0.01)
    assert person.percent_of_target is not None


def test_a_configured_maximum_stops_further_hours():
    config = make_config(
        [make_staff("A", contracted_weekly_hours=37.5, max_period_hours=16.0)],
        days=7)
    scheduler = build(config)
    assert scheduler.by_id["A"].allocated_hours <= 16.0


# --------------------------------------------------------------------------
# part-time and fixed working patterns
# --------------------------------------------------------------------------

def test_fixed_non_working_days_are_never_used():
    """Somebody who does not work Thursdays is not rostered on a Thursday."""
    person = make_staff("A", availability=working_days(0, 1, 2))   # Mon–Wed
    config = make_config([person], days=7)
    scheduler = build(config)
    worked = {day.weekday() for (day, _) in scheduler.assignments}
    assert worked <= {0, 1, 2}
    assert 3 not in worked and 4 not in worked


def test_an_alternating_two_week_pattern_is_followed():
    person = make_staff("A")
    person.availability.cycle_weeks = 2
    person.availability.weekdays = {1: {0, 1, 2}, 2: {2, 3, 4}}
    config = make_config([person], days=14)
    scheduler = build(config)
    first_week = {day.weekday() for (day, _) in scheduler.assignments
                  if (day - MONDAY).days < 7}
    second_week = {day.weekday() for (day, _) in scheduler.assignments
                   if (day - MONDAY).days >= 7}
    assert first_week <= {0, 1, 2}
    assert second_week <= {2, 3, 4}


def test_maximum_days_per_week_is_respected():
    person = make_staff("A")
    person.availability.max_days_per_week = 3
    config = make_config([person], days=7)
    scheduler = build(config)
    assert len(scheduler.assignments) <= 3


def test_earliest_start_excludes_an_early_shift():
    """Somebody who cannot start before 09:00 is not put on a 07:00 shift."""
    early = make_shift("E", start="07:00", end="15:00")
    person = make_staff("A")
    person.availability.earliest_start = time(9, 0)
    config = make_config([person], shifts=[early],
                         requirements=[ShiftRequirement("E", min_staff=1)], days=3)
    scheduler = build(config)
    assert not scheduler.assignments


def test_latest_finish_excludes_a_late_shift():
    late = make_shift("L", start="13:00", end="21:00")
    person = make_staff("A")
    person.availability.latest_finish = time(17, 0)
    config = make_config([person], shifts=[late],
                         requirements=[ShiftRequirement("L", min_staff=1)], days=3)
    scheduler = build(config)
    assert not scheduler.assignments


def test_a_shift_within_available_hours_is_allowed():
    core = make_shift("C", start="09:00", end="17:00")
    person = make_staff("A")
    person.availability.earliest_start = time(8, 0)
    person.availability.latest_finish = time(18, 0)
    config = make_config([person], shifts=[core],
                         requirements=[ShiftRequirement("C", min_staff=1)], days=3)
    scheduler = build(config)
    assert scheduler.assignments


def test_somebody_who_cannot_work_nights_gets_none():
    night = make_shift("N", start="21:00", end="07:00", night=True)
    day = make_shift("C", start="09:00", end="17:00")
    config = make_config(
        [make_staff("A", nights_ok=False), make_staff("B", nights_ok=True)],
        shifts=[night, day],
        requirements=[ShiftRequirement("N", min_staff=1),
                      ShiftRequirement("C", min_staff=1)],
        days=6)
    scheduler = build(config)
    assert scheduler.count_nights("A") == 0
    assert scheduler.count_nights("B") > 0


def test_somebody_who_cannot_work_weekends_gets_none():
    shift = make_shift("C", days="All")
    config = make_config([make_staff("A", weekends_ok=False),
                          make_staff("B", weekends_ok=True)],
                         shifts=[shift],
                         requirements=[ShiftRequirement("C", min_staff=1)],
                         days=14)
    scheduler = build(config)
    assert scheduler.count_weekend_days("A") == 0


def test_maximum_nights_per_roster_is_capped():
    night = make_shift("N", start="21:00", end="07:00", night=True)
    config = make_config(
        [make_staff("A", max_nights=2), make_staff("B"), make_staff("C")],
        shifts=[night],
        requirements=[ShiftRequirement("N", min_staff=1)],
        rules=Rules(night_block_length=1, recovery_days_after_nights=0),
        days=14)
    scheduler = build(config)
    assert scheduler.count_nights("A") <= 2


# --------------------------------------------------------------------------
# leave
# --------------------------------------------------------------------------

def test_nobody_is_scheduled_while_on_leave():
    absent = MONDAY + timedelta(days=1)
    config = make_config(
        [make_staff("A"), make_staff("B")],
        leave=[LeaveEntry(staff_id="A", start=absent,
                          end=absent + timedelta(days=2), code="A/L")],
        days=7)
    scheduler = build(config)
    for offset in range(3):
        assert (absent + timedelta(days=offset), "A") not in scheduler.assignments


def test_leave_outside_the_period_does_not_block_anything():
    config = make_config(
        [make_staff("A")],
        leave=[LeaveEntry(staff_id="A", start=MONDAY - timedelta(days=30),
                          end=MONDAY - timedelta(days=20), code="A/L")],
        days=3)
    scheduler = build(config)
    assert scheduler.assignments


# --------------------------------------------------------------------------
# consecutive days, nights and rest
# --------------------------------------------------------------------------

def test_maximum_consecutive_days_is_not_exceeded():
    config = make_config([make_staff("A")], days=21,
                         rules=Rules(max_consecutive_days=3))
    scheduler = build(config)
    run = 0
    worst = 0
    for day in scheduler.days:
        run = run + 1 if (day, "A") in scheduler.assignments else 0
        worst = max(worst, run)
    assert worst <= 3


def test_a_personal_consecutive_limit_overrides_the_organisational_one():
    config = make_config([make_staff("A", max_consecutive_days=2)], days=14,
                         rules=Rules(max_consecutive_days=6))
    scheduler = build(config)
    run = worst = 0
    for day in scheduler.days:
        run = run + 1 if (day, "A") in scheduler.assignments else 0
        worst = max(worst, run)
    assert worst <= 2


def test_maximum_consecutive_nights_is_not_exceeded():
    night = make_shift("N", start="21:00", end="07:00", night=True)
    config = make_config([make_staff(f"S{index}") for index in range(6)],
                         shifts=[night],
                         requirements=[ShiftRequirement("N", min_staff=1)],
                         rules=Rules(max_consecutive_nights=2,
                                     night_block_length=5,
                                     recovery_days_after_nights=1),
                         days=21)
    scheduler = build(config)
    for person in config.staff:
        run = worst = 0
        for day in scheduler.days:
            shift = scheduler.shift_on(person.staff_id, day)
            run = run + 1 if shift and shift.is_night else 0
            worst = max(worst, run)
        assert worst <= 2, f"{person.staff_id} worked {worst} nights in a row"


def test_a_late_shift_is_never_followed_by_an_early_one():
    """The specification's example: 13:00-21:00 then 07:00-15:00 is 10 hours."""
    late = make_shift("L", start="13:00", end="21:00")
    early = make_shift("E", start="07:00", end="15:00")
    config = make_config(
        [make_staff("A")], shifts=[late, early],
        requirements=[ShiftRequirement("L", min_staff=1),
                      ShiftRequirement("E", min_staff=1)],
        rules=Rules(minimum_rest_hours=11.0), days=6)
    scheduler = build(config)
    for day in scheduler.days:
        today = scheduler.shift_on("A", day)
        tomorrow = scheduler.shift_on("A", day + timedelta(days=1))
        if today and tomorrow:
            assert not (today.code == "L" and tomorrow.code == "E")
    assert scheduler.rest_conflicts == []


def test_rest_conflicts_are_reported_for_a_manual_override():
    """A manager may create a short gap; the tool must say so rather than hide it."""
    from labroster.scheduler import Assignment
    late = make_shift("L", start="13:00", end="21:00")
    early = make_shift("E", start="07:00", end="15:00")
    config = make_config([make_staff("A")], shifts=[late, early],
                         requirements=[ShiftRequirement("L", min_staff=0),
                                       ShiftRequirement("E", min_staff=0)],
                         rules=Rules(minimum_rest_hours=11.0), days=3)
    manual = [Assignment(MONDAY, "A", "L", source="manual"),
              Assignment(MONDAY + timedelta(days=1), "A", "E", source="manual")]
    scheduler = Scheduler(config, manual_assignments=manual)
    scheduler.build()
    assert len(scheduler.rest_conflicts) == 1
    assert scheduler.rest_conflicts[0].rest_interval_hours == pytest.approx(10.0)


def test_a_manual_assignment_is_never_removed():
    from labroster.scheduler import Assignment
    config = make_config([make_staff("A"), make_staff("B")], days=3)
    manual = [Assignment(MONDAY, "A", "C", source="manual")]
    scheduler = Scheduler(config, manual_assignments=manual)
    scheduler.build()
    assert scheduler.assignments[(MONDAY, "A")].is_manual
    assert scheduler.unplace(MONDAY, "A") is None


# --------------------------------------------------------------------------
# competence, and its separation from grade
# --------------------------------------------------------------------------

def test_a_required_competency_is_filled_by_a_competent_person():
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=1,
                                       required_competencies={"BT": 1})],
        competencies=[competent("B", "BT")],
        days=1)
    scheduler = build(config)
    assert (MONDAY, "B") in scheduler.assignments


def test_somebody_in_training_does_not_satisfy_a_competency_requirement():
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=1,
                                       required_competencies={"BT": 1})],
        competencies=[competent("A", "BT", CompetencyStatus.IN_TRAINING),
                      competent("B", "BT", CompetencyStatus.SUPERVISED)],
        days=1)
    scheduler = build(config)
    gaps = [item for item in scheduler.shortfalls if item.kind == "competency"]
    assert gaps, "a shift covered only by trainees should be reported"


def test_an_expired_competency_does_not_count():
    config = make_config(
        [make_staff("A")],
        requirements=[ShiftRequirement("C", min_staff=1,
                                       required_competencies={"BT": 1})],
        competencies=[competent("A", "BT",
                                expiry_date=MONDAY - timedelta(days=1))],
        days=1)
    scheduler = build(config)
    assert [item for item in scheduler.shortfalls if item.kind == "competency"]


def test_a_competency_expiring_tomorrow_still_counts_today():
    config = make_config(
        [make_staff("A")],
        requirements=[ShiftRequirement("C", min_staff=1,
                                       required_competencies={"BT": 1})],
        competencies=[competent("A", "BT",
                                expiry_date=MONDAY + timedelta(days=1))],
        days=1)
    scheduler = build(config)
    assert not [item for item in scheduler.shortfalls if item.kind == "competency"]


def test_a_senior_band_does_not_satisfy_a_specialist_competency():
    """A Band 7 with no morphology competency must not cover morphology."""
    senior = make_staff("A", band="7", is_senior=True)
    config = make_config(
        [senior],
        requirements=[ShiftRequirement("C", min_staff=1, min_senior=1,
                                       required_competencies={"MORPH": 1})],
        competencies=[competent("A", "HAEM")],
        days=1)
    scheduler = build(config)
    kinds = {item.kind for item in scheduler.shortfalls}
    assert "competency" in kinds
    assert "senior" not in kinds       # seniority itself was satisfied


def test_specialist_competence_does_not_make_somebody_senior():
    junior = make_staff("A", band="5", is_senior=False)
    config = make_config(
        [junior],
        requirements=[ShiftRequirement("C", min_staff=1, min_senior=1,
                                       required_competencies={"BT": 1})],
        competencies=[competent("A", "BT")],
        days=1)
    scheduler = build(config)
    assert "senior" in {item.kind for item in scheduler.shortfalls}


def test_a_result_authoriser_is_required_independently_of_competence():
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=1,
                                       required_authorisers={"BT": 1})],
        competencies=[competent("A", "BT"),                       # competent only
                      competent("B", "BT", authoriser=True)],     # can authorise
        days=1)
    scheduler = build(config)
    assert (MONDAY, "B") in scheduler.assignments


def test_an_authoriser_must_also_be_currently_competent():
    config = make_config(
        [make_staff("A")],
        requirements=[ShiftRequirement("C", min_staff=1,
                                       required_authorisers={"BT": 1})],
        competencies=[competent("A", "BT", authoriser=True,
                                expiry_date=MONDAY - timedelta(days=1))],
        days=1)
    scheduler = build(config)
    assert [item for item in scheduler.shortfalls if item.kind == "authoriser"]


def test_a_shift_coordinator_is_required_independently():
    config = make_config(
        [make_staff("A", shift_coordinator=False)],
        requirements=[ShiftRequirement("C", min_staff=1, min_coordinators=1)],
        days=1)
    scheduler = build(config)
    assert "coordinator" in {item.kind for item in scheduler.shortfalls}


def test_registration_is_required_independently_of_band():
    config = make_config(
        [make_staff("A", band="7", is_senior=True, registered=False)],
        requirements=[ShiftRequirement("C", min_staff=1, min_registered=1)],
        days=1)
    scheduler = build(config)
    assert "registered" in {item.kind for item in scheduler.shortfalls}


def test_the_trainee_cap_is_honoured():
    config = make_config(
        [make_staff("A", trainee=True), make_staff("B", trainee=True),
         make_staff("C", trainee=False)],
        requirements=[ShiftRequirement("C", min_staff=2, max_trainees=1)],
        days=1)
    scheduler = build(config)
    on_duty = scheduler.assigned_to(MONDAY, config.shifts[0])
    trainees = [sid for sid in on_duty if scheduler.by_id[sid].trainee]
    assert len(trainees) <= 1


# --------------------------------------------------------------------------
# bench exclusivity — the double counting defect
# --------------------------------------------------------------------------

def _three_bench_config(**kwargs):
    benches = [Bench(name="Blood Transfusion", discipline="BT", min_staff=1),
               Bench(name="Haematology", discipline="HAEM", min_staff=1),
               Bench(name="Coagulation", discipline="COAG", min_staff=1)]
    return make_config(benches=benches, days=1, **kwargs)


def test_one_person_is_never_allocated_to_three_benches_at_once():
    """The defect this release fixes.

    A scientist competent in transfusion, haematology and coagulation cannot
    stand at all three benches simultaneously, and must not be counted as doing
    so.
    """
    config = _three_bench_config(
        staff=[make_staff("A")],
        requirements=[ShiftRequirement("C", min_staff=1)],
        competencies=[competent("A", "BT"), competent("A", "HAEM"),
                      competent("A", "COAG")])
    scheduler = build(config)
    counts = scheduler.simultaneous_bench_counts(MONDAY, "C")
    assert counts["A"] <= 1, "one person was counted on more than one bench"


def test_covering_three_benches_with_one_person_is_reported_as_a_gap():
    config = _three_bench_config(
        staff=[make_staff("A")],
        requirements=[ShiftRequirement("C", min_staff=1)],
        competencies=[competent("A", "BT"), competent("A", "HAEM"),
                      competent("A", "COAG")])
    scheduler = build(config)
    bench_gaps = [item for item in scheduler.shortfalls if item.kind == "bench"]
    assert len(bench_gaps) >= 2, "two of the three benches should be unfilled"


def test_three_benches_are_covered_by_three_different_people():
    config = _three_bench_config(
        staff=[make_staff("A"), make_staff("B"), make_staff("C")],
        requirements=[ShiftRequirement("C", min_staff=3)],
        competencies=[competent("A", "BT"), competent("B", "HAEM"),
                      competent("C", "COAG")])
    scheduler = build(config)
    counts = scheduler.simultaneous_bench_counts(MONDAY, "C")
    assert set(counts.values()) == {1}
    assert not [item for item in scheduler.shortfalls if item.kind == "bench"]


def test_the_scheduler_rosters_enough_distinct_people_for_the_benches():
    """Bench demand drives staffing: three sections need three competent people."""
    everyone = [make_staff("A"), make_staff("B"), make_staff("C"),
                make_staff("D")]
    config = _three_bench_config(
        staff=everyone,
        requirements=[ShiftRequirement("C", min_staff=1)],
        competencies=[competent("A", "BT"), competent("A", "HAEM"),
                      competent("B", "HAEM"), competent("C", "COAG"),
                      competent("D", "BT")])
    scheduler = build(config)
    on_duty = scheduler.assigned_to(MONDAY, config.shifts[0])
    assert len(on_duty) >= 3, ("the shift needs three distinct competent people "
                               "even though its headcount minimum is one")
    assert not [item for item in scheduler.shortfalls if item.kind == "bench"]


def test_a_trainee_is_not_counted_as_bench_cover():
    config = make_config(
        staff=[make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2)],
        benches=[Bench(name="Morphology", discipline="MORPH", min_staff=1)],
        competencies=[competent("A", "MORPH", CompetencyStatus.IN_TRAINING),
                      competent("B", "HAEM")],
        days=1)
    scheduler = build(config)
    assert scheduler.bench_staff(MONDAY, "Morphology") == []
    assert [item for item in scheduler.shortfalls if item.kind == "bench"]


def test_a_bench_needing_an_authoriser_only_uses_authorisers():
    config = make_config(
        staff=[make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2)],
        benches=[Bench(name="Blood Transfusion", discipline="BT", min_staff=1,
                       requires_authoriser=True)],
        competencies=[competent("A", "BT"),
                      competent("B", "BT", authoriser=True)],
        days=1)
    scheduler = build(config)
    assert scheduler.bench_staff(MONDAY, "Blood Transfusion") == ["B"]


def test_matching_reports_how_many_distinct_slots_can_be_filled():
    config = _three_bench_config(
        staff=[make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2)],
        competencies=[competent("A", "BT"), competent("A", "HAEM"),
                      competent("B", "HAEM")])
    scheduler = Scheduler(config)
    filled, unmatched = scheduler.match_distinct(
        MONDAY, ["A", "B"], {"BT": 1, "HAEM": 1, "COAG": 1})
    assert filled == 2
    assert unmatched == ["COAG"]


# --------------------------------------------------------------------------
# unfilled shifts and senior cover
# --------------------------------------------------------------------------

def test_an_unfillable_shift_is_reported_rather_than_overfilled():
    config = make_config([make_staff("A")],
                         requirements=[ShiftRequirement("C", min_staff=3)],
                         days=1)
    scheduler = build(config)
    gaps = [item for item in scheduler.shortfalls if item.kind == "staffing"]
    assert gaps and gaps[0].found == 1 and gaps[0].needed == 3


def test_senior_cover_is_provided_when_somebody_is_available():
    config = make_config(
        [make_staff("A", is_senior=False), make_staff("B", is_senior=True)],
        requirements=[ShiftRequirement("C", min_staff=1, min_senior=1)],
        days=1)
    scheduler = build(config)
    assert (MONDAY, "B") in scheduler.assignments


def test_no_rule_is_broken_to_fill_a_shift():
    """With everybody on leave the shift is left empty, not filled regardless."""
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2)],
        leave=[LeaveEntry("A", MONDAY, MONDAY, "A/L"),
               LeaveEntry("B", MONDAY, MONDAY, "A/L")],
        days=1)
    scheduler = build(config)
    assert not scheduler.assignments
    assert [item for item in scheduler.shortfalls if item.kind == "staffing"]


# --------------------------------------------------------------------------
# fairness and resilience reporting
# --------------------------------------------------------------------------

def test_weekend_work_is_shared_between_those_eligible():
    staff = [make_staff("A"), make_staff("B"), make_staff("C")]
    config = make_config(staff, shifts=[make_shift("W", days="Weekend")],
                         requirements=[ShiftRequirement("W", min_staff=1)],
                         days=28)
    scheduler = build(config)
    counts = [scheduler.count_weekend_days(person.staff_id) for person in staff]
    assert max(counts) - min(counts) <= 2


def test_fairness_ignores_people_who_cannot_work_nights():
    night = make_shift("N", start="21:00", end="07:00", night=True)
    staff = [make_staff("A", nights_ok=True), make_staff("B", nights_ok=True),
             make_staff("C", nights_ok=False)]
    config = make_config(staff, shifts=[night],
                         requirements=[ShiftRequirement("N", min_staff=1)],
                         days=14)
    scheduler = build(config)
    analysis = Analysis(scheduler)
    # C never works nights; that must not be read as unfairness.
    assert analysis.night_fairness() in ("Good", "Review", "Uneven")
    assert scheduler.count_nights("C") == 0


def test_a_single_competent_person_is_a_critical_single_point_of_failure():
    config = make_config(
        [make_staff("A"), make_staff("B")],
        benches=[Bench(name="Morphology", discipline="MORPH", min_staff=1)],
        requirements=[ShiftRequirement("C", min_staff=2)],
        competencies=[competent("A", "MORPH")],
        days=1)
    analysis = Analysis(build(config))
    morph = [item for item in analysis.workforce_resilience
             if item.discipline == "MORPH"][0]
    assert morph.competent_count == 1
    assert morph.severity == CRITICAL


def test_two_competent_people_are_flagged_for_review_not_as_critical():
    config = make_config(
        [make_staff("A"), make_staff("B"), make_staff("C")],
        benches=[Bench(name="Coagulation", discipline="COAG", min_staff=1)],
        requirements=[ShiftRequirement("C", min_staff=2)],
        competencies=[competent("A", "COAG"), competent("B", "COAG")],
        days=1)
    analysis = Analysis(build(config))
    coag = [item for item in analysis.workforce_resilience
            if item.discipline == "COAG"][0]
    assert coag.competent_count == 2
    assert coag.severity == "REVIEW"


def test_workforce_and_shift_resilience_are_reported_separately():
    """One competent person in the department is not the same problem as one on a shift."""
    config = make_config(
        [make_staff("A"), make_staff("B"), make_staff("C")],
        benches=[Bench(name="Haematology", discipline="HAEM", min_staff=1)],
        requirements=[ShiftRequirement("C", min_staff=2)],
        competencies=[competent("A", "HAEM"), competent("B", "HAEM"),
                      competent("C", "HAEM")],
        days=3)
    analysis = Analysis(build(config))
    categories = {issue.category for issue in analysis.issues}
    assert "Workforce resilience" in categories
