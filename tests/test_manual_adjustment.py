"""Manual roster adjustment: removing, adding, persistence and consequences."""

from __future__ import annotations

from datetime import timedelta

import pytest

from conftest import MONDAY, competent, make_config, make_shift, make_staff
from labroster import api
from labroster.analysis import CRITICAL, Analysis
from labroster.models import Bench, LeaveEntry, ShiftRequirement
from labroster.scheduler import Scheduler


def build(config, **kwargs):
    scheduler = Scheduler(config, **kwargs)
    scheduler.build()
    return scheduler


# --- removing somebody -----------------------------------------------------

def test_removing_somebody_keeps_them_off_that_day():
    """Without a block the scheduler would simply put them back."""
    config = make_config([make_staff("A"), make_staff("B")],
                         requirements=[ShiftRequirement("C", min_staff=1)],
                         days=1)
    plain = build(config)
    assert (MONDAY, "A") in plain.assignments or (MONDAY, "B") in plain.assignments

    config = make_config([make_staff("A"), make_staff("B")],
                         requirements=[ShiftRequirement("C", min_staff=1)],
                         days=1)
    scheduler = build(config, blocked={(MONDAY, "A")})
    assert (MONDAY, "A") not in scheduler.assignments


def test_the_shift_is_refilled_by_somebody_else_where_possible():
    config = make_config([make_staff("A"), make_staff("B")],
                         requirements=[ShiftRequirement("C", min_staff=1)],
                         days=1)
    scheduler = build(config, blocked={(MONDAY, "A")})
    assert (MONDAY, "B") in scheduler.assignments


def test_removing_the_only_person_available_that_day_is_reported():
    """A change that breaks cover must show up, not pass quietly.

    Three people hold the competency, so the workforce itself is resilient; only
    one of them works Mondays, so taking that person off leaves the section
    uncovered on the Monday.
    """
    from conftest import working_days

    def department():
        return make_config(
            [make_staff("A", availability=working_days(0)),          # Monday
             make_staff("B", availability=working_days(1)),          # Tuesday
             make_staff("C", availability=working_days(2))],         # Wednesday
            requirements=[ShiftRequirement("C", min_staff=1)],
            benches=[Bench(name="Morphology", discipline="MORPH", min_staff=1)],
            competencies=[competent("A", "MORPH"), competent("B", "MORPH"),
                          competent("C", "MORPH")],
            start=MONDAY, days=1)

    before = Analysis(build(department()))
    assert before.metrics["critical_count"] == 0
    assert before.metrics["shifts_meeting_all_requirements_percent"] == 100.0

    after = Analysis(build(department(), blocked={(MONDAY, "A")}))
    assert after.metrics["critical_count"] > 0
    assert after.metrics["shifts_meeting_all_requirements_percent"] < 100.0


# --- adding somebody -------------------------------------------------------

def test_an_added_assignment_is_marked_as_manual():
    data = api.balanced_workbook_bytes()
    plain = api.generate(data)
    day = plain["roster"]["days"][0]["iso"]

    # Somebody not working that day.
    spare = next(row["staff_id"] for row in plain["roster"]["rows"]
                 if plain["roster"]["rows"] and
                 row["cells"][day]["kind"] == "off")

    result = api.generate(data, adjustments=[
        {"date": day, "shift": "C", "remove": None, "add": spare,
         "note": "extra cover"}])
    cell = next(row["cells"][day] for row in result["roster"]["rows"]
                if row["staff_id"] == spare)
    assert cell["kind"] == "shift"
    assert cell["source"] == "manual"


def test_a_manual_assignment_is_not_overwritten_when_rebuilt():
    data = api.balanced_workbook_bytes()
    plain = api.generate(data)
    day = plain["roster"]["days"][0]["iso"]
    spare = next(row["staff_id"] for row in plain["roster"]["rows"]
                 if row["cells"][day]["kind"] == "off")

    adjustments = [{"date": day, "shift": "C", "remove": None, "add": spare,
                    "note": ""}]
    # A different generation ID rebuilds the whole roster from scratch.
    for alternative in (11, 22, 33):
        result = api.generate(data, alternative=alternative,
                              adjustments=adjustments)
        cell = next(row["cells"][day] for row in result["roster"]["rows"]
                    if row["staff_id"] == spare)
        assert cell["source"] == "manual", \
            f"the manual assignment was lost at generation {alternative}"


def test_undoing_an_adjustment_restores_the_automatic_roster():
    data = api.balanced_workbook_bytes()
    plain = api.generate(data)
    day = plain["roster"]["days"][0]["iso"]
    spare = next(row["staff_id"] for row in plain["roster"]["rows"]
                 if row["cells"][day]["kind"] == "off")

    changed = api.generate(data, adjustments=[
        {"date": day, "shift": "C", "remove": None, "add": spare, "note": ""}])
    assert changed["manual_count"] == 1

    undone = api.generate(data, adjustments=[])
    assert undone["manual_count"] == 0
    for row_a, row_b in zip(plain["roster"]["rows"], undone["roster"]["rows"]):
        assert row_a["cells"] == row_b["cells"], \
            "undoing should give back exactly the automatic roster"


def test_every_check_is_rerun_after_an_adjustment():
    data = api.balanced_workbook_bytes()
    plain = api.generate(data)
    day = plain["roster"]["days"][0]["iso"]
    victim = next(row["staff_id"] for row in plain["roster"]["rows"]
                  if row["cells"][day]["kind"] == "shift")

    result = api.generate(data, adjustments=[
        {"date": day, "shift": plain["roster"]["rows"][0]["cells"][day].get("code", "C"),
         "remove": victim, "add": None, "note": ""}])
    # The whole analysis is regenerated, not patched.
    for key in ("staffing_slot_coverage_percent",
                "shifts_meeting_all_requirements_percent",
                "critical_count", "rest_conflicts", "weekend_fairness"):
        assert key in result["dashboard"]
    assert result["hours"], "hours are recalculated"
    assert result["resilience"], "resilience is recalculated"


# --- who is offered as a replacement --------------------------------------

def test_eligible_staff_excludes_people_on_leave():
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=1)],
        leave=[LeaveEntry("B", MONDAY, MONDAY, "A/L")],
        days=1)
    scheduler = build(config)
    options = api.eligible_staff(scheduler, MONDAY.isoformat(), "C")
    entry = next(item for item in options if item["staff_id"] == "B")
    assert entry["eligible"] is False
    assert "A/L" in entry["why_not"]


def test_eligible_staff_excludes_a_day_they_do_not_work():
    from conftest import working_days
    config = make_config(
        [make_staff("A", availability=working_days(0, 1))],   # Mon and Tue only
        requirements=[ShiftRequirement("C", min_staff=1)],
        days=5)
    scheduler = build(config)
    thursday = (MONDAY + timedelta(days=3)).isoformat()
    entry = api.eligible_staff(scheduler, thursday, "C")[0]
    assert entry["eligible"] is False
    assert "Thursday" in entry["why_not"]


def test_eligible_staff_excludes_somebody_who_cannot_work_nights():
    night = make_shift("N", start="21:00", end="07:00", night=True)
    config = make_config([make_staff("A", nights_ok=False)],
                         shifts=[night],
                         requirements=[ShiftRequirement("N", min_staff=0)],
                         days=1)
    scheduler = build(config)
    entry = api.eligible_staff(scheduler, MONDAY.isoformat(), "N")[0]
    assert entry["eligible"] is False
    assert "nights" in entry["why_not"]


def test_people_already_on_the_shift_are_flagged_rather_than_hidden():
    config = make_config([make_staff("A")],
                         requirements=[ShiftRequirement("C", min_staff=1)],
                         days=1)
    scheduler = build(config)
    entry = api.eligible_staff(scheduler, MONDAY.isoformat(), "C")[0]
    assert entry["on_this_shift"] is True


def test_eligible_staff_reports_current_competencies():
    config = make_config(
        [make_staff("A")], requirements=[ShiftRequirement("C", min_staff=1)],
        competencies=[competent("A", "BT"), competent("A", "HAEM")], days=1)
    scheduler = build(config)
    entry = api.eligible_staff(scheduler, MONDAY.isoformat(), "C")[0]
    assert entry["competencies"] == ["BT", "HAEM"]


def test_an_expired_competency_is_not_listed_as_current():
    config = make_config(
        [make_staff("A")], requirements=[ShiftRequirement("C", min_staff=1)],
        competencies=[competent("A", "BT",
                                expiry_date=MONDAY - timedelta(days=1))],
        days=1)
    scheduler = build(config)
    entry = api.eligible_staff(scheduler, MONDAY.isoformat(), "C")[0]
    assert entry["competencies"] == []


def test_an_unknown_shift_code_returns_nothing_rather_than_failing():
    config = make_config([make_staff("A")], days=1)
    scheduler = build(config)
    assert api.eligible_staff(scheduler, MONDAY.isoformat(), "ZZZ") == []
