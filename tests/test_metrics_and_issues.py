"""The two coverage measures, and consolidation of related issues."""

from __future__ import annotations

from datetime import timedelta

import pytest

from conftest import MONDAY, competent, make_config, make_shift, make_staff
from labroster import api
from labroster.analysis import CRITICAL, PASSED, REVIEW, Analysis
from labroster.models import Bench, CompetencyStatus, ShiftRequirement
from labroster.scheduler import Scheduler


def build(config):
    scheduler = Scheduler(config)
    scheduler.build()
    return scheduler, Analysis(scheduler)


# --- the two measures are genuinely different -------------------------------

def test_occupied_positions_do_not_imply_a_satisfactory_shift():
    """Every position filled, but nobody competent for the section.

    This is the case the old single "shift coverage" figure hid: it reported
    100% because the headcount was met.
    """
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2)],
        benches=[Bench(name="Morphology", discipline="MORPH", min_staff=1)],
        competencies=[competent("A", "HAEM")],
        days=1)
    _, analysis = build(config)
    metrics = analysis.metrics

    assert metrics["staffing_slot_coverage_percent"] == 100.0
    assert metrics["shifts_meeting_all_requirements_percent"] == 0.0
    assert "Section coverage" in metrics["requirement_failure_causes"]


def test_a_fully_satisfied_shift_counts_towards_both_measures():
    config = make_config(
        [make_staff("A", is_senior=True), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2, min_senior=1,
                                       required_competencies={"HAEM": 1})],
        competencies=[competent("A", "HAEM"), competent("B", "HAEM")],
        days=1)
    _, analysis = build(config)
    assert analysis.metrics["staffing_slot_coverage_percent"] == 100.0
    assert analysis.metrics["shifts_meeting_all_requirements_percent"] == 100.0
    assert analysis.metrics["requirement_failure_causes"] == {}


def test_the_compliance_measure_counts_shift_instances_not_slots():
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=1)], days=5)
    _, analysis = build(config)
    assert analysis.metrics["shift_instances"] == 5
    assert analysis.metrics["shift_instances_met"] == 5


def test_a_missing_senior_fails_compliance_but_not_slot_coverage():
    config = make_config(
        [make_staff("A", is_senior=False), make_staff("B", is_senior=False)],
        requirements=[ShiftRequirement("C", min_staff=2, min_senior=1)],
        days=1)
    _, analysis = build(config)
    assert analysis.metrics["staffing_slot_coverage_percent"] == 100.0
    assert analysis.metrics["shifts_meeting_all_requirements_percent"] == 0.0
    assert "Senior cover" in analysis.metrics["requirement_failure_causes"]


def test_the_trainee_limit_is_part_of_compliance():
    config = make_config(
        [make_staff("A", trainee=True), make_staff("B", trainee=True)],
        requirements=[ShiftRequirement("C", min_staff=2, max_trainees=1)],
        days=1)
    scheduler, analysis = build(config)
    # The scheduler honours the cap, so the shift is short rather than over-full.
    assert analysis.metrics["shifts_meeting_all_requirements_percent"] < 100.0


# --- issue consolidation ---------------------------------------------------

def test_a_missing_competency_produces_one_issue_not_two():
    """A section that cannot be staffed and the competency behind it are one problem."""
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2)],
        benches=[Bench(name="Morphology", discipline="MORPH", min_staff=1)],
        competencies=[competent("A", "HAEM")],
        days=1)
    _, analysis = build(config)
    critical = [issue for issue in analysis.issues
                if issue.severity == CRITICAL
                and issue.day == MONDAY]
    assert len(critical) == 1, (
        f"expected one consolidated issue, got: "
        f"{[issue.title for issue in critical]}")


def test_the_consolidated_issue_reads_as_a_management_problem():
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2)],
        benches=[Bench(name="Morphology", discipline="MORPH", min_staff=1)],
        competencies=[competent("A", "HAEM")],
        days=1)
    _, analysis = build(config)
    issue = [item for item in analysis.issues
             if item.severity == CRITICAL and item.day == MONDAY][0]
    assert "cannot be covered" in issue.title
    assert issue.required and issue.available == "0"
    assert issue.impact
    assert issue.review_point


def test_traceability_back_to_the_underlying_checks_is_kept():
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2)],
        benches=[Bench(name="Morphology", discipline="MORPH", min_staff=1)],
        competencies=[competent("A", "HAEM")],
        days=1)
    _, analysis = build(config)
    issue = [item for item in analysis.issues
             if item.severity == CRITICAL and item.day == MONDAY][0]
    assert issue.causes, "the underlying failed checks should be recorded"
    assert any("bench" in cause or "competency" in cause
               for cause in issue.causes)


def test_separate_shifts_still_produce_separate_issues():
    """Consolidation groups by cause, not across the whole period."""
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2)],
        benches=[Bench(name="Morphology", discipline="MORPH", min_staff=1)],
        competencies=[competent("A", "HAEM")],
        days=3)
    _, analysis = build(config)
    days = {issue.day for issue in analysis.issues
            if issue.severity == CRITICAL and issue.day}
    assert len(days) == 3


def test_expired_competency_is_not_counted_as_section_cover():
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2)],
        benches=[Bench(name="Blood Transfusion", discipline="BT", min_staff=1)],
        competencies=[competent("A", "BT",
                                expiry_date=MONDAY - timedelta(days=1))],
        days=1)
    _, analysis = build(config)
    assert analysis.metrics["shifts_meeting_all_requirements_percent"] == 0.0


# --- summaries -------------------------------------------------------------

def test_main_causes_are_summarised_for_the_manager():
    result = api.generate(api.demo_workbook_bytes())
    causes = result["dashboard"]["main_causes"]
    assert causes, "the dashboard should summarise what is driving the issues"
    assert all(isinstance(count, int) for count in causes.values())
    # Ordered most common first so the top line is where to start.
    counts = list(causes.values())
    assert counts == sorted(counts, reverse=True)


def test_the_dashboard_exposes_both_measures_by_their_new_names():
    result = api.generate(api.demo_workbook_bytes())
    dashboard = result["dashboard"]
    assert "staffing_slot_coverage_percent" in dashboard
    assert "shifts_meeting_all_requirements_percent" in dashboard
    assert "shift_coverage_percent" not in dashboard, \
        "the misleading combined figure should be gone"


def test_the_two_measures_differ_on_the_challenging_example():
    """The example laboratory should demonstrate why the distinction matters."""
    result = api.generate(api.demo_workbook_bytes())
    dashboard = result["dashboard"]
    assert dashboard["staffing_slot_coverage_percent"] > \
        dashboard["shifts_meeting_all_requirements_percent"]


def test_issue_payload_carries_the_consolidated_fields():
    result = api.generate(api.demo_workbook_bytes())
    critical = [issue for issue in result["issues"]
                if issue["severity"] == CRITICAL]
    assert critical
    assert any(issue.get("required") for issue in critical)
    assert any(issue.get("impact") for issue in critical)


# --- discipline terminology -------------------------------------------------

def test_disciplines_are_expressed_as_short_codes():
    """BT, HAEM, COAG and MORPH are the everyday shorthand in a laboratory.

    The requirement line must use the code, not an expanded competency name, so the
    workbook and the interface stay in the same terms.
    """
    result = api.generate(api.demo_workbook_bytes())
    text = " ".join(f"{issue['title']} {issue['explanation']} {issue['required']}"
                    for issue in result["issues"])
    for expanded in ("Blood film morphology scientist", "Haematology scientist",
                     "Coagulation scientist", "Blood Transfusion scientist"):
        assert expanded not in text, f"expanded name used instead of a code: {expanded}"
    assert "MORPH scientist" in text or "HAEM scientist" in text


def test_resilience_reports_disciplines_by_code():
    result = api.generate(api.demo_workbook_bytes())
    disciplines = {item["discipline"] for item in result["resilience"]}
    assert disciplines == {"BT", "HAEM", "COAG", "MORPH"}


def test_section_names_still_appear_in_the_issue_title():
    """Codes for the requirement, the section's own name for the headline."""
    config = make_config(
        [make_staff("A"), make_staff("B")],
        requirements=[ShiftRequirement("C", min_staff=2)],
        benches=[Bench(name="Morphology", discipline="MORPH", min_staff=1)],
        competencies=[competent("A", "HAEM")],
        days=1)
    _, analysis = build(config)
    issue = [item for item in analysis.issues
             if item.severity == CRITICAL and item.day == MONDAY][0]
    assert issue.title == "Morphology cannot be covered"
    assert "MORPH" in issue.required
