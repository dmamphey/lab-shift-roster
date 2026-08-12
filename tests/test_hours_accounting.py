"""Contracted hours accounting, including credited absence.

The behaviour under test: hours somebody was away for must not be clawed back
with extra shifts.  Before this was added, a person with a week of annual leave
was given the same workload as colleagues who had not been away, compressed into
fewer days.
"""

from __future__ import annotations

from datetime import date, time, timedelta

import pytest

from conftest import MONDAY, make_config, make_shift, make_staff, working_days
from labroster.analysis import Analysis
from labroster.models import (
    CREDIT_FIXED, CREDIT_FROM_PATTERN, CREDIT_NONE, LeaveEntry, LeaveType,
    ShiftRequirement,
)
from labroster.scheduler import Scheduler


def four_weeks(staff, leave=None, leave_types=None, min_staff=4):
    """A four-week period of weekday day shifts."""
    shift = make_shift("C", start="09:00", end="17:00", days="Weekday")
    config = make_config(
        staff, shifts=[shift],
        requirements=[ShiftRequirement("C", min_staff=min_staff)],
        leave=leave or [], days=28)
    if leave_types:
        config.leave_types = leave_types
    return config


def build(config):
    scheduler = Scheduler(config)
    scheduler.build()
    return scheduler


# --- credited absence ------------------------------------------------------

def test_annual_leave_is_credited_towards_contracted_hours():
    person = make_staff("A", contracted_weekly_hours=37.5,
                        availability=working_days(0, 1, 2, 3, 4))
    config = four_weeks([person], leave=[
        LeaveEntry("A", MONDAY + timedelta(days=7),
                   MONDAY + timedelta(days=11), "A/L")], min_staff=1)
    build(config)
    # Five working days at 7.5 hours each.
    assert person.credited_absence_hours == pytest.approx(37.5)


def test_credited_hours_follow_a_part_timers_own_pattern():
    """A part-timer absent for a week is credited a part-time week."""
    part_time = make_staff("P", contracted_weekly_hours=22.5,
                           availability=working_days(0, 1, 2))     # Mon–Wed
    config = four_weeks([part_time], leave=[
        LeaveEntry("P", MONDAY + timedelta(days=7),
                   MONDAY + timedelta(days=11), "A/L")], min_staff=1)
    build(config)
    # Only Monday, Tuesday and Wednesday are days they would have worked.
    assert part_time.credited_absence_hours == pytest.approx(22.5)


def test_total_accounted_hours_is_worked_plus_credited():
    person = make_staff("A", contracted_weekly_hours=37.5,
                        availability=working_days(0, 1, 2, 3, 4))
    config = four_weeks([person], leave=[
        LeaveEntry("A", MONDAY + timedelta(days=7),
                   MONDAY + timedelta(days=11), "A/L")], min_staff=1)
    build(config)
    assert person.total_accounted_hours == pytest.approx(
        person.allocated_hours + person.credited_absence_hours)
    assert person.hours_variance == pytest.approx(
        person.total_accounted_hours - person.target_period_hours, abs=0.01)


def test_hours_lost_to_leave_are_not_made_up_with_extra_shifts():
    """The defect this accounting exists to prevent."""
    on_leave = make_staff("L", contracted_weekly_hours=37.5)
    others = [make_staff(f"S{index}", contracted_weekly_hours=37.5)
              for index in range(5)]
    config = four_weeks([on_leave, *others], leave=[
        LeaveEntry("L", MONDAY + timedelta(days=7),
                   MONDAY + timedelta(days=11), "A/L")])
    scheduler = build(config)

    average_other = sum(p.allocated_hours for p in others) / len(others)
    # A week away should mean roughly a week less work, not the same amount.
    assert on_leave.allocated_hours < average_other - 15, (
        f"person on leave worked {on_leave.allocated_hours} against a colleague "
        f"average of {average_other}")


def test_sickness_does_not_trigger_make_up_shifts():
    sick = make_staff("K", contracted_weekly_hours=37.5)
    others = [make_staff(f"S{index}", contracted_weekly_hours=37.5)
              for index in range(5)]
    config = four_weeks([sick, *others], leave=[
        LeaveEntry("K", MONDAY, MONDAY + timedelta(days=11), "S/L")])
    config.leave_types = {
        "sl": LeaveType("S/L", "Sickness absence", credits_hours=True)}
    build(config)
    average_other = sum(p.allocated_hours for p in others) / len(others)
    assert sick.allocated_hours < average_other
    assert sick.credited_absence_hours > 0


def test_long_term_absence_does_not_produce_a_huge_deficit():
    """Somebody absent for the whole period should not look wildly under target."""
    absent = make_staff("M", contracted_weekly_hours=37.5)
    config = four_weeks([absent, make_staff("B"), make_staff("C"),
                         make_staff("D"), make_staff("E")], leave=[
        LeaveEntry("M", MONDAY - timedelta(days=5),
                   MONDAY + timedelta(days=40), "M/L")])
    build(config)
    assert absent.allocated_hours == 0
    assert absent.credited_absence_hours > 100
    # Within a reasonable distance of target rather than 150 hours short.
    assert abs(absent.hours_variance) < 30


def test_a_leave_type_can_be_configured_not_to_credit_hours():
    person = make_staff("A", contracted_weekly_hours=37.5)
    config = four_weeks([person], leave=[
        LeaveEntry("A", MONDAY, MONDAY + timedelta(days=4), "U/A")],
        min_staff=1)
    config.leave_types = {"ua": LeaveType("U/A", "Unpaid absence",
                                          credits_hours=False)}
    build(config)
    assert person.credited_absence_hours == 0


def test_a_fixed_daily_credit_can_be_configured():
    person = make_staff("A", contracted_weekly_hours=37.5,
                        availability=working_days(0, 1, 2, 3, 4))
    config = four_weeks([person], leave=[
        LeaveEntry("A", MONDAY, MONDAY + timedelta(days=4), "S/D")],
        min_staff=1)
    config.leave_types = {"sd": LeaveType(
        "S/D", "Study day", credits_hours=True,
        credited_method=CREDIT_FIXED, fixed_daily_hours=4.0)}
    build(config)
    assert person.credited_absence_hours == pytest.approx(20.0)   # 5 x 4h


def test_an_explicit_credited_figure_on_the_leave_row_wins():
    person = make_staff("A", contracted_weekly_hours=37.5)
    config = four_weeks([person], leave=[
        LeaveEntry("A", MONDAY, MONDAY + timedelta(days=4), "A/L",
                   credited_hours=12.0)], min_staff=1)
    build(config)
    assert person.credited_absence_hours == pytest.approx(12.0)


def test_credit_is_pro_rated_when_leave_straddles_the_period_edge():
    person = make_staff("A", contracted_weekly_hours=37.5)
    # Ten-day absence, only the last five days inside the period.
    config = four_weeks([person], leave=[
        LeaveEntry("A", MONDAY - timedelta(days=5),
                   MONDAY + timedelta(days=4), "A/L",
                   credited_hours=20.0)], min_staff=1)
    build(config)
    assert person.credited_absence_hours == pytest.approx(10.0, abs=0.5)


# --- reporting -------------------------------------------------------------

def test_the_hours_summary_shows_the_full_accounting():
    person = make_staff("A", contracted_weekly_hours=37.5,
                        availability=working_days(0, 1, 2, 3, 4))
    config = four_weeks([person, make_staff("B"), make_staff("C"),
                         make_staff("D")], leave=[
        LeaveEntry("A", MONDAY + timedelta(days=7),
                   MONDAY + timedelta(days=11), "A/L")])
    analysis = Analysis(build(config))
    row = [item for item in analysis.hours_rows if item.staff_id == "A"][0]
    assert row.worked_hours > 0
    assert row.credited_absence_hours == pytest.approx(37.5)
    assert row.total_accounted_hours == pytest.approx(
        row.worked_hours + row.credited_absence_hours)
    assert row.variance == pytest.approx(
        row.total_accounted_hours - row.target_hours, abs=0.01)


def test_the_hours_payload_exposes_the_new_figures():
    config = four_weeks([make_staff("A"), make_staff("B"), make_staff("C"),
                         make_staff("D")])
    analysis = Analysis(build(config))
    row = analysis.hours_payload()[0]
    for key in ("target", "worked", "credited", "accounted", "variance"):
        assert key in row


# --- weekly ceiling --------------------------------------------------------

def test_a_maximum_weekly_hours_ceiling_is_respected():
    person = make_staff("A", contracted_weekly_hours=37.5, max_weekly_hours=16.0)
    config = four_weeks([person], min_staff=1)
    scheduler = build(config)
    for day in scheduler.days:
        if day.weekday() == 0:
            assert scheduler.hours_in_week("A", day) <= 16.0


# --- the model must stay hours-based, not shift-count-based -----------------

def test_workload_is_compared_by_hours_not_shift_count():
    """A long-shift worker and a short-shift worker should even out on hours."""
    long_shift = make_shift("N", start="20:00", end="08:00", night=True)
    short_shift = make_shift("C", start="09:00", end="14:00")
    staff = [make_staff(f"S{index}", contracted_weekly_hours=37.5)
             for index in range(4)]
    config = make_config(
        staff, shifts=[long_shift, short_shift],
        requirements=[ShiftRequirement("N", min_staff=1),
                      ShiftRequirement("C", min_staff=1)],
        days=28)
    scheduler = build(config)
    hours = [person.allocated_hours for person in staff]
    counts = [sum(1 for (_, sid) in scheduler.assignments
                  if sid == person.staff_id) for person in staff]
    spread_hours = max(hours) - min(hours)
    # Hours should be closer together than a naive equal-shifts split would give.
    assert spread_hours < 12 * max(1, (max(counts) - min(counts)) + 1)
    assert sum(hours) > 0
