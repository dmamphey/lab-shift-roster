"""Small, explicit fixtures.

Each test builds the smallest laboratory that can demonstrate the rule under
test, so a failure points at one behaviour rather than at a month-long roster.
"""

from __future__ import annotations

from datetime import date, time

import pytest

from labroster.models import (
    Availability, Bench, Competency, CompetencyStatus, Config, LeaveEntry,
    LeaveType, Period, RosterDetails, Rules, ShiftRequirement, ShiftType, Staff,
)

MONDAY = date(2026, 9, 7)          # a Monday, so weekday maths is readable


def make_staff(staff_id, name=None, **kwargs) -> Staff:
    defaults = dict(
        name=name or f"Person {staff_id}",
        contracted_weekly_hours=37.5,
        fte=1.0,
        registered=True,
    )
    defaults.update(kwargs)
    return Staff(staff_id=staff_id, **defaults)


def working_days(*weekdays, cycle=1) -> Availability:
    """Availability restricted to the given weekday numbers (Monday = 0)."""
    return Availability(cycle_weeks=cycle, weekdays={1: set(weekdays)})


def make_shift(code="C", name=None, start="09:00", end="17:00", days="All",
               night=False) -> ShiftType:
    hour, minute = (int(part) for part in start.split(":"))
    end_hour, end_minute = (int(part) for part in end.split(":"))
    return ShiftType(code=code, name=name or code,
                     start=time(hour, minute), end=time(end_hour, end_minute),
                     days=days, is_night=night)


def competent(staff_id, discipline, status=CompetencyStatus.COMPETENT,
              **kwargs) -> Competency:
    return Competency(staff_id=staff_id, discipline=discipline, status=status,
                      **kwargs)


def make_config(staff, shifts=None, requirements=None, benches=None,
                competencies=None, leave=None, rules=None,
                start=MONDAY, days=1) -> Config:
    shifts = shifts or [make_shift()]
    if requirements is None:
        requirements = [ShiftRequirement(shift_code=shift.code, min_staff=1)
                        for shift in shifts]
    return Config(
        details=RosterDetails(),
        period=Period(start=start, end=start.fromordinal(start.toordinal() + days - 1)),
        rules=rules or Rules(),
        staff=list(staff),
        competencies=list(competencies or []),
        shifts=list(shifts),
        requirements=list(requirements),
        benches=list(benches or []),
        leave=list(leave or []),
        leave_types={"al": LeaveType("A/L", "Annual leave")},
    )


@pytest.fixture
def monday():
    return MONDAY
