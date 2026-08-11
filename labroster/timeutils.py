"""Shift-time arithmetic: durations, rest intervals and week patterns.

Everything the roster needs to reason about *when* a shift happens lives here,
deliberately free of spreadsheet or scheduling concerns so it can be tested in
isolation.

Two things this fixes relative to the first version of the tool:

* shift length is derived from the configured start and end times rather than
  typed in by hand, so the stated times and the hours can no longer disagree
* a shift that runs past midnight is measured correctly, so 20:00-08:00 is
  twelve hours rather than a negative number
"""

from __future__ import annotations

import re
from datetime import date, datetime, time, timedelta

HOURS_PER_DAY = 24.0

# Excel hands times over in several shapes depending on how the cell was typed:
# a real time, a datetime, a fraction of a day, or free text.
_TIME_TEXT = re.compile(r"^(\d{1,2})\s*[:.,h]?\s*(\d{2})?\s*(am|pm)?$", re.I)


class TimeError(ValueError):
    """A shift time could not be understood, or a duration makes no sense."""


def parse_time(value, label: str = "time") -> time:
    """Read a shift time from anything a spreadsheet is likely to contain.

    Accepts ``datetime.time``, ``datetime.datetime``, an Excel day fraction
    (0.5 -> 12:00), and text such as ``08:00``, ``8.00``, ``0800``, ``8am``.
    """
    if value is None or (isinstance(value, str) and not value.strip()):
        raise TimeError(f"No {label} was given.")

    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value

    # Excel stores a bare time as a fraction of a day.
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        fraction = float(value)
        if not 0.0 <= fraction < 1.0:
            raise TimeError(
                f"{label} of {value} is not a time. Enter it as 08:00.")
        minutes = round(fraction * HOURS_PER_DAY * 60)
        return time(hour=(minutes // 60) % 24, minute=minutes % 60)

    text = str(value).strip()
    match = _TIME_TEXT.match(text)
    if not match:
        raise TimeError(f"{label} of '{value}' is not a time. Use 24-hour "
                        f"format such as 08:00 or 21:30.")

    hour = int(match.group(1))
    minute = int(match.group(2) or 0)
    meridiem = (match.group(3) or "").lower()

    if meridiem == "pm" and hour < 12:
        hour += 12
    elif meridiem == "am" and hour == 12:
        hour = 0

    # "0800" arrives as a single 4-digit group.
    if not match.group(2) and len(text) == 4 and text.isdigit():
        hour, minute = int(text[:2]), int(text[2:])

    if hour == 24 and minute == 0:
        hour = 0
    if not 0 <= hour <= 23 or not 0 <= minute <= 59:
        raise TimeError(f"{label} of '{value}' is not a valid time of day.")
    return time(hour=hour, minute=minute)


def crosses_midnight(start: time, end: time) -> bool:
    """True when the finish time falls on the day after the start time."""
    return end <= start


def duration_hours(start: time, end: time, label: str = "shift") -> float:
    """Length of a shift in hours, measured across midnight where needed.

    08:00 to 16:00 is 8.0.  20:00 to 08:00 is 12.0.  A shift that starts and
    finishes at the same time is rejected: in a laboratory rota that is a data
    entry mistake rather than a 24-hour or zero-length shift.
    """
    if start == end:
        raise TimeError(
            f"The {label} starts and finishes at {start.strftime('%H:%M')}. "
            f"Check the start and finish times.")

    minutes = (end.hour * 60 + end.minute) - (start.hour * 60 + start.minute)
    if minutes < 0:
        minutes += int(HOURS_PER_DAY * 60)
    return round(minutes / 60.0, 4)


def shift_window(day: date, start: time, end: time) -> tuple[datetime, datetime]:
    """The real start and finish moments of a shift worked on ``day``.

    The finish rolls onto the following calendar day when the shift runs
    through midnight, which is what makes rest intervals come out right.
    """
    begins = datetime.combine(day, start)
    finishes = datetime.combine(day, end)
    if crosses_midnight(start, end):
        finishes += timedelta(days=1)
    return begins, finishes


def rest_hours(previous_end: datetime, next_start: datetime) -> float:
    """Hours between finishing one shift and starting the next.

    Negative when the two shifts overlap, which is itself worth reporting.
    """
    return round((next_start - previous_end).total_seconds() / 3600.0, 4)


def overlaps(first_start: datetime, first_end: datetime,
             second_start: datetime, second_end: datetime) -> bool:
    """Whether two periods share any time.

    Not needed for whole-shift bench allocation, where one assignment covers the
    shift, but this is the primitive that part-shift bench blocks
    (morphology 09:00-13:00, coagulation 13:00-17:00) will be checked with, so
    the rule lives here from the start.
    """
    return first_start < second_end and second_start < first_end


def week_index(day: date, anchor: date) -> int:
    """Which week of a repeating pattern ``day`` falls in, counting from 0.

    Weeks are measured from the Monday of the anchor date's week so that an
    alternating pattern does not shift when a roster happens to start midweek.
    """
    anchor_monday = anchor - timedelta(days=anchor.weekday())
    return max(0, (day - anchor_monday).days) // 7


def pattern_week(day: date, anchor: date, cycle_length: int) -> int:
    """Position in a repeating cycle, 1-based, for ``cycle_length`` weeks.

    A two-week cycle returns 1, 2, 1, 2 … so a manager can write "Week 1" and
    "Week 2" in the workbook and mean what they expect.
    """
    if cycle_length < 1:
        raise TimeError("A working pattern cycle must be at least one week.")
    return (week_index(day, anchor) % cycle_length) + 1
