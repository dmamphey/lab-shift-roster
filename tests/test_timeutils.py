"""Tests for shift-time arithmetic: durations, midnight crossing, rest, patterns."""

from datetime import date, datetime, time, timedelta

import pytest

from labroster.timeutils import (
    TimeError, crosses_midnight, duration_hours, overlaps, parse_time,
    pattern_week, rest_hours, shift_window, week_index,
)


# --- parsing ---------------------------------------------------------------

@pytest.mark.parametrize("value,expected", [
    ("08:00", time(8, 0)),
    ("8:00", time(8, 0)),
    ("08.00", time(8, 0)),
    ("0800", time(8, 0)),
    ("21:30", time(21, 30)),
    ("8am", time(8, 0)),
    ("9pm", time(21, 0)),
    ("12am", time(0, 0)),
    ("12pm", time(12, 0)),
    ("24:00", time(0, 0)),
    (time(13, 45), time(13, 45)),
    (datetime(2026, 9, 1, 7, 15), time(7, 15)),
    (0.5, time(12, 0)),                      # Excel day fraction
    (0.875, time(21, 0)),
])
def test_parse_time_accepts_spreadsheet_shapes(value, expected):
    assert parse_time(value) == expected


@pytest.mark.parametrize("value", ["", None, "half eight", "99:00", "8:99", 1.5])
def test_parse_time_rejects_nonsense(value):
    with pytest.raises(TimeError):
        parse_time(value, "start time")


def test_parse_time_error_names_the_field():
    with pytest.raises(TimeError, match="finish time"):
        parse_time("banana", "finish time")


# --- durations, including across midnight ----------------------------------

@pytest.mark.parametrize("start,end,hours", [
    ("08:00", "16:00", 8.0),      # the specification's first example
    ("20:00", "08:00", 12.0),     # the specification's second example
    ("21:00", "07:00", 10.0),     # night shift
    ("09:00", "17:30", 8.5),
    ("13:00", "21:00", 8.0),      # late
    ("07:00", "15:00", 8.0),      # early
    ("22:30", "06:15", 7.75),
    ("23:59", "00:01", 0.0333),
])
def test_duration_hours(start, end, hours):
    assert duration_hours(parse_time(start), parse_time(end)) == pytest.approx(hours, abs=1e-3)


def test_duration_rejects_equal_start_and_finish():
    with pytest.raises(TimeError, match="starts and finishes"):
        duration_hours(time(8, 0), time(8, 0))


@pytest.mark.parametrize("start,end,expected", [
    ("21:00", "07:00", True),
    ("08:00", "16:00", False),
    ("13:00", "21:00", False),
    ("23:00", "23:30", False),
])
def test_crosses_midnight(start, end, expected):
    assert crosses_midnight(parse_time(start), parse_time(end)) is expected


def test_a_night_shift_totals_the_same_however_it_is_measured():
    """Twenty 10-hour nights must not be mistaken for twenty 7.5-hour days."""
    night = duration_hours(parse_time("21:00"), parse_time("07:00"))
    day = duration_hours(parse_time("09:00"), parse_time("16:30"))
    assert night * 20 == pytest.approx(200.0)
    assert day * 20 == pytest.approx(150.0)
    assert night * 20 != day * 20


# --- shift windows and rest ------------------------------------------------

def test_shift_window_rolls_finish_onto_the_next_day():
    begins, finishes = shift_window(date(2026, 9, 1), time(21, 0), time(7, 0))
    assert begins == datetime(2026, 9, 1, 21, 0)
    assert finishes == datetime(2026, 9, 2, 7, 0)


def test_shift_window_stays_on_the_same_day_when_it_should():
    begins, finishes = shift_window(date(2026, 9, 1), time(9, 0), time(17, 30))
    assert begins == datetime(2026, 9, 1, 9, 0)
    assert finishes == datetime(2026, 9, 1, 17, 30)


def test_late_then_early_is_the_ten_hour_case_from_the_specification():
    """Late 13:00-21:00 followed by early 07:00-15:00 leaves only 10 hours."""
    _, late_end = shift_window(date(2026, 9, 1), time(13, 0), time(21, 0))
    early_start, _ = shift_window(date(2026, 9, 2), time(7, 0), time(15, 0))
    assert rest_hours(late_end, early_start) == pytest.approx(10.0)


def test_night_then_next_day_shift_leaves_very_little_rest():
    _, night_end = shift_window(date(2026, 9, 1), time(21, 0), time(7, 0))
    early_start, _ = shift_window(date(2026, 9, 2), time(9, 0), time(17, 0))
    assert rest_hours(night_end, early_start) == pytest.approx(2.0)


def test_rest_is_ample_after_a_day_off():
    _, end = shift_window(date(2026, 9, 1), time(9, 0), time(17, 0))
    start, _ = shift_window(date(2026, 9, 3), time(9, 0), time(17, 0))
    assert rest_hours(end, start) == pytest.approx(40.0)


def test_rest_goes_negative_when_shifts_overlap():
    _, end = shift_window(date(2026, 9, 1), time(21, 0), time(7, 0))
    start, _ = shift_window(date(2026, 9, 2), time(6, 0), time(14, 0))
    assert rest_hours(end, start) == pytest.approx(-1.0)


# --- overlap primitive for future part-shift bench blocks ------------------

def test_overlapping_and_touching_periods():
    morning = shift_window(date(2026, 9, 1), time(9, 0), time(13, 0))
    afternoon = shift_window(date(2026, 9, 1), time(13, 0), time(17, 0))
    straddling = shift_window(date(2026, 9, 1), time(12, 0), time(14, 0))

    assert overlaps(*morning, *afternoon) is False   # back to back, not overlapping
    assert overlaps(*morning, *straddling) is True
    assert overlaps(*afternoon, *straddling) is True


# --- repeating week patterns ----------------------------------------------

def test_week_index_counts_from_the_anchor_week_monday():
    anchor = date(2026, 9, 2)                 # a Wednesday
    assert week_index(date(2026, 8, 31), anchor) == 0   # Monday of that week
    assert week_index(anchor, anchor) == 0
    assert week_index(date(2026, 9, 6), anchor) == 0    # Sunday, same week
    assert week_index(date(2026, 9, 7), anchor) == 1    # next Monday


def test_alternating_two_week_pattern_repeats():
    anchor = date(2026, 9, 1)
    mondays = [date(2026, 8, 31), date(2026, 9, 7),
               date(2026, 9, 14), date(2026, 9, 21)]
    assert [pattern_week(day, anchor, 2) for day in mondays] == [1, 2, 1, 2]


def test_three_week_pattern_repeats():
    anchor = date(2026, 9, 7)
    weeks = [anchor + timedelta(weeks=offset) for offset in range(6)]
    assert [pattern_week(day, anchor, 3) for day in weeks] == [1, 2, 3, 1, 2, 3]


def test_pattern_cycle_must_be_at_least_one_week():
    with pytest.raises(TimeError):
        pattern_week(date(2026, 9, 1), date(2026, 9, 1), 0)
