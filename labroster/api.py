"""One entry point, used by both the browser and the command line.

The browser calls :func:`generate` with the bytes of an uploaded workbook and
gets back a plain dictionary: the dashboard, the issues, a roster grid it can
draw, and the bytes of the exported workbook.  Nothing in here touches the
network or the filesystem, which is what keeps the workforce data on the user's
own device.
"""

from __future__ import annotations

import io
from datetime import date, datetime

from .analysis import CRITICAL, PASSED, REVIEW, Analysis
from .export import write_workbook
from .models import Config
from .scheduler import Assignment, Scheduler
from .template import build_blank_template, build_demo_workbook
from .workbook import ERROR, Problem, WorkbookError, read_workbook


def blank_template_bytes() -> bytes:
    buffer = io.BytesIO()
    build_blank_template(buffer)
    return buffer.getvalue()


def demo_workbook_bytes() -> bytes:
    buffer = io.BytesIO()
    build_demo_workbook(buffer)
    return buffer.getvalue()


def _problem_payload(problems: list[Problem]) -> list[dict]:
    return [{"severity": problem.severity, "sheet": problem.sheet,
             "row": problem.row, "message": problem.message,
             "location": problem.location}
            for problem in problems]


def _roster_payload(scheduler: Scheduler, analysis: Analysis) -> dict:
    """Everything the browser preview needs to draw the roster."""
    config = scheduler.config
    shift_by_code = config.shift_by_code()
    leave_types = config.leave_types

    days = [{
        "iso": day.isoformat(),
        "day": day.day,
        "weekday": day.strftime("%a"),
        "weekday_full": day.strftime("%A"),
        "month": day.strftime("%b %Y"),
        "is_weekend": scheduler.is_weekend(day),
    } for day in scheduler.days]

    rows = []
    for person in config.staff:
        cells = {}
        for day in scheduler.days:
            assignment = scheduler.assignments.get((day, person.staff_id))
            if assignment:
                shift = shift_by_code.get(assignment.shift_code)
                cells[day.isoformat()] = {
                    "kind": "shift",
                    "code": assignment.shift_code,
                    "label": shift.name if shift else assignment.shift_code,
                    "times": shift.times_label if shift else "",
                    "hours": round(shift.hours, 2) if shift else 0,
                    "colour": shift.colour if shift else "D9E1F2",
                    "font": shift.font_colour if shift else "000000",
                    "source": assignment.source,
                    "benches": [allocation.bench_name
                                for allocation in scheduler.bench_allocations
                                if allocation.day == day
                                and allocation.staff_id == person.staff_id],
                }
                continue
            code = scheduler.leave.get((day, person.staff_id))
            if code:
                entry = leave_types.get(
                    "".join(ch for ch in code.lower() if ch.isalnum()))
                cells[day.isoformat()] = {
                    "kind": "leave", "code": code,
                    "label": entry.label if entry else code,
                    "colour": entry.colour if entry else "FFFF00",
                    "font": entry.font_colour if entry else "000000",
                }
            else:
                cells[day.isoformat()] = {"kind": "off", "code": "",
                                          "label": "Not working"}

        rows.append({
            "staff_id": person.staff_id,
            "name": person.name,
            "band": person.band,
            "job_title": person.job_title,
            "group": person.group,
            "senior": person.is_senior,
            "coordinator": person.shift_coordinator,
            "cells": cells,
        })

    # Bench coverage per day and shift, from the actual allocations.
    coverage = []
    for day in scheduler.days:
        for shift in config.shifts:
            if not shift.applies_on(day, scheduler.rules.weekend_days):
                continue
            benches = scheduler.benches_for(day, shift)
            if not benches:
                continue
            for bench in benches:
                wanted = bench.required_on(day, scheduler.rules.weekend_days)
                allocated = scheduler.bench_staff(day, bench.name, shift.code)
                if not wanted and not allocated:
                    continue
                if len(allocated) >= wanted and wanted:
                    state = "Covered"
                elif allocated:
                    state = "Review"
                else:
                    state = "Not covered"
                coverage.append({
                    "date": day.isoformat(),
                    "shift": shift.code,
                    "shift_name": shift.name,
                    "bench": bench.name,
                    "discipline": bench.discipline,
                    "required": wanted,
                    "allocated": len(allocated),
                    "state": state,
                    "staff": [analysis.name_of(sid) for sid in allocated],
                })

    return {
        "days": days,
        "rows": rows,
        "coverage": coverage,
        "shifts": [{"code": shift.code, "name": shift.name,
                    "times": shift.times_label, "colour": shift.colour,
                    "font": shift.font_colour, "night": shift.is_night,
                    "hours": round(shift.hours, 2)}
                   for shift in config.shifts],
        "leave_types": [{"code": entry.code, "label": entry.label,
                         "colour": entry.colour, "font": entry.font_colour}
                        for entry in config.leave_types.values()],
        "benches": [{"name": bench.name, "discipline": bench.discipline}
                    for bench in config.benches],
    }


def _details_payload(config: Config) -> dict:
    return {
        "rota_name": config.details.rota_name,
        "organisation": config.details.organisation,
        "department": config.details.department,
        "site": config.details.site,
        "prepared_by": config.details.prepared_by,
        "heading": config.details.heading,
        "period_start": config.period.start.isoformat(),
        "period_end": config.period.end.isoformat(),
        "period_label": (f"{config.period.start:%d %b %Y} to "
                         f"{config.period.end:%d %b %Y}"),
        "generated": datetime.now().strftime("%d %b %Y at %H:%M"),
        "minimum_rest_hours": config.rules.minimum_rest_hours,
        "hours_tolerance_percent": config.rules.hours_tolerance_percent,
    }


def generate(data: bytes, start: str | None = None, end: str | None = None,
             alternative: int | None = None,
             manual: list[dict] | None = None) -> dict:
    """Read a workbook, build a draft roster and report on it.

    Returns a dictionary with ``ok`` set to False and a list of problems when the
    workbook cannot be used, so the caller can show every correction needed at
    once rather than one at a time.
    """
    try:
        config, problems = read_workbook(io.BytesIO(data))
    except WorkbookError as error:
        return {"ok": False, "fatal": str(error),
                "problems": _problem_payload(error.problems)}

    # Optional overrides from the interface.
    if start:
        parsed = date.fromisoformat(start)
        config.period.start = parsed
    if end:
        parsed = date.fromisoformat(end)
        config.period.end = parsed
    if config.period.end < config.period.start:
        return {"ok": False,
                "fatal": "The end date is before the start date.",
                "problems": _problem_payload(problems)}
    if alternative is not None:
        config.rules.seed = int(alternative)

    blocking = [problem for problem in problems if problem.severity == ERROR]
    if blocking:
        return {"ok": False,
                "fatal": ("The workbook needs correcting before a roster can be "
                          "produced. Everything found is listed below so you can "
                          "fix it in one pass."),
                "problems": _problem_payload(problems)}

    manual_assignments = [
        Assignment(day=date.fromisoformat(item["date"]),
                   staff_id=item["staff_id"],
                   shift_code=item["shift_code"], source="manual")
        for item in (manual or [])
    ]

    scheduler = Scheduler(config, manual_assignments=manual_assignments)
    scheduler.build()
    analysis = Analysis(scheduler)

    buffer = io.BytesIO()
    write_workbook(scheduler, analysis, buffer)

    return {
        "ok": True,
        "problems": _problem_payload(problems),
        "details": _details_payload(config),
        "dashboard": analysis.metrics,
        "issues": analysis.issues_payload(),
        "hours": analysis.hours_payload(),
        "roster": _roster_payload(scheduler, analysis),
        "resilience": [{"discipline": item.discipline,
                        "competent": item.competent_count,
                        "authorisers": item.authoriser_count,
                        "severity": item.severity,
                        "message": item.message}
                       for item in analysis.workforce_resilience],
        "expiring": [{"name": item["name"], "discipline": item["discipline"],
                      "competency": item["name_of_competency"],
                      "expiry": item["expiry"].isoformat() if item["expiry"] else None,
                      "days": item["days"], "state": item["state"]}
                     for item in analysis.expiring],
        "workbook": buffer.getvalue(),
        "severities": {"critical": CRITICAL, "review": REVIEW, "passed": PASSED},
    }
