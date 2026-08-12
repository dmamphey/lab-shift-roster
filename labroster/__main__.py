"""Command line for LabRoster, for developers and batch use.

    python -m labroster template --out blank.xlsx
    python -m labroster example  --out example-laboratory.xlsx
    python -m labroster generate --input my-workbook.xlsx --out draft.xlsx

Laboratory managers use the browser interface; nothing here appears in the
manager-facing workbooks.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import api


def _write(path: Path, data: bytes) -> int:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        path.write_bytes(data)
    except PermissionError:
        print(f"Could not write {path}: the file is open in another program, "
              f"most likely Excel.\nClose it and run this again.")
        return 1
    print(f"Written to {path}")
    return 0


def cmd_template(args) -> int:
    return _write(Path(args.out), api.blank_template_bytes())


def cmd_example(args) -> int:
    return _write(Path(args.out), api.demo_workbook_bytes())


def cmd_generate(args) -> int:
    source = Path(args.input)
    if not source.exists():
        print(f"Workbook not found: {source}")
        return 1

    result = api.generate(source.read_bytes(), start=args.start, end=args.end,
                          alternative=args.variation)
    if not result["ok"]:
        print(result.get("fatal", "The workbook could not be used."))
        for problem in result.get("problems", []):
            where = f"{problem['location']}: " if problem["location"] else ""
            print(f"  [{problem['severity']}] {where}{problem['message']}")
        return 1

    for problem in result["problems"]:
        print(f"  [{problem['severity']}] {problem['location']}: "
              f"{problem['message']}")

    dashboard = result["dashboard"]
    print(f"\n{result['details']['heading']}")
    print(f"{result['details']['period_label']}  "
          f"({dashboard['day_count']} days, {dashboard['staff_count']} staff)")
    print(f"\nRoster status      : {dashboard['roster_status']}")
    print(f"Shift coverage     : {dashboard['shift_coverage_percent']}%")
    for label, key in [
        ("Unfilled shifts", "unfilled_shifts"),
        ("Uncovered sections", "uncovered_benches"),
        ("Senior cover gaps", "senior_cover_gaps"),
        ("Competency gaps", "competency_gaps"),
        ("Rest rule conflicts", "rest_conflicts"),
        ("Outside target hours", "staff_outside_target_hours"),
        ("Expiring competencies", "competencies_expiring_soon"),
        ("Expired competencies", "competencies_expired"),
        ("Single points of failure", "single_points_of_failure"),
    ]:
        print(f"{label:<19}: {dashboard[key]}")
    print(f"Weekend fairness   : {dashboard['weekend_fairness']}")
    print(f"Night fairness     : {dashboard['night_fairness']}")
    print(f"\nIssues: {dashboard['critical_count']} critical, "
          f"{dashboard['review_count']} to review, "
          f"{dashboard['passed_count']} passed")
    print("This roster is a draft requiring managerial review.\n")

    return _write(Path(args.out), result["workbook"])


def main(argv=None) -> int:
    # Windows consoles often default to a codepage that cannot show the
    # separators used in headings, which turns them into question marks.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):    # pragma: no cover - old streams
            pass

    parser = argparse.ArgumentParser(
        prog="python -m labroster",
        description="LabRoster — competency-aware workforce planning for "
                    "diagnostic laboratories.")
    sub = parser.add_subparsers(dest="command", required=True)

    blank = sub.add_parser("template", help="write a blank workbook to fill in")
    blank.add_argument("--out", default="LabRoster-blank-template.xlsx")
    blank.set_defaults(func=cmd_template)

    example = sub.add_parser("example",
                             help="write the fictional example laboratory")
    example.add_argument("--out", default="LabRoster-example-laboratory.xlsx")
    example.set_defaults(func=cmd_example)

    generate = sub.add_parser("generate", help="build a draft roster")
    generate.add_argument("--input", required=True)
    generate.add_argument("--out", default="LabRoster-draft.xlsx")
    generate.add_argument("--start", help="override the period start (yyyy-mm-dd)")
    generate.add_argument("--end", help="override the period end (yyyy-mm-dd)")
    generate.add_argument("--variation", type=int,
                          help="different draft from the same information")
    generate.set_defaults(func=cmd_generate)

    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
