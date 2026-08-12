"""Building the workbook a manager fills in.

Two workbooks come out of here:

* a **blank template** — headings, instructions and dropdown lists, with no
  employee records at all
* an **example laboratory** — a complete fictional department, deliberately built
  so that every kind of warning the tool can raise is triggered by it

Every name, competency and date in the example is invented.  Real staff details
must never ship inside a template: it is distributed with the product, and it
would carry personal data into every copy.

Instructions written into the workbook are for laboratory managers.  Nothing
here mentions Python, command lines or file paths.
"""

from __future__ import annotations

from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import SCHEMA_VERSION
from .models import CompetencyStatus, WEEKDAY_SHORT

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
SECTION_FILL = PatternFill("solid", fgColor="2E5496")
NOTE_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WHITE_BOLD = Font(bold=True, color="FFFFFF", size=10)
WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")

MANAGER_STEPS = [
    "Complete the Staff sheet: one row per member of staff.",
    "Complete the Competencies sheet: one row per person per discipline.",
    "Complete the Shifts sheet: the shifts your laboratory runs, with times.",
    "Add leave and availability: the Leave sheet, and Week Patterns if anyone "
    "works alternating weeks.",
    "Configure requirements: the Shift Requirements and Benches sheets set the "
    "minimum staffing and competencies each shift needs.",
    "Check the Roster Details and Rules sheets.",
    "Save the workbook.",
    "Upload it to LabRoster in your browser.",
    "Generate the draft roster.",
    "Review the warnings LabRoster raises.",
    "Export the finished workbook.",
]

INSTRUCTION_NOTES = [
    ("How to use this workbook", True),
    ("", False),
] + [(f"{index}.  {text}", False)
     for index, text in enumerate(MANAGER_STEPS, start=1)] + [
    ("", False),
    ("Filling in the sheets", True),
    ("Staff — Staff ID must be unique and is what the other sheets refer to. "
     "Contracted Weekly Hours drives how much work each person is given.", False),
    ("Staff — the Mon to Sun columns are the days that person normally works. "
     "Leave them all blank if somebody is fully flexible.", False),
    ("Staff — Senior, Shift Coordinator and Registered are recorded separately "
     "from Band, because they are different things. A Band 7 is not "
     "automatically competent in every discipline.", False),
    ("Competencies — only Competent, Trainer and Assessor count as cover. "
     "In Training and Supervised do not, because those staff need somebody with "
     "them. Expired competencies never count.", False),
    ("Shift Requirements — Required Competencies is written as discipline and "
     "number, for example: BT:1, HAEM:2, COAG:1", False),
    ("Benches — each bench needs its own person. One member of staff is not "
     "counted as covering two benches at the same time.", False),
    ("Rules — these are your laboratory's own operating rules, including the "
     "rest interval. They are settings for planning, not a determination of "
     "legal or regulatory compliance.", False),
    ("", False),
    ("What LabRoster produces", True),
    ("A draft roster for you to review, together with a list of staffing, "
     "competency and coverage gaps. Every roster it generates is a draft "
     "requiring managerial review.", False),
]


def _sheet(workbook, title, headers, widths=None, note=None, freeze=True):
    """Create a sheet with a title row, a header row and sensible widths."""
    sheet = workbook.create_sheet(title)
    sheet.sheet_view.showGridLines = False
    row = 1
    if note:
        sheet.cell(row=1, column=1, value=note).font = Font(bold=True, size=10,
                                                            color="7F4F00")
        sheet.cell(row=1, column=1).fill = NOTE_FILL
        sheet.merge_cells(start_row=1, start_column=1,
                          end_row=1, end_column=max(len(headers), 4))
        row = 2
    for column, heading in enumerate(headers, start=1):
        cell = sheet.cell(row=row, column=column, value=heading)
        cell.fill = HEADER_FILL
        cell.font = WHITE_BOLD
        cell.alignment = WRAP
        cell.border = BORDER
    sheet.row_dimensions[row].height = 34
    for column, width in enumerate(widths or [], start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    if freeze:
        sheet.freeze_panes = sheet.cell(row=row + 1, column=1).coordinate
    sheet.auto_filter.ref = (f"A{row}:"
                             f"{get_column_letter(len(headers))}{row}")
    return sheet, row


def _write_rows(sheet, header_row, rows):
    for offset, values in enumerate(rows):
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row=header_row + 1 + offset, column=column,
                              value=value)
            cell.border = BORDER
            cell.alignment = LEFT
            if isinstance(value, date):
                cell.number_format = "DD/MM/YYYY"


def _yes_no(sheet, header_row, columns, last_row=400):
    validation = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    sheet.add_data_validation(validation)
    for column in columns:
        letter = get_column_letter(column)
        validation.add(f"{letter}{header_row + 1}:{letter}{last_row}")


def _list_validation(sheet, header_row, column, options, last_row=400):
    formula = '"' + ",".join(options) + '"'
    validation = DataValidation(type="list", formula1=formula, allow_blank=True)
    sheet.add_data_validation(validation)
    letter = get_column_letter(column)
    validation.add(f"{letter}{header_row + 1}:{letter}{last_row}")


# --------------------------------------------------------------------------
# column definitions, shared by both workbooks
# --------------------------------------------------------------------------

STAFF_HEADERS = [
    "Staff ID", "Name", "Job Title", "Band", "Registered BMS", "Senior",
    "Shift Coordinator", "Trainee", "Contracted Weekly Hours", "FTE",
    "Working Pattern", "Pattern Cycle (weeks)",
    *WEEKDAY_SHORT,
    "Earliest Start", "Latest Finish", "Max Days Per Week",
    "Max Consecutive Days", "Works Nights", "Works Weekends", "Max Nights",
    "Max Weekends", "Max Hours This Period", "Group", "Restrictions", "Notes",
]
STAFF_WIDTHS = [10, 20, 20, 7, 11, 8, 12, 8, 13, 6, 16, 11,
                5, 5, 5, 5, 5, 5, 5, 11, 11, 11, 12, 10, 11, 10, 11, 13, 10, 20, 20]

COMPETENCY_HEADERS = [
    "Staff ID", "Name", "Discipline", "Competency", "Status", "Date Achieved",
    "Review Date", "Expiry Date", "Trainer", "Assessor", "Result Authoriser",
    "Notes",
]
COMPETENCY_WIDTHS = [10, 20, 11, 26, 13, 13, 13, 13, 9, 9, 14, 22]

SHIFT_HEADERS = ["Code", "Name", "Start", "End", "Days", "Night Shift",
                 "Colour", "Font Colour"]
SHIFT_WIDTHS = [8, 16, 9, 9, 12, 11, 10, 11]

REQUIREMENT_HEADERS = [
    "Shift Code", "Days", "Min Staff", "Min Registered BMS", "Min Senior",
    "Min Band", "Staff At Min Band", "Min Coordinators", "Min Trainers",
    "Max Trainees", "Required Competencies", "Required Authorisers", "Notes",
]
REQUIREMENT_WIDTHS = [10, 12, 10, 12, 10, 9, 12, 13, 11, 11, 24, 22, 20]

BENCH_HEADERS = ["Bench", "Discipline", "Days", "Min Staff",
                 "Min Staff Weekend", "Shift Codes", "Requires Authoriser",
                 "Rotation Interval (days)"]
BENCH_WIDTHS = [24, 12, 12, 10, 16, 14, 16, 18]

LEAVE_HEADERS = ["Staff ID", "Name", "From", "To", "Type", "Reason"]
LEAVE_WIDTHS = [10, 20, 13, 13, 10, 24]

PATTERN_HEADERS = ["Staff ID", "Week", *WEEKDAY_SHORT]
PATTERN_WIDTHS = [10, 8, 5, 5, 5, 5, 5, 5, 5]

LEAVE_TYPE_HEADERS = ["Code", "Label", "Colour", "Font Colour"]
LEAVE_TYPE_WIDTHS = [10, 34, 10, 11]

DEFAULT_SHIFTS = [
    ("E", "Early", "07:00", "15:00", "Weekday", "N", "FFF2CC", "000000"),
    ("C", "Core", "09:00", "17:30", "Weekday", "N", "D9E1F2", "000000"),
    ("L", "Late", "13:00", "21:00", "Weekday", "N", "00B0F0", "000000"),
    ("N", "Night", "21:00", "07:00", "All", "Y", "7030A0", "FFFFFF"),
    ("W", "Weekend Day", "09:00", "17:30", "Weekend", "N", "00B050", "FFFFFF"),
]

DEFAULT_LEAVE_TYPES = [
    ("A/L", "Annual leave", "FFFF00", "000000"),
    ("S/L", "Sickness absence", "FF9999", "000000"),
    ("C/L", "Carers or compassionate leave", "FFC000", "000000"),
    ("M/L", "Maternity or paternity leave", "D9D9D9", "000000"),
    ("S/D", "Study or training day", "CC99FF", "000000"),
]

#: Shift Codes says which shifts a section must be staffed during.  Transfusion
#: and haematology run whenever the laboratory is open; coagulation and
#: morphology only during the main day service.
DEFAULT_BENCHES = [
    ("Blood Transfusion", "BT", "All", 1, 1, "C, L, W", "Y", 42),
    ("Haematology", "HAEM", "All", 1, 1, "C, L, W", "N", 42),
    ("Coagulation", "COAG", "Weekday", 1, 0, "C", "N", 56),
    ("Morphology", "MORPH", "Weekday", 1, 0, "C", "N", 56),
]

# Sized so the example workforce can genuinely deliver it, while still leaving
# the deliberate morphology gap that demonstrates a single point of failure.
DEFAULT_REQUIREMENTS = [
    ("E", "Weekday", 2, 1, 0, 0, 0, 0, 0, 1, "", "", ""),
    ("C", "Weekday", 5, 3, 1, 6, 1, 1, 0, 2,
     "BT:1, HAEM:1, COAG:1, MORPH:1", "BT:1", "Main day service"),
    ("L", "Weekday", 2, 1, 1, 0, 0, 0, 0, 1, "BT:1", "BT:1", ""),
    ("N", "All", 1, 1, 1, 0, 0, 0, 0, 0, "BT:1, HAEM:1", "BT:1",
     "Lone working out of hours"),
    ("W", "Weekend", 3, 2, 1, 0, 0, 1, 0, 1, "BT:1, HAEM:1", "BT:1", ""),
]

RULES_ROWS = [
    ("Senior band threshold", 6,
     "Band at or above which somebody counts as senior, if you use band for this."),
    ("Minimum rest hours between shifts", 11,
     "Your laboratory's own rest rule. Not a legal compliance determination."),
    ("Maximum consecutive days", 6, "Longest unbroken run of working days."),
    ("Maximum consecutive nights", 4, "Longest unbroken run of night shifts."),
    ("Night block length", 3, "Nights are rostered in blocks of this length."),
    ("Recovery days after nights", 2, "Days off after a block of nights."),
    ("Hours tolerance (%)", 10,
     "How far from contracted hours before somebody is flagged for review."),
    ("Competency expiry warnings (days)", "30, 60, 90",
     "Warn when a competency expires within these periods."),
    ("Max simultaneous bench assignments", 1,
     "How many benches one person may cover at the same time. Normally 1."),
    ("Cross cover allowed", "N",
     "Reserved for future use: allowing one person to cover two sections."),
    ("Weekend days", "Saturday, Sunday", "Which days count as the weekend."),
    ("Section rotation warning (days)", 56,
     "Warn if somebody has not worked a section for this long."),
    ("Alternative roster number", 42,
     "Change this to generate a different draft from the same information."),
]


def _instructions_sheet(workbook, is_demo: bool) -> None:
    sheet = workbook.active
    sheet.title = "Instructions"
    sheet.sheet_view.showGridLines = False
    sheet.column_dimensions["A"].width = 104

    sheet["A1"] = "LabRoster"
    sheet["A1"].font = Font(bold=True, size=18, color="1F3864")
    sheet["A2"] = "Competency-aware workforce planning for diagnostic laboratories"
    sheet["A2"].font = Font(size=11, italic=True, color="5B6470")

    row = 4
    if is_demo:
        sheet.cell(row=row, column=1,
                   value="EXAMPLE LABORATORY — every member of staff, competency, "
                         "date and requirement in this workbook is fictional. "
                         "Replace all of it with your own information.")
        sheet.cell(row=row, column=1).font = Font(bold=True, size=11,
                                                  color="9C0006")
        sheet.cell(row=row, column=1).fill = PatternFill("solid", fgColor="FBE6E8")
        sheet.cell(row=row, column=1).alignment = Alignment(wrap_text=True,
                                                           vertical="center")
        sheet.row_dimensions[row].height = 44
        row += 2
    else:
        sheet.cell(row=row, column=1,
                   value="BLANK TEMPLATE — the Staff, Competencies, Leave and "
                         "Week Patterns sheets are empty and ready for your own "
                         "information. The Shifts, Shift Requirements and Benches "
                         "sheets contain a common starting configuration for you "
                         "to adjust.")
        sheet.cell(row=row, column=1).font = Font(bold=True, size=10,
                                                  color="1D6B3F")
        sheet.cell(row=row, column=1).fill = PatternFill("solid", fgColor="E3F4E9")
        sheet.cell(row=row, column=1).alignment = Alignment(wrap_text=True,
                                                           vertical="center")
        sheet.row_dimensions[row].height = 44
        row += 2

    for text, is_heading in INSTRUCTION_NOTES:
        cell = sheet.cell(row=row, column=1, value=text)
        if is_heading:
            cell.font = Font(bold=True, size=12, color="1F3864")
        else:
            cell.font = Font(size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if len(text) > 90:
                sheet.row_dimensions[row].height = 28
        row += 1

    row += 1
    sheet.cell(row=row, column=1,
               value="Your workforce information is processed on your own device, "
                     "in your browser. It is not uploaded to Optymum SS. Please do "
                     "not add sensitive personal details or any patient "
                     "information: this workbook does not need them.")
    sheet.cell(row=row, column=1).font = Font(size=10, italic=True,
                                              color="1D6B3F")
    sheet.cell(row=row, column=1).alignment = Alignment(wrap_text=True)
    sheet.row_dimensions[row].height = 40


def _details_sheet(workbook, period_start: date, period_end: date,
                   is_demo: bool) -> None:
    sheet, header_row = _sheet(workbook, "Roster Details", ["Detail", "Value"],
                               widths=[30, 40], freeze=False)
    rows = [
        ("Rota name", "Laboratory Staff Rota"),
        ("Organisation", "Example NHS Trust" if is_demo else ""),
        ("Department", "Blood Sciences" if is_demo else ""),
        ("Site", "Example Hospital" if is_demo else ""),
        ("Prepared by", "A. Manager (example)" if is_demo else ""),
        ("Roster period start", period_start),
        ("Roster period end", period_end),
        ("Workbook version", SCHEMA_VERSION),
    ]
    _write_rows(sheet, header_row, rows)
    for offset in range(len(rows)):
        sheet.cell(row=header_row + 1 + offset, column=1).font = Font(bold=True,
                                                                      size=10)


def _rules_sheet(workbook) -> None:
    sheet, header_row = _sheet(workbook, "Rules",
                               ["Setting", "Value", "What it means"],
                               widths=[36, 18, 68], freeze=False)
    _write_rows(sheet, header_row, RULES_ROWS)
    for offset in range(len(RULES_ROWS)):
        sheet.cell(row=header_row + 1 + offset, column=1).font = Font(bold=True,
                                                                      size=10)
        sheet.cell(row=header_row + 1 + offset, column=3).font = Font(size=9,
                                                                      color="5B6470")
        sheet.cell(row=header_row + 1 + offset, column=3).alignment = \
            Alignment(wrap_text=True, vertical="center")


def _configuration_sheets(workbook) -> None:
    """Shifts, requirements, benches and leave types: the same in both workbooks."""
    sheet, header_row = _sheet(workbook, "Shifts", SHIFT_HEADERS, SHIFT_WIDTHS)
    _write_rows(sheet, header_row, DEFAULT_SHIFTS)
    _yes_no(sheet, header_row, [6], last_row=60)
    _list_validation(sheet, header_row, 5,
                     ["All", "Weekday", "Weekend"], last_row=60)

    sheet, header_row = _sheet(
        workbook, "Shift Requirements", REQUIREMENT_HEADERS, REQUIREMENT_WIDTHS,
        note="Required Competencies and Required Authorisers are written as "
             "discipline and number, for example:  BT:1, HAEM:2, COAG:1")
    _write_rows(sheet, header_row, DEFAULT_REQUIREMENTS)
    _list_validation(sheet, header_row, 2,
                     ["All", "Weekday", "Weekend", "Monday", "Tuesday",
                      "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
                     last_row=100)

    sheet, header_row = _sheet(
        workbook, "Benches", BENCH_HEADERS, BENCH_WIDTHS,
        note="Each bench needs its own person. One member of staff is never "
             "counted as covering two benches at the same time. Shift Codes says "
             "which shifts the section must be staffed during — leave it blank to "
             "mean every shift except nights.")
    _write_rows(sheet, header_row, DEFAULT_BENCHES)
    _yes_no(sheet, header_row, [7], last_row=60)
    _list_validation(sheet, header_row, 3, ["All", "Weekday", "Weekend"],
                     last_row=60)

    sheet, header_row = _sheet(workbook, "Leave Types", LEAVE_TYPE_HEADERS,
                               LEAVE_TYPE_WIDTHS)
    _write_rows(sheet, header_row, DEFAULT_LEAVE_TYPES)


def _empty_people_sheets(workbook) -> None:
    sheet, header_row = _sheet(
        workbook, "Staff", STAFF_HEADERS, STAFF_WIDTHS,
        note="One row per member of staff. Staff ID must be unique — the other "
             "sheets refer to it. Leave the Mon–Sun columns blank for somebody "
             "who is fully flexible.")
    _yes_no(sheet, header_row, [5, 6, 7, 8, *range(13, 20), 24, 25])

    sheet, header_row = _sheet(
        workbook, "Competencies", COMPETENCY_HEADERS, COMPETENCY_WIDTHS,
        note="Only Competent, Trainer and Assessor count as coverage. "
             "In Training and Supervised do not. Expired competencies never count.")
    _list_validation(sheet, header_row, 5, CompetencyStatus.ALL)
    _yes_no(sheet, header_row, [9, 10, 11])

    sheet, header_row = _sheet(workbook, "Week Patterns", PATTERN_HEADERS,
                               PATTERN_WIDTHS,
                               note="Only needed for staff who work alternating "
                                    "weeks. One row per person per week.")
    _yes_no(sheet, header_row, list(range(3, 10)))

    sheet, header_row = _sheet(workbook, "Leave", LEAVE_HEADERS, LEAVE_WIDTHS,
                               note="From and To are both included in the "
                                    "absence.")
    _list_validation(sheet, header_row, 5,
                     [code for code, *_ in DEFAULT_LEAVE_TYPES])


def build_blank_template(path) -> None:
    """A workbook with headings, instructions and dropdowns, and no staff."""
    workbook = Workbook()
    today = date.today()
    first = date(today.year + (today.month == 12), (today.month % 12) + 1, 1)
    last = (date(first.year + (first.month == 12), (first.month % 12) + 1, 1)
            - timedelta(days=1))

    _instructions_sheet(workbook, is_demo=False)
    _details_sheet(workbook, first, last, is_demo=False)
    _rules_sheet(workbook)
    _empty_people_sheets(workbook)
    _configuration_sheets(workbook)
    workbook.save(path)


# --------------------------------------------------------------------------
# the example laboratory
# --------------------------------------------------------------------------

def _demo_staff(period_start: date):
    """A fictional department, built to exercise every warning.

    Deliberate features: part-time staff with fixed non-working days, an
    alternating two-week pattern, staff who cannot work nights or weekends, a
    trainee, a single morphology-competent scientist (a workforce single point of
    failure), an expired competency and one expiring shortly.
    """
    Y, N = "Y", "N"
    # id, name, title, band, reg, senior, coord, trainee, hours, fte, pattern,
    # cycle, Mon..Sun, earliest, latest, maxdays, maxconsec, nights, weekends,
    # maxnights, maxweekends, maxhours, group, restrictions, notes
    return [
        ("S01", "Alex Sample", "Laboratory Manager", "8a", Y, Y, Y, N, 37.5, 1.0,
         "Full time", 1, Y, Y, Y, Y, Y, N, N, "", "", "", "", N, N, "", "", 0,
         "Main", "Management time; no nights", ""),
        ("S02", "Jordan Test", "Senior BMS", "7", Y, Y, Y, N, 37.5, 1.0,
         "Full time", 1, Y, Y, Y, Y, Y, Y, Y, "", "", "", "", Y, Y, 0, 0, 0,
         "Main", "", ""),
        ("S03", "Casey Example", "Senior BMS", "6", Y, Y, Y, N, 37.5, 1.0,
         "Full time", 1, Y, Y, Y, Y, Y, Y, Y, "", "", "", "", Y, Y, 0, 0, 0,
         "Main", "", ""),
        ("S04", "Morgan Demo", "BMS", "6", Y, Y, N, N, 30.0, 0.8,
         "Part time", 1, Y, Y, N, Y, Y, Y, Y, "", "", 4, "", Y, Y, 0, 2, 0,
         "Main", "No Wednesdays; maximum 2 weekends", ""),
        ("S05", "Riley Placeholder", "BMS", "6", Y, Y, N, N, 22.5, 0.6,
         "Part time", 1, N, Y, Y, Y, N, N, N, "", "", 3, "", N, N, 0, 0, 0,
         "Main", "No Mondays or Fridays; no nights or weekends", ""),
        ("S06", "Jamie Mock", "BMS", "5", Y, N, N, N, 22.5, 0.6,
         "Alternating weeks", 2, "", "", "", "", "", "", "", "", "", "", "",
         Y, Y, 0, 0, 0, "Main",
         "Alternating two-week pattern, three days each week", ""),
        ("S07", "Taylor Dummy", "BMS", "5", Y, N, N, N, 37.5, 1.0,
         "Full time", 1, Y, Y, Y, Y, Y, Y, Y, "", "", "", "", Y, Y, 0, 0, 0,
         "Main", "", ""),
        ("S08", "Avery Specimen", "BMS", "5", Y, N, N, N, 37.5, 1.0,
         "Full time", 1, Y, Y, Y, Y, Y, Y, Y, "", "", "", "", Y, Y, 0, 0, 0,
         "Main", "", ""),
        ("S09", "Quinn Template", "BMS", "5", Y, N, N, N, 18.75, 0.5,
         "Part time", 1, N, N, Y, Y, Y, N, N, "", "16:00", 3, "", N, N, 0, 0, 0,
         "Main", "Wednesday to Friday, finishing by 16:00", ""),
        ("S10", "Rowan Draft", "BMS", "5", Y, N, N, N, 37.5, 1.0,
         "Full time", 1, Y, Y, Y, Y, Y, Y, Y, "", "", "", "", Y, Y, 0, 0, 0,
         "Main", "", ""),
        ("S11", "Skyler Trial", "Associate Practitioner", "4", N, N, N, N,
         37.5, 1.0, "Full time", 1, Y, Y, Y, Y, Y, Y, Y, "", "", "", "",
         N, Y, 0, 0, 0, "Main", "Not registered; no nights", ""),
        ("S12", "Parker Proxy", "Trainee BMS", "4", N, N, N, Y, 37.5, 1.0,
         "Full time", 1, Y, Y, Y, Y, Y, N, N, "", "", "", "", N, N, 0, 0, 0,
         "Main", "Trainee; supervised working only", ""),
        ("S13", "Reese Stub", "BMS", "6", Y, Y, Y, N, 37.5, 1.0,
         "Full time", 1, Y, Y, Y, Y, Y, Y, Y, "", "", "", "", Y, Y, 0, 0, 0,
         "Main", "", ""),
        ("S14", "Devon Filler", "BMS", "5", Y, N, N, N, 37.5, 1.0,
         "Full time", 1, Y, Y, Y, Y, Y, Y, Y, "", "", "", "", Y, Y, 0, 0, 0,
         "Main", "", ""),
        ("S15", "Harper Model", "Support Worker", "3", N, N, N, N, 22.5, 0.6,
         "Part time", 1, N, N, Y, Y, Y, N, N, "", "", 3, "", N, N, 0, 0, 0,
         "Support", "Specimen reception only", ""),
    ]


def _demo_patterns():
    """Jamie Mock works Mon–Wed one week, Wed–Fri the next."""
    Y, N = "Y", "N"
    return [
        ("S06", 1, Y, Y, Y, N, N, N, N),
        ("S06", 2, N, N, Y, Y, Y, N, N),
    ]


def _demo_competencies(period_start: date):
    """Fictional competencies, arranged to trigger each competency warning."""
    soon = period_start + timedelta(days=21)          # expires within 30 days
    later = period_start + timedelta(days=200)
    lapsed = period_start - timedelta(days=14)        # already expired
    achieved = period_start - timedelta(days=400)
    C, T, A, S, IT = (CompetencyStatus.COMPETENT, CompetencyStatus.TRAINER,
                      CompetencyStatus.ASSESSOR, CompetencyStatus.SUPERVISED,
                      CompetencyStatus.IN_TRAINING)
    Y, N = "Y", "N"
    # id, name, discipline, competency, status, achieved, review, expiry,
    # trainer, assessor, authoriser, notes
    return [
        ("S01", "Alex Sample", "HAEM", "Full blood count reporting", T,
         achieved, None, later, Y, Y, Y, ""),
        ("S01", "Alex Sample", "BT", "Transfusion authorisation", C,
         achieved, None, later, N, N, Y, ""),
        ("S02", "Jordan Test", "BT", "Crossmatch and electronic issue", T,
         achieved, None, later, Y, Y, Y, ""),
        ("S02", "Jordan Test", "HAEM", "Full blood count reporting", C,
         achieved, None, later, N, N, Y, ""),
        ("S02", "Jordan Test", "COAG", "Routine coagulation", C,
         achieved, None, later, N, N, N, ""),
        ("S03", "Casey Example", "BT", "Crossmatch and electronic issue", C,
         achieved, None, soon, N, N, Y, "Reassessment due"),
        ("S03", "Casey Example", "HAEM", "Full blood count reporting", C,
         achieved, None, later, N, N, N, ""),
        ("S04", "Morgan Demo", "BT", "Crossmatch and electronic issue", C,
         achieved, None, later, N, N, Y, ""),
        ("S04", "Morgan Demo", "COAG", "Routine coagulation", C,
         achieved, None, later, N, N, N, ""),
        ("S05", "Riley Placeholder", "MORPH", "Blood film morphology", C,
         achieved, None, later, Y, N, Y, "Only morphology scientist"),
        ("S05", "Riley Placeholder", "HAEM", "Full blood count reporting", C,
         achieved, None, later, N, N, N, ""),
        ("S06", "Jamie Mock", "HAEM", "Full blood count reporting", C,
         achieved, None, later, N, N, N, ""),
        ("S06", "Jamie Mock", "BT", "Crossmatch and electronic issue", C,
         achieved, None, lapsed, N, N, N, "Lapsed — needs reassessment"),
        ("S07", "Taylor Dummy", "HAEM", "Full blood count reporting", C,
         achieved, None, later, N, N, N, ""),
        ("S07", "Taylor Dummy", "COAG", "Routine coagulation", C,
         achieved, None, later, N, N, N, ""),
        ("S08", "Avery Specimen", "BT", "Crossmatch and electronic issue", C,
         achieved, None, later, N, N, N, ""),
        ("S08", "Avery Specimen", "HAEM", "Full blood count reporting", C,
         achieved, None, later, N, N, N, ""),
        ("S09", "Quinn Template", "COAG", "Routine coagulation", C,
         achieved, None, later, N, N, N, ""),
        ("S10", "Rowan Draft", "HAEM", "Full blood count reporting", C,
         achieved, None, later, N, N, N, ""),
        ("S10", "Rowan Draft", "BT", "Crossmatch and electronic issue", S,
         None, None, None, N, N, N, "Working towards sign-off"),
        ("S11", "Skyler Trial", "HAEM", "Sample processing", C,
         achieved, None, later, N, N, N, ""),
        ("S12", "Parker Proxy", "HAEM", "Full blood count reporting", IT,
         None, None, None, N, N, N, "In training"),
        ("S12", "Parker Proxy", "BT", "Crossmatch and electronic issue", IT,
         None, None, None, N, N, N, "In training"),
        ("S12", "Parker Proxy", "MORPH", "Blood film morphology", IT,
         None, None, None, N, N, N,
         "In training — does not count as morphology cover"),
        ("S13", "Reese Stub", "BT", "Crossmatch and electronic issue", C,
         achieved, None, later, N, N, Y, ""),
        ("S13", "Reese Stub", "HAEM", "Full blood count reporting", C,
         achieved, None, later, N, N, Y, ""),
        ("S14", "Devon Filler", "HAEM", "Full blood count reporting", C,
         achieved, None, later, N, N, N, ""),
        ("S14", "Devon Filler", "COAG", "Routine coagulation", C,
         achieved, None, later, N, N, N, ""),
        ("S15", "Harper Model", "HAEM", "Sample reception", S,
         None, None, None, N, N, N, "Supervised"),
    ]


def _demo_leave(period_start: date, period_end: date):
    return [
        ("S02", "Jordan Test", period_start + timedelta(days=7),
         period_start + timedelta(days=11), "A/L", "Booked leave"),
        ("S07", "Taylor Dummy", period_start + timedelta(days=14),
         period_start + timedelta(days=20), "A/L", "Booked leave"),
        ("S13", "Reese Stub", period_start + timedelta(days=2),
         period_start + timedelta(days=3), "S/L", "Recorded absence"),
        ("S08", "Avery Specimen", period_start + timedelta(days=21),
         period_start + timedelta(days=21), "S/D", "Training day"),
    ]


def build_demo_workbook(path) -> None:
    """A complete fictional laboratory that triggers every kind of warning."""
    workbook = Workbook()
    today = date.today()
    first = date(today.year + (today.month == 12), (today.month % 12) + 1, 1)
    last = (date(first.year + (first.month == 12), (first.month % 12) + 1, 1)
            - timedelta(days=1))

    _instructions_sheet(workbook, is_demo=True)
    _details_sheet(workbook, first, last, is_demo=True)
    _rules_sheet(workbook)

    sheet, header_row = _sheet(
        workbook, "Staff", STAFF_HEADERS, STAFF_WIDTHS,
        note="EXAMPLE DATA — every person below is fictional. Replace with your "
             "own staff before using this for real planning.")
    _write_rows(sheet, header_row, _demo_staff(first))
    _yes_no(sheet, header_row, [5, 6, 7, 8, *range(13, 20), 24, 25])

    sheet, header_row = _sheet(
        workbook, "Competencies", COMPETENCY_HEADERS, COMPETENCY_WIDTHS,
        note="EXAMPLE DATA — fictional competencies. Only Competent, Trainer and "
             "Assessor count as coverage.")
    _write_rows(sheet, header_row, _demo_competencies(first))
    _list_validation(sheet, header_row, 5, CompetencyStatus.ALL)
    _yes_no(sheet, header_row, [9, 10, 11])

    sheet, header_row = _sheet(workbook, "Week Patterns", PATTERN_HEADERS,
                               PATTERN_WIDTHS,
                               note="EXAMPLE DATA — one fictional member of staff "
                                    "on an alternating two-week pattern.")
    _write_rows(sheet, header_row, _demo_patterns())
    _yes_no(sheet, header_row, list(range(3, 10)))

    sheet, header_row = _sheet(workbook, "Leave", LEAVE_HEADERS, LEAVE_WIDTHS,
                               note="EXAMPLE DATA — fictional absence records.")
    _write_rows(sheet, header_row, _demo_leave(first, last))
    _list_validation(sheet, header_row, 5,
                     [code for code, *_ in DEFAULT_LEAVE_TYPES])

    _configuration_sheets(workbook)
    workbook.save(path)
