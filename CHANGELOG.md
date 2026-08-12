# Changelog

## 2.0.0 — competency-aware workforce planning

A rebuild of the roster generator into a laboratory workforce planning product.
Workbooks from version 1 are **not** compatible: they do not carry contracted
hours, competencies or shift requirements, and those cannot safely be guessed.
An older workbook is recognised and explained rather than crashing or being given
invented defaults.

### The defect that prompted much of this

One scientist competent in transfusion, haematology and coagulation was recorded
as covering **all three benches at once**, and no coverage problem was reported.
Already-allocated staff were sorted lower in the candidate list but not removed,
and coverage was judged on who happened to be on duty holding a competency rather
than on who was actually allocated.

Fixing the allocator alone was not enough. The requirement check also counted a
multi-skilled person once per competency, so the scheduler would roster two people
for three sections and only then discover the conflict. Requirements are now
satisfied by distinct-person matching, so three sections cause three people to be
rostered.

### Scheduling

- **Contracted hours replace shift counts.** Staff carry contracted weekly hours,
  FTE, a target for the period, allocated hours and variance. Twenty ten-hour
  nights total 200 hours where twenty seven-and-a-half hour days total 150.
- **Shift length is derived from configured times**, so the stated times and the
  hours can no longer disagree. Shifts crossing midnight are measured correctly:
  20:00–08:00 is twelve hours.
- **Working patterns are hard constraints**: fixed non-working days, alternating
  week cycles, earliest start and latest finish, days per week, night and weekend
  eligibility, caps on nights and weekends, and a personal consecutive-day limit.
- **Rest is calculated from real shift times** in both directions. A late
  finishing at 21:00 followed by an early starting at 07:00 is ten hours, below
  the configured interval, and is prevented — then re-checked against the finished
  roster so a manual override cannot hide one.
- **Fairness compares peers only.** Night counts are compared between people
  eligible for nights, weekends between those eligible for weekends. Total hours,
  nights, Saturdays, Sundays, full weekends, lates and earlies are all tracked.

### Competency

- Structured records replace binary skills: status from Not Trained through In
  Training, Supervised, Competent, Trainer and Assessor to Expired, with dates
  achieved, review and expiry, and trainer, assessor and authoriser flags.
- **In Training and Supervised never satisfy a requirement** for an independently
  competent member of staff. **Expired records never count.**
- **Grade, seniority and competence are three separate things.** Registration,
  seniority, shift coordination, discipline competence, result authorisation,
  training and assessing are each checked independently.
- Expiry warnings at configurable thresholds, 30/60/90 days by default.

### Reporting

- **Issue engine** with CRITICAL, REVIEW and PASSED severities. Each issue
  carries date, shift, section, staff, a plain-English explanation and the point a
  manager should check.
- **Single point of failure analysis**, distinguishing workforce-level resilience
  (one morphology scientist in the department) from shift-level resilience (one
  rostered on Tuesday). Thin shift cover is reported once per shift and discipline
  with a count of affected dates, not once per day.
- **Manager-facing dashboard**: roster status, shift coverage percentage, unfilled
  shifts, uncovered sections, senior cover gaps, competency gaps, rest conflicts,
  staff outside target hours, expiring competencies, and weekend and night
  fairness.
- **Section recency** collected and reported, with a configurable interval. No
  rotation rule is invented.

### Interface

- Repositioned as LabRoster, with the local-processing guarantee prominent.
- Six-step explanation, then two distinct downloads: **blank template** (headings,
  guidance and dropdowns, no employee records) and **example laboratory**
  (a complete fictional department, clearly labelled as fictional).
- **In-browser roster preview**: staff down the side, dates across, horizontal
  scrolling with a sticky name column, and filters for staff, shift and date
  range. Section coverage shown per date and section.
- Technical language removed from the normal interface. Loading says
  "Preparing LabRoster…" then "LabRoster ready". Pyodide, WebAssembly and openpyxl
  are described only inside an optional "Privacy and technical information"
  section. The random seed became "Roster variation" with a
  **Generate alternative roster** button, under Advanced options.
- Accessibility: every input labelled, real buttons, tab roles, `aria-live` status,
  visible focus rings, status carried by words as well as colour, and contrast
  measured at 15.6:1 for body text and 8.2:1 for status pills in dark mode.

### Excel export

Ten worksheets: Instructions, Roster, Staff, Competencies, Shift Requirements,
Bench Allocations, Issues, Hours Summary, Fairness Summary and Competency Expiry.
Freeze panes, filters, sensible widths, landscape fit-to-width for printing, a
legend, and conditional colour used alongside words. The roster keeps the shape
managers already read and adds per-shift staffing counts that highlight when a
shift is below its minimum.

### Validation

Every problem is reported in one pass, naming the sheet, row and column heading
rather than a Python key. Checks include required sheets and columns, duplicate
staff IDs and shift codes, negative hours, invalid competency statuses, unknown
staff references, zero-length shifts, leave dates the wrong way round, dates
outside the period, and availability rules that would make somebody impossible to
roster.

### Organisation-neutral

The default rota name is "Laboratory Staff Rota". Organisation, department, site,
rota name, period and preparer are configurable metadata and appear in the export.

### Architecture and tests

Split from a single 1,665-line module into a flat, dependency-light package —
`timeutils`, `models`, `workbook`, `template`, `scheduler`, `analysis`, `export`,
`api` — kept loadable into Pyodide with no build step. 130 tests.

### Not yet done

Self-hosted runtime dependencies, a manual roster editing interface, rotation as a
scheduling input, part-shift section allocation, and bank holiday handling. See
the README's *Known limitations*.

---

## 1.x — lab shift roster generator

- Draft roster generation from an Excel workbook, in the browser via Pyodide
- Colour-coded calendar grid, bench allocation rows, per-person summary
- Rules: leave respected, one shift per day, no shift after a night, maximum
  consecutive days, nights in blocks with rest days
- Shift-count fairness, with a senior defined as Band 6 or above
- Double-click launcher for Windows, and a command-line interface
