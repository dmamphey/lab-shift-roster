# Changelog

## 3.0.0-beta — Optymum SS release

### Branding and positioning

Optymum SS wordmark and brand palette (#1878C4 primary, #17395F navy) applied
throughout, with a branded hero panel, favicon, meta description and Open Graph
tags. Version, company and contact details come from `labroster.__version__` alone,
so nothing duplicates a version string. Footer links for Privacy, User guide,
Feedback and About.

The logo file is a recreation, because the artwork was supplied as an image rather
than a vector. Replacing `assets/optymum-ss-logo.svg` with the original needs no
other change.

### Hours and absence

- Hours are accounted as **target, worked, credited absence, total accounted,
  variance**, and balancing uses total accounted. Previously a person with five
  days of annual leave worked the same hours as colleagues with none, compressed
  into fewer days.
- Credited hours derive from each person's own working pattern, configurable per
  leave type, with an optional explicit figure per absence.
- `Max Weekly Hours` added as a hard ceiling.
- Removed the `Senior band threshold`, which was read from the workbook and never
  used by anything. Seniority is recorded per person; grades use `Min Band`.

### Measurement

- `Shift coverage` renamed **`Staffing slot coverage`** — it only ever measured
  whether positions were occupied.
- Added **`Shifts meeting all configured requirements`**, which checks every
  condition per shift instance. On the challenging example: 100% of positions
  occupied, 63.5% of shifts actually satisfying everything.
- Coverage gaps are recomputed from the **finished** roster rather than the
  scheduler's build log, which could report problems that had since been resolved.
- Related warnings are consolidated by root cause, with Required, Available, Impact
  and a review point, and traceability to the underlying checks retained.
- Issues open with severity counts and main causes by frequency.

### Fairness

- Weekend days are shared by who has done fewest, not by hours owed. Ordering by
  absolute deficit concentrated weekends on full-timers, whose larger target gave
  them a larger deficit.
- **Nights are shared evenly by default.** Concentrating them while eligible
  colleagues do none is not acceptable in practice. The cost is documented: night
  blocks book recovery days, so sharing can reduce day-service cover. Configurable.
- Night fairness is judged on **blocks**, since ten three-night blocks between eight
  people cannot come out level.

### Manager control

- **Manual adjustment in the browser.** Change who works a shift; only eligible
  staff are offered, with their competencies; excluded staff are counted with the
  reason. Every check re-runs. Removals are honoured as blocked days, or the
  scheduler would simply refill with the same person. Changes survive
  `Generate alternative roster` and can be undone individually or all at once.
- **Start new roster** clears everything without a page reload.
- The random seed is gone from the interface. `Generate alternative roster` picks
  its own; the number survives as `Roster generation ID` under Advanced settings.

### Outputs

- **Two exports.** Staff rota (three sheets, data-minimised: no competency records,
  hours, fairness, vulnerabilities or absence reasons) and manager report (all ten
  analytical sheets). Both from one draft.
- Shift and absence colours are shown as colours in the workbook, with a Preview
  cell, alongside the hex codes.
- `Download blank template` renamed `Download blank workbook`.

### Demonstrations

- **Balanced example** added: 100% slot coverage, 100% requirement compliance, zero
  criticals, no rest conflicts, no single points of failure.
- The **challenging example** is now labelled as deliberate and not typical.

### Deployment

- The runtime can be **self-hosted** with `python tools/fetch_runtime.py`: 13.7 MB
  into a git-ignored `vendor/`. With it present, loading and generating contacts no
  host but the one serving the page. Falls back to the CDN when absent, so
  development and the Pages demo need no download.
- Sizes, deployment, caching, update procedure and licence obligations documented.

### Fixed

- A single-handed night requiring `BT:1, HAEM:1` was impossible to satisfy:
  distinct-person matching is right for sections, which are physical stations, but
  not for a shift's competency list. Compliance on the balanced example went from
  69% to 100%.
- Disciplines are written as short codes (`BT`, `HAEM`, `COAG`, `MORPH`), matching
  laboratory usage and keeping existing workbooks readable.
- `.nojekyll` added: GitHub Pages was silently excluding `labroster/__init__.py`
  because Jekyll drops underscore-prefixed files, so the deployed site could not
  start.
- Browser hours table was showing a column that no longer existed.
- Credited hours divided by availability rather than contracted days, so somebody
  marked as able to work any day was treated as working a seven-day week.

### Known limitations

Bank holidays are not modelled. Rotation and recency are reported but are not
scheduling inputs. Section allocation is whole-shift. Workbook sheet
simplification is still to come.

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

- Repositioned as Lab Shift Roster, with the local-processing guarantee prominent.
- Six-step explanation, then two distinct downloads: **blank template** (headings,
  guidance and dropdowns, no employee records) and **example laboratory**
  (a complete fictional department, clearly labelled as fictional).
- **In-browser roster preview**: staff down the side, dates across, horizontal
  scrolling with a sticky name column, and filters for staff, shift and date
  range. Section coverage shown per date and section.
- Technical language removed from the normal interface. Loading says
  "Preparing Lab Shift Roster…" then "Lab Shift Roster ready". Pyodide, WebAssembly and openpyxl
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
