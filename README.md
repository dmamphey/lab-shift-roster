# LabRoster

**Competency-aware workforce planning for diagnostic laboratories.**

Import your workforce information, define your shift and competency
requirements, and LabRoster creates a draft rota while identifying staffing,
skills and coverage gaps. Workforce data is processed locally in your browser and
is not uploaded to Optymum SS.

The question LabRoster is built to answer is:

> *Do I have the right people, with the right competencies, in the right
> laboratory areas for this shift?*

## Who it is for

Diagnostic laboratory managers, initially in UK biomedical science and NHS-style
laboratory environments — blood sciences, haematology, transfusion, coagulation.
It assumes a workforce with mixed grades, mixed competencies, part-time and
flexible working patterns, and sections that must each be staffed by somebody
competent to work in them.

## What it does

- Balances work by **contracted hours**, not shift counts, so twenty ten-hour
  nights are not treated as equivalent to twenty seven-and-a-half hour days
- Honours **part-time and fixed working patterns**, including alternating weeks,
  availability windows, and caps on days, nights and weekends
- Models **competency properly**: a status from Not Trained to Assessor, with
  expiry dates, trainer, assessor and result-authoriser flags
- Keeps **grade and competence separate** — a Band 7 is not automatically
  competent in morphology, and a competent scientist is not automatically senior
- Allocates **one person to one section at a time**, so section coverage reflects
  who is actually standing there
- Calculates **rest from real shift times**, including across midnight
- Reports **single points of failure** at both workforce and shift level
- Produces a **ten-sheet Excel workbook** and an in-browser preview

## What it does not do

- It is **not** artificial intelligence. The engine is rules-based: it fills the
  shifts you define, hardest requirements first, and refuses to break a rule you
  have set. Where it cannot fill a shift without breaking one, it leaves the
  shift short and tells you.
- It does **not** make staffing decisions. Every roster it produces is a
  **draft requiring managerial review**.
- It does **not** determine legal or regulatory compliance. Rest intervals and
  other limits are rules *you* configure for your own laboratory. The wording
  throughout is "configured rest rule conflict", never "illegal shift".
- It does **not** replace organisational HR, payroll, ESR, legal, regulatory or
  managerial responsibilities.
- There are no user accounts, no external database, no cloud storage, no payroll
  or ESR integration, and no AI services.

## Using it

Open the application, then:

1. Download a workbook — blank, or the example laboratory
2. Add your workforce and requirements
3. Upload the workbook
4. Generate a draft roster
5. Review the issues raised
6. Export to Excel

Nothing is installed and nothing is uploaded.

## How local processing works

LabRoster runs a Python runtime compiled to WebAssembly (Pyodide) plus the
openpyxl spreadsheet library inside your browser tab. Your workbook is read by
JavaScript, handed to Python in the same tab, and the exported workbook is
written back out for download.

There is no server component. Once the page has loaded you can disconnect from
the network and it will still work. This is what allows the tool to be served as
static files while keeping workforce data on your own device.

We make no absolute security guarantees about your device or network. Include
only the workforce information the tool needs; do not add sensitive personal
details, and never add patient-identifiable information — LabRoster does not
require any.

## Running and deploying it

Any static host will serve it. The `labroster` package must sit beside
`index.html`, because the page fetches the modules at run time.

Locally:

```bash
python -m http.server 8765
```

Then open <http://localhost:8765>.

Deployment targets, all static:

- **GitHub Pages** — development and demonstration
- **tools.optymumss.com** — intended production home
- **Cloudflare Pages** or similar — if preferred

No conventional Python server is needed or wanted.

### Runtime dependencies

The Pyodide runtime and the openpyxl wheel are currently loaded from a public
CDN. In NHS-managed environments that means additional domains may need to be
reachable. Self-hosting them so everything is served from one trusted domain is
planned but **not yet implemented** — see *Known limitations*.

## Workbook structure

| Sheet | What it holds |
| --- | --- |
| **Instructions** | How to fill the workbook in |
| **Roster Details** | Organisation, department, site, rota name, period, preparer |
| **Rules** | Rest interval, consecutive limits, tolerances, warning thresholds |
| **Staff** | One row per person: grade, role, hours, FTE, working pattern, eligibility |
| **Week Patterns** | Only for staff on alternating weeks |
| **Competencies** | One row per person per discipline |
| **Shifts** | Shift codes with real start and finish times |
| **Shift Requirements** | Minimum staffing and competencies per shift |
| **Benches** | Laboratory sections, the competency each needs, and when |
| **Leave** | Absence, inclusive of both dates |
| **Leave Types** | Absence codes and colours |

Notes:

- **Staff ID** must be unique; the other sheets refer to it
- **Required Competencies** is written as discipline and number:
  `BT:1, HAEM:2, COAG:1`
- The **Mon–Sun** columns on the Staff sheet are the days somebody normally
  works; leave all seven blank for a fully flexible member of staff
- **Shift Codes** on the Benches sheet says which shifts a section must be
  staffed during
- Headers are matched loosely, so capitalisation, spacing and column order do
  not matter
- Dates are read as Excel dates or as text in `dd/mm/yyyy`, `dd.mm.yyyy` or
  `yyyy-mm-dd`

## Scheduling assumptions

- Nights are scheduled first, because they are the most constrained and they
  generate the recovery days everything else works around
- Nights are placed in blocks, followed by recovery days
- Day shifts are filled hardest-first: the shift with the fewest eligible people
  goes first, and within a shift the rarest capability is satisfied first
- One shift per person per day
- Competency demands are satisfied by **distinct-person matching**: three
  sections need three people, even if one person holds all three competencies
- Bench demand drives staffing, so a shift whose headcount minimum is one will
  still roster three people if three sections need cover
- Fairness compares only staff eligible for the same work

## Rule configuration

All on the **Rules** sheet, all yours to set:

| Rule | Default |
| --- | --- |
| Senior band threshold | 6 |
| Minimum rest hours between shifts | 11 |
| Maximum consecutive days | 6 |
| Maximum consecutive nights | 4 |
| Night block length | 3 |
| Recovery days after nights | 2 |
| Hours tolerance | 10% |
| Competency expiry warnings | 30, 60, 90 days |
| Max simultaneous bench assignments | 1 |
| Section rotation warning | 56 days |

The 11-hour rest default is a common starting point, not a compliance
determination. Set it to whatever your organisation has agreed.

## Issues and severities

| Severity | Meaning |
| --- | --- |
| **CRITICAL** | Resolve before publishing: unfilled shift, missing competency or authoriser, absent senior cover, somebody rostered while unavailable, a section relying on one person twice |
| **REVIEW** | Worth a look: materially off contracted hours, uneven weekend distribution, competency expiring, thin section cover, single point of failure |
| **PASSED** | Checked and satisfied |

Each issue carries the date, shift, section, staff involved, an explanation, and
the point a manager should check.

## Architecture

```
index.html            browser interface: dashboard, issues, preview, filters
labroster/
  __init__.py         version and workbook schema version
  timeutils.py        shift durations, midnight crossing, rest, week patterns
  models.py           staff, competency, shift, requirement, bench, rules
  workbook.py         reading and validating a manager's workbook
  template.py         blank template and example laboratory
  scheduler.py        the scheduling engine and bench allocation
  analysis.py         issue engine, hours, fairness, resilience, expiry
  export.py           the ten-sheet Excel export
  api.py              single entry point for browser and command line
tests/                pytest suite
```

Modules are small and independently testable, but the package is kept flat and
dependency-light so it loads into Pyodide with no build step. `index.html`
fetches each module at run time, which is why the repository stays plain text and
deploys to any static host.

## Testing

```bash
pip install -r requirements.txt pytest
python -m pytest tests/ -q
```

The suite covers contracted hour calculation, overnight durations, rest
intervals, part-time and alternating patterns, leave exclusion, night and weekend
eligibility and caps, consecutive day and night limits, expired competencies,
trainees not counting as cover, grade separated from specialist competence,
required authorisers, senior cover, unfilled shift detection, weekend fairness,
single points of failure, and — specifically — that one person is never counted
as covering more than one section at a time.

The example laboratory is a fixture in its own right: it is arranged to trigger
every category of warning, and tests assert that it does.

## Known limitations

- **Runtime dependencies are not yet self-hosted.** Pyodide and openpyxl load
  from a public CDN. Self-hosting is planned so everything can be served from one
  trusted domain.
- **Manual roster adjustment is prepared for but not exposed.** Assignments are
  tracked as auto-generated or manually adjusted, manual ones are never
  overwritten, and rules are re-checked against the finished roster — but there
  is no editing interface yet.
- **Section recency is reported, not scheduled around.** The data is collected and
  warnings are raised; rotation is not yet an input to the scheduler.
- **Part-shift bench blocks are not implemented.** The data model and the overlap
  primitive are in place so morphology 09:00–13:00 then coagulation 13:00–17:00
  can be added, but sections are currently covered for a whole shift.
- **The engine is greedy, not an optimiser.** It reliably produces a workable
  draft with a tight hours spread, but it does not prove optimality, and a
  human-built rota may sometimes be better.
- **Bank holidays are not modelled** as a distinct shift category.

## Roadmap

1. Self-hosted Pyodide and openpyxl, served from one domain
2. Manual roster adjustment in the browser, with rules re-checked live
3. Rotation and recency as scheduling inputs
4. Part-shift section allocation
5. Bank holiday handling

## Privacy summary

- Workforce data is processed locally in your browser
- Uploaded workbooks are not sent to Optymum SS
- LabRoster does not store roster information on an external database
- The generated workbook is created locally for download
- No absolute security guarantees are made
- Patient-identifiable data is never required and should never be added

## A note on GitHub Pages

The repository contains an empty `.nojekyll` file at its root. Without it, GitHub
Pages runs the files through Jekyll, which **silently excludes anything whose name
begins with an underscore** — including `labroster/__init__.py`. The site then
loads but cannot start, reporting that part of LabRoster is missing. Keep that
file if you deploy to Pages.
