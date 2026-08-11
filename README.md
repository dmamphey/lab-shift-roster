# Lab shift roster generator

Generates a lab shift rota as a formatted Excel workbook: a colour-coded
calendar grid (staff as rows, dates as columns) with bench/section allocation
rows underneath, plus a summary sheet of shifts worked per person.

## Install

Python 3.9+ and one dependency:

```bash
pip install openpyxl
```

## Use — in a browser, nothing installed

`index.html` runs the whole tool client-side. It loads Pyodide (CPython compiled
to WebAssembly), installs `openpyxl`, then imports `lab_roster.py` and calls the
same `Scheduler` and `write_rota` as the command line does — no reimplementation,
no second copy of the rules.

**Your data never leaves the machine.** There is no server and no upload: the
workbook you choose is read by JavaScript, handed to Python inside the browser
tab, and the generated rota comes back as a download. You can go offline once
the page has loaded. That matters here, because a filled-in input workbook
contains named staff and sick leave dates — health data about identifiable
people — and this design keeps it off other people's infrastructure entirely.

It has to be served over HTTP, not opened as a `file://` path, because it fetches
`lab_roster.py` alongside it. Locally:

```bash
python -m http.server 8765
```

Then open <http://localhost:8765>. To share it, put `index.html` and
`lab_roster.py` on any static host — GitHub Pages, an intranet site, a
SharePoint library. Note that GitHub Pages for a **private** repo needs a paid
plan; on the free tier the repository must be public.

First load pulls roughly 10–25 MB of Python runtime and is cached afterwards.
The "Try it with example data" button builds a fictional workbook and schedules
against it, so anyone can see what the tool does without filling anything in.

## Use — double-click

Double-click **`Generate Rota.bat`**. That is the whole workflow.

- **First run** creates `roster_input.xlsx` and opens it. Replace the example
  staff with your real staff, set the date range on the Settings sheet, add any
  leave, then save and close it.
- **Every run after that** reads that workbook, writes a rota named
  `rota_<date>_<time>.xlsx` and opens it. The timestamp means a new file every
  run, so you never overwrite a rota you have already shared.

The launcher finds Python itself, installs `openpyxl` if it is missing, and
keeps the window open so you can read any message. To start over from a clean
template, delete or rename `roster_input.xlsx` and run it again.

## Use — command line

Two steps. First write yourself an input workbook:

```bash
python lab_roster.py template --out roster_input.xlsx
```

Fill it in (it ships pre-populated with example data so you can see the shape),
then generate the rota:

```bash
python lab_roster.py generate --input roster_input.xlsx --out rota.xlsx
```

Useful overrides — handy for producing next month without editing the workbook:

```bash
python lab_roster.py generate --input roster_input.xlsx --out oct.xlsx --start 01/10/2026 --end 31/10/2026
```

`--seed N` changes the tie-breaking. Same inputs + same seed always produce the
same rota, so a rota is reproducible; change the seed to shuffle a rota you
don't like.

Neither workbook is committed to this repository — `.gitignore` excludes all
spreadsheets. Both are reproducible from the two commands above, and the input
workbook is the one you fill in with real staff names, bands and leave. Git
history is permanent, so the file type is excluded outright rather than relying
on remembering which copy is safe to commit.

## The input workbook

| Sheet | What it holds |
| --- | --- |
| **Settings** | Date range, rule limits, weekend days, random seed |
| **Staff** | Name, band, skills, whether they can do nights, group, optional personal shift cap |
| **Shifts** | Shift code, name, times, hours, which days it runs, how many people, whether a senior is required, whether it counts as a night, colour |
| **Leave** | Name, from, to (inclusive), leave type |
| **Benches** | Section name, required skill, which days, minimum staff on weekdays and at weekends |
| **Leave types** | Absence codes, labels and colours |

Notes:

- **Band drives seniority.** Anything at or above `Senior band threshold`
  (default 6) counts as senior. `8a` sorts above `8`, and above `7`.
- **Skills** are free text codes; they only need to match between the Staff
  sheet and the Benches sheet. The defaults use `BT`, `COAG`, `HAEM`, `MORPH`.
- **Groups** split the calendar into labelled blocks, mirroring the
  Main / Extras split in the existing rotas.
- **Min staff weekend** lets a bench drop to a skeleton service at weekends,
  or close entirely with `0`.
- Dates are read as real Excel dates or as text in `dd/mm/yyyy`, `dd.mm.yyyy`
  or `yyyy-mm-dd`. Headers are matched loosely, so case, spacing and column
  order don't matter.
- The bands, skills and night flags in the generated template are
  **placeholders** — replace them with your real staff details.

## Rules

Hard constraints — the scheduler will never break these. If it cannot fill a
shift without breaking one, it leaves the shift short and reports it:

- nobody is scheduled on a day they are on leave
- one shift per person per day
- no shift on the day after a night (this is what bans night → early)
- no more than `Max consecutive days` worked in an unbroken run
- nights only go to staff flagged as night-capable
- nights are placed in blocks (default 3) followed by rest days (default 2)

Soft targets, in priority order:

1. total shifts spread as evenly as possible
2. weekend duty spread as evenly as possible
3. at least one senior on every shift flagged `Requires senior`
4. bench skill coverage each day

Nights are scheduled first because they are the most constrained and they
generate the rest days everything else has to work around. Day shifts are then
filled hardest-first, and a final repair pass swaps in skilled staff where a
bench would otherwise have nobody — but only when the incoming person has
strictly fewer shifts than the person they replace, so coverage repair can
never widen the fairness spread.

## Output

**Rota sheet** — title, month bands, day numbers and weekday names, then one row
per staff member with the shift code in each cell, colour-coded by shift type.
Leave shows as its own code and colour. Weekends are tinted. Underneath:
a "staff on duty" count row, the bench allocation rows (staff shown as
initials, uncovered benches highlighted red), and the KEY. Senior staff names
are bold. Panes are frozen at the first date column and the page is set up
landscape, fit to width.

**Summary sheet** — per person: a column per shift type, total shifts,
weekends, nights, leave days, days off and total hours, with a totals row.
Highest and lowest totals are highlighted. Below that, a distribution block
(min / max / mean / spread) and a **rule check block** that reports PASS or the
number of issues for each rule, with a detail table listing every one.

Always read the rule check block. It is how the tool tells you a rota is short
of cover rather than pretending it isn't.

## Limits worth knowing

- The scheduler is greedy with a targeted repair pass, not an exhaustive
  optimiser. It reliably lands a tight spread (1–2 shifts across a month on
  realistic inputs) but it doesn't prove optimality.
- Part-time contracted hours and fixed non-working days are not modelled. Every
  schedulable person is treated as available on every non-leave day, subject to
  the rules above. Use `Max shifts` on the Staff sheet, or leave entries, to
  approximate a part-timer.
- There is no on-call, no shift swapping and no carry-over of fairness between
  runs: each run balances only within its own date range.
