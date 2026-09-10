# Improvement backlog — status

Everything proposed in the review has been implemented. This file records what
each change does and what was learned along the way. The calibrated pipeline
logic — the brand/category/gender classifiers and the cost lookup — was not
touched; every change is a guard, a new aggregate, or presentation.

## Process — done

| # | Change | Where |
|---|---|---|
| P1 | Sanity gate: refuses to overwrite `index.html` if item count falls >20% or revenue moves >30% vs the previous run. `--force` overrides. | `refresh.py` |
| P2 | `history.json`: one snapshot per refresh, driving month-over-month deltas. | `refresh.py` |
| P3 | Schema guard: fingerprints both header rows on first run and aborts on drift, naming the columns that moved. Also probes that numeric columns are numeric. | `refresh.py` |
| P4 | `--dry-run`: prints the summary, touches nothing. | `refresh.py` |
| P5 | Smoke test over synthetic fixtures — happy path, dry run, gate, schema guard. Runs on every push. | `tests/`, `.github/workflows/smoke.yml` |
| P6 | `openpyxl` pinned to `==3.1.5`. | `requirements.txt` |

## UI — done

| # | Change |
|---|---|
| U1 | Phone layout at ≤700px: 2-up KPIs, horizontally scrolling tabs, wide tables scroll inside their own box. Verified at 390px with no page-level overflow. |
| U2 | Month-over-month delta chips on all eight headline KPIs, direction-aware (a rise in days-to-sell reads as bad). Appear once a second refresh exists. |
| U3 | "What the new export changed" now reports the actual change between exports instead of describing on-hand aging. |
| U4 | New Overview panel: the confidence split, an invoiced-cost-only profit figure, and a band showing where profit lands if the estimates are off. |
| U5 | Filter bar is sticky. |
| U6 | Tab and filter state live in the URL hash, so a view can be bookmarked. |
| U7 | Already worked — `cost_confidence` was in the CSV export all along. No change needed. |
| U8 | New "Capital parked in stale stock" panel: headline figures plus a brand × category worklist ranked by cost tied up. |

## Data findings — surfaced in the dashboard

- **170 sold items carry no sale date** (~$11.7K commission). They sit in the
  headline totals but cannot land in a month, so the monthly charts will always
  sum slightly below the KPI row. The Monthly tab now says so rather than
  leaving it to be discovered as an apparent bug.
- **Pre-2022 is 2.7% of the catalog** (1,099 of 40,538 items) but spanned eight
  of thirteen years on the x-axis. Monthly charts now default to 2022 onward,
  with full history one click away.
- **22 of 61 cohorts hold fewer than 50 items**, several showing 100%
  sell-through on samples of 8–30. Those rows are now marked THIN.

## Bugs found and fixed along the way

- **The Overview charts never rendered on first load.** Chart building was wired
  only to tab clicks, and nothing invoked the Overview handler at startup, so
  the default landing tab showed two blank canvases until you clicked away and
  back. The deep-link routing (U6 — the idea I had recommended skipping) drives
  the initial render and fixes it.
- **Three different cost-confidence figures were shown at once.** The masthead
  said 18/69/13 and was updated by `refresh.py`; the five-line summary hardcoded
  "94%"; the caveats paragraph hardcoded "31% / 63% / 6%". Only the masthead was
  right. All three now compute from the data.
- **The refresh date was never stamped** into the masthead (regex expected a
  space where the markup has a newline) — fixed earlier in the session.
- **The first Vercel build failed**, having auto-detected a Python project from
  `requirements.txt`.

## Not done, deliberately

Nothing outstanding. `history.json` will be empty of comparisons until the
second real refresh, at which point the delta chips and the "what changed"
paragraph populate themselves.
