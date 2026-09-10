# Improvement backlog

Observations from a pass over `index.html` (v4) and `refresh.py`. Nothing here is
implemented — the calibrated pipeline logic (brand/category/gender classifiers,
cost lookup) has deliberately not been touched. Ordered by value per unit of risk.

## Process

**1. Sanity gate before overwrite (highest value).**
`refresh.py` overwrites `index.html` unconditionally. A truncated or partial TRR
export — one bad download — silently republishes a wrong dashboard, and the only
tell is that the numbers look off. Before writing, compare against the previous
run and refuse to proceed if `total_items` drops more than ~20% or revenue swings
more than ~30%, unless `--force` is passed. This is the guard that most directly
protects the live site.

**2. Snapshot history.**
Append `meta` + `headline` to a small `history.json` on every run. It costs a few
hundred bytes a month and unlocks everything in UI #2 and #3 below: real
month-over-month deltas, and a sell-through trend line over time. Purely
additive — no change to how any number is computed.

**3. Fail loudly on a column-layout change.**
The parsers address columns positionally: the wholesale invoice reads `r[7]`,
`r[19]`, `r[25]`, `r[28]`, and the inventory unpacks exactly 11 columns. If TRR
or Icon reorders or inserts a column, the likely outcome is not a crash but a
quiet mis-parse — brands read from the wrong field, costs off by a column. Assert
the header row matches the expected names and abort with a diff if it does not.

**4. `--dry-run`.**
Print the summary block without touching `index.html`, so a new export can be
eyeballed before it becomes the live site.

**5. Smoke test in CI.**
A GitHub Action that runs `refresh.py` against a tiny synthetic fixture on every
push would catch a broken pipeline before a monthly refresh does. A fixture
generator matching both schemas was written and validated during setup.

**6. Pin `openpyxl`.**
`requirements.txt` says `openpyxl>=3.1.0`. Pinning to the exact version in use
removes the chance that an upstream release changes parsing behaviour between one
monthly refresh and the next.

## UI

**1. Phone layout.**
There is a single breakpoint, at `max-width: 1100px`, and it only collapses the
2- and 3-column grids. On a phone the wide tables — all-brands, and the
brand × category × gender matrix — will overflow the viewport. Worth a ≤700px
breakpoint: KPIs 2-up, each table in its own `overflow-x: auto` container,
reduced chart heights. Relevant if the dashboard ever gets checked away from a
desk.

**2. Month-over-month deltas on the headline KPIs.**
The KPI row shows absolute values only. For a dashboard whose whole cadence is
monthly, "84.6% sell-through" is much less useful than "84.6%, +1.3pt vs last
refresh". Needs process #2.

**3. "What the new export changed" does not report a change.**
The `story-4` heading promises a delta; the text it renders describes current
on-hand aging. Either retitle it to match what it says, or wire it to real
deltas once history exists.

**4. Show how much of profit is estimated (most decision-relevant).**
The masthead reports the cost match as 18% exact / 69% average / 13% portfolio,
which means the large majority of the profit figure rests on brand×category
averages rather than invoiced cost. Presenting a single blended profit number
understates that. Showing profit computed from HIGH-confidence items alongside
the blended figure — or a band rather than a point — would make the uncertainty
legible at the point where decisions get made.

**5. Sticky filter bar.**
The brand / category / gender / search controls scroll out of view on the long
tables, so changing a filter means scrolling back up. `position: sticky` on that
row fixes it.

**6. Deep-linkable state.**
Encode the active tab and filter selection in the URL hash so a particular view
("Gucci / Bags / Womens") can be bookmarked or reopened directly.

**7. Carry cost confidence into the CSV export.**
The export drops the per-item confidence flag, so any downstream analysis of the
exported rows cannot tell invoiced cost from an estimate.

**8. Make aging actionable.**
`on_hand_aging_cost` already holds capital tied up by age bucket. Surfacing the
>365-day bucket as an explicit "capital parked in stale inventory" figure, with
the underlying items listed, turns that panel from a report into a decision.
