# TRR Dashboard

Static analytics dashboard for The RealReal consignment operation. One HTML file, all data embedded, deploys to Vercel free tier.

## Live URL

- Production: https://trr-dashboard-jonch95-oss-projects.vercel.app
- Vercel project: `trr-dashboard` (team `jonch95-oss-projects`)
- Production branch: `claude/blissful-tesla-9ucnfs` — pushes to it auto-deploy in ~20s.

## Monthly refresh — the one-liner

Every month when the new TRR export lands:

```bash
cp ~/Downloads/<new_inventory>.xlsx data/inventory.xlsx
python refresh.py
git add index.html history.json && git commit -m "refresh $(date +%F)" && git push
```

Only replace `data/sales.xlsx` when a new Icon wholesale invoice has shipped.
The xlsx files are gitignored — only the rebuilt `index.html` is pushed.

To see the numbers before committing to them, add `--dry-run`: it prints the
same summary and leaves `index.html` untouched.

### Or just double-click `refresh.command`

In Finder, double-click `refresh.command`. It finds the newest
`*Inventory*Export*.xlsx` in `~/Downloads`, shows you the numbers first, asks
before publishing, then rebuilds, commits and pushes. No terminal, and nothing
leaves the machine except the finished `index.html`.

It needs `data/sales.xlsx` to already be in place — the Icon wholesale invoice
is the source of *all* brand and cost data, not just cost. Without it only
about 84 of your 205 brands are recognised and no item gets a cost, so the
refresh is not worth running.

### If the refresh refuses to run

`refresh.py` will stop rather than publish numbers it does not trust:

- **"SANITY GATE: this export looks wrong"** — item count fell more than 20%, or
  revenue moved more than 30%, against the previous run. A partial or truncated
  download is the usual cause. Check the export; if the numbers really are
  correct, re-run with `--force`.
- **"header no longer matches the recorded layout"** — a column moved in one of
  the exports. Both parsers read columns by position, so this would otherwise
  produce plausible-looking wrong numbers. Fix the indices at the top of
  `refresh.py`, then delete `schema.json` to re-record the layout.

`history.json` holds one small snapshot per refresh and drives the
month-over-month deltas on the dashboard — commit it along with `index.html`.
`schema.json` records the expected column layout of both exports.

## Refreshing with a new TRR export

1. Drop new files into `data/`:
   - `data/inventory.xlsx` — the TRR "Icon inventory export"
   - `data/sales.xlsx` — the Icon wholesale invoice (only replace this when a new wholesale batch has shipped)
2. Run the refresh script:
   ```bash
   python refresh.py
   ```
   This rebuilds `index.html` in place with the new numbers.
3. Ship it:
   ```bash
   git add index.html
   git commit -m "refresh $(date +%F)"
   git push
   ```
   Vercel picks up the push and redeploys in ~20 seconds.

## First-time setup

```bash
# Python 3.9+
pip install -r requirements.txt

# Vercel CLI (once)
npm i -g vercel
vercel login
vercel link           # link to a Vercel project (or create new)
vercel --prod         # first deploy
```

Or connect the git repo to Vercel via the dashboard (Import Project) — every push then auto-deploys.

## Privacy

The dashboard has real financial numbers embedded directly in `index.html`.
Two separate surfaces expose them, and they need separate decisions.

**The live site — currently protected.** The Vercel team is on the Pro plan and
Deployment Protection (Vercel Authentication) is already enabled on this project,
inherited from the team default. Visiting the URL without a Vercel login on this
account redirects to SSO, and responses carry `x-robots-tag: noindex`. Nothing
extra was turned on to achieve this. To change it: Vercel project settings →
Deployment Protection.

**The GitHub repo — public.** `jonch95-oss/TRRDash` is a public repository, so
`index.html` and every number in it are readable by anyone who finds the repo,
regardless of the protection on the Vercel side. Making the repo private closes
that surface: repo Settings → General → Danger Zone → Change visibility.

The deployment itself is minimal: `.vercelignore` keeps `refresh.py`,
`requirements.txt`, the `data/` drop zone, and this README out of the deployed
output, so only `index.html` is served.

## Layout

```
trr-dashboard/
├── index.html          # the dashboard (static, self-contained)
├── refresh.py          # rebuilds index.html from data/*.xlsx
├── requirements.txt    # openpyxl
├── history.json        # one snapshot per refresh — drives the deltas
├── schema.json         # recorded column layout of both exports
├── tests/              # fixture generator + pipeline smoke test
├── vercel.json         # static-site deploy config (no build step)
├── .vercelignore       # keeps the pipeline out of the deployment
├── IDEAS.md            # improvement backlog (UI + process)
├── data/               # xlsx exports (gitignored)
│   ├── inventory.xlsx
│   └── sales.xlsx
└── README.md
```

## Notes

- `refresh.py` reads two xlsx files, joins them (SKU + brand×category lookup for cost), computes all aggregates, and writes the entire DATA blob back into `index.html`. No separate JSON file is served — everything is inlined so the site works with zero backend.
- The wholesale invoice (`data/sales.xlsx`) is the source of cost data. Items received after the invoice's cutoff fall back to brand×category averages — flagged as MED confidence in the dashboard.
- If TRR ever changes the column layout of the inventory export, the top of `refresh.py` (the `for i, r in enumerate(ws.iter_rows(...))` block) is where to adjust column indices.
