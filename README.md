# TRR Dashboard

Static analytics dashboard for The RealReal consignment operation. One HTML file, all data embedded, deploys to Vercel free tier.

## Live URL

Set after first deploy:
- Production: `https://<your-project>.vercel.app`

## Monthly refresh — the one-liner

Every month when the new TRR export lands:

```bash
cp ~/Downloads/<new_inventory>.xlsx data/inventory.xlsx
python refresh.py
git add index.html && git commit -m "refresh $(date +%F)" && git push
```

Only replace `data/sales.xlsx` when a new Icon wholesale invoice has shipped.
The xlsx files are gitignored — only the rebuilt `index.html` is pushed.

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

The dashboard contains real financial numbers embedded in the HTML. Anyone with the URL sees everything. Options:

- **Free**: rely on the unguessable Vercel URL and don't share it. Do not add the URL to any indexable page.
- **Pro ($20/mo)**: enable "Deployment Protection" in Vercel project settings → require a Vercel login or password to view.

## Layout

```
trr-dashboard/
├── index.html          # the dashboard (static, self-contained)
├── refresh.py          # rebuilds index.html from data/*.xlsx
├── requirements.txt    # openpyxl
├── data/               # xlsx exports (gitignored)
│   ├── inventory.xlsx
│   └── sales.xlsx
└── README.md
```

## Notes

- `refresh.py` reads two xlsx files, joins them (SKU + brand×category lookup for cost), computes all aggregates, and writes the entire DATA blob back into `index.html`. No separate JSON file is served — everything is inlined so the site works with zero backend.
- The wholesale invoice (`data/sales.xlsx`) is the source of cost data. Items received after the invoice's cutoff fall back to brand×category averages — flagged as MED confidence in the dashboard.
- If TRR ever changes the column layout of the inventory export, the top of `refresh.py` (the `for i, r in enumerate(ws.iter_rows(...))` block) is where to adjust column indices.
