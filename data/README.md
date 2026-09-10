# Data drop zone

Place your xlsx files here before running `python refresh.py` from the project root:

- `inventory.xlsx` — the TRR "Icon inventory export" (currently the whole catalog + status + commission)
- `sales.xlsx` — the Icon wholesale invoice ("Sales Order Detail — Vertical Sizes"), used as the cost source

These files are gitignored — they never get pushed. Only the resulting `index.html` (with numbers baked in) gets deployed.
