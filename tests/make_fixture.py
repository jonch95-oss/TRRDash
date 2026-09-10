#!/usr/bin/env python3
"""Generate synthetic xlsx fixtures matching the two export layouts refresh.py reads.

Nothing here touches real data — the point is to exercise the pipeline in CI
without any TRR export being present. Column positions mirror the indices
refresh.py uses; if those change, this file must change with them.
"""
import datetime, random, sys
from pathlib import Path
from openpyxl import Workbook

# Wholesale invoice: refresh.py reads r[5] category, r[7] brand, r[8] gender,
# r[19] style, r[25] units, r[28] std_cost.
WH_COLS = 30
BRANDS = [('GUCCI', 'SHOES', 'WOMENS'), ('PRADA', 'BAGS', 'WOMENS'),
          ('TOM FORD', 'RTW', 'MENS'), ('BURBERRY', 'RTW', 'WOMENS'),
          ('FENDI', 'ACCESSORIES', 'UNISEX'), ('RAYBAN', 'GLASSES', 'UNISEX')]
DISPLAY = {'GUCCI': 'Gucci', 'PRADA': 'Prada', 'TOM FORD': 'Tom Ford',
           'BURBERRY': 'Burberry', 'FENDI': 'Fendi', 'RAYBAN': 'Ray-Ban'}
NAMES = {'SHOES': ['Leather Sneakers', 'Suede Ankle Boots', 'Patent Heels'],
         'BAGS': ['Shoulder Bag', 'Leather Tote', 'Mini Crossbody'],
         'RTW': ["Men's Wool Blazer", 'Silk Dress', 'Cotton Shirt'],
         'ACCESSORIES': ['Leather Belt', 'Wool Scarf', 'Card Holder'],
         'GLASSES': ['Aviator Sunglasses', 'Square Sunglasses', 'Round Sunglasses']}
STATUSES = (['sold'] * 60 + ['commission_paid'] * 15 + ['on_sale'] * 15
            + ['consigner_returned'] * 6 + ['accepted'] * 4)


def build(out_dir, items=1200, seed=7):
    out = Path(out_dir)
    (out / 'data').mkdir(parents=True, exist_ok=True)
    rng = random.Random(seed)

    wb = Workbook(); ws = wb.active
    ws.append([f'wh_col_{i}' for i in range(WH_COLS)])
    styles = []
    for n, (brand, cat, gender) in enumerate(BRANDS):
        for k in range(6):
            style = f'{brand[:3]}{n}{k:02d}'
            styles.append((style, brand, cat, gender))
            row = [None] * WH_COLS
            row[5], row[7], row[8], row[19] = cat, brand, gender, style
            row[25], row[28] = rng.randint(5, 40), round(rng.uniform(40, 600), 2)
            ws.append(row)
    wb.save(out / 'data' / 'sales.xlsx')

    # Inventory: two header rows, then exactly 11 columns.
    wb = Workbook(); ws = wb.active
    ws.append(['TRR Inventory Export'])
    ws.append(['SKU', 'VSKU', 'Name', 'Price', 'Status', 'Image', 'Description',
               'Comm Rate', 'Comm Amount', 'Received', 'Sale Date'])
    base = datetime.date(2025, 1, 1)
    for i in range(items):
        style, brand, cat, gender = rng.choice(styles)
        name = f'{DISPLAY[brand]} {rng.choice(NAMES[cat])}'
        price = round(rng.uniform(90, 900), 2)
        status = rng.choice(STATUSES)
        recv = base + datetime.timedelta(days=rng.randint(0, 400))
        sold = status in ('sold', 'commission_paid')
        # A few sold items deliberately carry no sale date, mirroring the real
        # export, so the dashboard's reconciliation note gets exercised.
        sale = (recv + datetime.timedelta(days=rng.randint(5, 300))
                if sold and rng.random() > 0.02 else None)
        ws.append([f'S{i:06d}',
                   style if rng.random() < 0.55 else f'{style}-{rng.randint(1, 9)}',
                   name, price, status, None, f'{gender.title()} {cat.lower()} item',
                   0.65, round(price * 0.65, 2) if sold else 0, recv, sale])
    wb.save(out / 'data' / 'inventory.xlsx')
    return out


if __name__ == '__main__':
    d = build(sys.argv[1] if len(sys.argv) > 1 else '.')
    print(f'fixtures written to {d}/data')
