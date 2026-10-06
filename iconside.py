"""Icon-side gross margin from the ship file, lined up against the TRR-side figure.

The dashboard's profit figures are TRR-side: commission paid on each sold item
minus a cost looked up from the Icon ship file (SKU match, else brand x category
average). Finance reads the same ship file the other way round: invoice amount
minus standard cost on every line shipped to REAL001, by calendar window. The
two do not agree, and this module builds the table that puts them side by side
so the gap is visible rather than argued about.

Pure functions, no I/O: refresh.py feeds it the parsed wholesale rows and the
finished brand_table; a one-off script can do the same.
"""
from collections import Counter, defaultdict
from datetime import date, datetime


def _d(v):
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    if isinstance(v, str):
        for f in ('%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y'):
            try: return datetime.strptime(v[:10], f).date()
            except ValueError: pass
    return None


def build_icon_vs_trr(wh_rows, brand_table, normalize_brand, today):
    """wh_rows: dicts with brand, units, std_cost and, when the export carries
    them, ship_date, inv_amount, ext_cost, gp. Rows without ship_date are
    ignored on the Icon side (the fixture, or an export cut to 29 columns)."""
    year = today.year
    ytd_lo, ytd_hi = date(year, 1, 1), today
    py_lo, py_hi = date(year - 1, 1, 1), date(year - 1, 12, 31)

    def blank():
        return {'ytd': [0.0, 0.0, 0], 'py': [0.0, 0.0, 0], 'units': 0,
                'dollar_units': 0, 'costs': Counter(), 'last': None}
    by_brand = defaultdict(blank)
    for r in wh_rows:
        sd = _d(r.get('ship_date'))
        if sd is None: continue
        b = normalize_brand(r.get('brand')) or 'Unbranded'
        if py_lo <= sd <= py_hi: w = 'py'
        elif ytd_lo <= sd <= ytd_hi: w = 'ytd'
        else: continue
        inv = float(r.get('inv_amount') or 0)
        cost = float(r.get('ext_cost') or 0)
        units = int(r.get('units') or 0)
        acc = by_brand[b]
        acc[w][0] += inv; acc[w][1] += cost; acc[w][2] += units
        acc['units'] += units
        acc['costs'][round(float(r.get('std_cost') or 0), 2)] += units
        if float(r.get('std_cost') or 0) <= 1.0: acc['dollar_units'] += units
        if acc['last'] is None or sd > acc['last']: acc['last'] = sd

    trr = {row['brand']: row for row in brand_table}

    def gm(sales, cost):
        return (sales - cost) / sales if sales else None

    rows, tot = [], {'ytd': [0.0, 0.0, 0], 'py': [0.0, 0.0, 0],
                     'trr_comm': 0.0, 'trr_cogs': 0.0, 'units': 0, 'dollar_units': 0}
    last_ship = None
    for b, acc in by_brand.items():
        ytd_s, ytd_c, ytd_u = acc['ytd']; py_s, py_c, py_u = acc['py']
        cs, cc, cu = ytd_s + py_s, ytd_c + py_c, ytd_u + py_u
        if cu == 0 and cs == 0: continue
        t = trr.get(b)
        flat_cost, flat_units = (acc['costs'].most_common(1)[0]
                                 if acc['costs'] else (None, 0))
        flat_share = flat_units / acc['units'] if acc['units'] else 0
        # One line of "why": the thing most likely to make this brand's number wrong.
        flag = None
        if acc['units'] and acc['dollar_units'] / acc['units'] >= 0.05:
            flag = f"${1:.2f} cost on {acc['dollar_units']:,} of {acc['units']:,} units"
        elif flat_share >= 0.5 and acc['units'] >= 20:
            flag = f"flat ${flat_cost:,.2f} on {flat_share:.0%} of units"
        rows.append({
            'brand': b,
            'ytd_sales': ytd_s, 'ytd_gm': gm(ytd_s, ytd_c),
            'py_sales': py_s, 'py_gm': gm(py_s, py_c),
            'sales': cs, 'cost': cc, 'gp': cs - cc, 'gm': gm(cs, cc), 'units': cu,
            'trr_comm': t['commission_you'] if t else None,
            'trr_profit': t['profit'] if t else None,
            'trr_gm': (t['profit'] / t['commission_you'])
                      if t and t.get('profit') is not None and t['commission_you'] else None,
            'trr_sold': t['sold'] if t else None,
            'trr_conf': t['cost_confidence'] if t else None,
            'gap_pts': None, 'flag': flag,
        })
        r = rows[-1]
        if r['gm'] is not None and r['trr_gm'] is not None:
            r['gap_pts'] = (r['gm'] - r['trr_gm']) * 100
        for w in ('ytd', 'py'):
            for i in range(3): tot[w][i] += acc[w][i]
        tot['units'] += acc['units']; tot['dollar_units'] += acc['dollar_units']
        if t and t.get('profit') is not None:
            tot['trr_comm'] += t['commission_you']; tot['trr_cogs'] += t['cogs'] or 0
        if acc['last'] and (last_ship is None or acc['last'] > last_ship):
            last_ship = acc['last']
    rows.sort(key=lambda r: -r['sales'])

    cs = tot['ytd'][0] + tot['py'][0]; cc = tot['ytd'][1] + tot['py'][1]
    trr_costed = [t for t in brand_table if t.get('profit') is not None and t['commission_you']]
    trr_comm_all = sum(t['commission_you'] for t in trr_costed)
    trr_profit_all = sum(t['profit'] for t in trr_costed)
    return {
        'windows': {
            'ytd': f"{year} YTD",
            'py': f"{year - 1} full year",
            'ytd_through': last_ship.isoformat() if last_ship else None,
        },
        'totals': {
            'ytd_sales': tot['ytd'][0], 'ytd_gm': gm(tot['ytd'][0], tot['ytd'][1]),
            'py_sales': tot['py'][0], 'py_gm': gm(tot['py'][0], tot['py'][1]),
            'sales': cs, 'cost': cc, 'gp': cs - cc, 'gm': gm(cs, cc),
            'units': tot['units'], 'dollar_units': tot['dollar_units'],
            'trr_comm': trr_comm_all, 'trr_profit': trr_profit_all,
            'trr_gm': (trr_profit_all / trr_comm_all) if trr_comm_all else None,
            'trr_comm_matched': tot['trr_comm'],
            'trr_gm_matched': gm(tot['trr_comm'], tot['trr_cogs']),
        },
        'rows': rows,
    }
