"""
Most Traded Universe - Excel Generator
Creates Most.Traded.Universe.xlsx with all spread/fly/condor permutations
for SFR, SFI, ER, IR, BA futures, with Bloomberg BDP volume formulas.

Ticker conventions (confirmed):
  Spreads:  SFR/SFI/IR/BA: {prefix}{near}{far}  e.g. SFRM6U6
            ER: ER{near}ER{far}                  e.g. ERM6ERU6
  Flies:    B{prefix}{near_wing}{far_wing}        e.g. BSFRM6Z6
            Prefixes: SFR→BSFR, SFI→BSFI, ER→BER, IR→BIR, BA→BBA(unconfirmed)
  Condors:  C{prefix}{first}{last}               e.g. CSFRM6H7
            SFR only (not exchange-listed for other products)
"""
import xlsxwriter

# ── Expiry codes ─────────────────────────────────────────────────────────────
EXPIRIES = [
    'M6','U6','Z6','H7','M7','U7','Z7','H8',
    'M8','U8','Z8','H9','M9','U9','Z9','H0',
    'M0','U0','Z0','H1',
]
LABELS = {
    'M6':'Jun-26','U6':'Sep-26','Z6':'Dec-26','H7':'Mar-27',
    'M7':'Jun-27','U7':'Sep-27','Z7':'Dec-27','H8':'Mar-28',
    'M8':'Jun-28','U8':'Sep-28','Z8':'Dec-28','H9':'Mar-29',
    'M9':'Jun-29','U9':'Sep-29','Z9':'Dec-29','H0':'Mar-30',
    'M0':'Jun-30','U0':'Sep-30','Z0':'Dec-30','H1':'Mar-31',
}
N = len(EXPIRIES)

# ── Product config ────────────────────────────────────────────────────────────
PRODUCTS = {
    'SFR': {
        'spread': lambda a, b: f'SFR{a}{b} Comdty',
        'fly':    lambda a, c: f'BSFR{a}{c} Comdty',
        'condor': lambda a, d: f'CSFR{a}{d} Comdty',
        'dark': '#1F497D', 'light': '#C5D9F1',
    },
    'SFI': {
        'spread': lambda a, b: f'SFI{a}{b} Comdty',
        'fly':    lambda a, c: f'BSFI{a}{c} Comdty',
        'condor': None,
        'dark': '#974706', 'light': '#FCE4D6',
    },
    'ER': {
        'spread': lambda a, b: f'ER{a}ER{b} Comdty',
        'fly':    lambda a, c: f'BER{a}{c} Comdty',
        'condor': None,
        'dark': '#375623', 'light': '#D9EAD3',
    },
    'IR': {
        'spread': lambda a, b: f'IR{a}{b} Comdty',
        'fly':    lambda a, c: f'BIR{a}{c} Comdty',
        'condor': None,
        'dark': '#5B0A9F', 'light': '#E8DAEF',
    },
    'BA': {
        'spread': lambda a, b: f'BA{a}{b} Comdty',
        'fly':    lambda a, c: f'BBA{a}{c} Comdty',
        'condor': None,
        'dark': '#C00000', 'light': '#FFE0E0',
    },
}

SPREAD_GAPS  = [(1,3),(2,6),(3,9),(4,12),(6,18),(8,24),(12,36)]
FLY_STEPS    = [(1,3),(2,6),(3,9),(4,12)]
CONDOR_STEPS = [(1,3),(2,6)]

# ── Build universe ────────────────────────────────────────────────────────────
all_rows = {}

for product, cfg in PRODUCTS.items():
    rows = []

    for gap, months in SPREAD_GAPS:
        for i in range(N):
            j = i + gap
            if j >= N: continue
            a, b = EXPIRIES[i], EXPIRIES[j]
            rows.append((product, f'Spread {months}m',
                         f'{LABELS[a]} / {LABELS[b]}',
                         cfg['spread'](a, b), f'{a}-{b}', ''))

    for n, months in FLY_STEPS:
        note = 'BBA prefix unconfirmed' if product == 'BA' else ''
        for i in range(N - 2*n):
            a, b, c = EXPIRIES[i], EXPIRIES[i+n], EXPIRIES[i+2*n]
            rows.append((product, f'Fly {months}m',
                         f'{LABELS[a]} / {LABELS[b]} / {LABELS[c]}',
                         cfg['fly'](a, c), f'{a}-{b}-{c}', note))

    if cfg['condor']:
        for n, months in CONDOR_STEPS:
            for i in range(N - 3*n):
                a, b, c, d = (EXPIRIES[i], EXPIRIES[i+n],
                              EXPIRIES[i+2*n], EXPIRIES[i+3*n])
                rows.append((product, f'Condor {months}m',
                             f'{LABELS[a]} / {LABELS[b]} / {LABELS[c]} / {LABELS[d]}',
                             cfg['condor'](a, d), f'{a}-{b}-{c}-{d}', ''))

    all_rows[product] = rows

# SFR packs/bundles
for exp in EXPIRIES[:16]:
    for tenor in ('1Y','2Y','3Y','4Y'):
        all_rows['SFR'].append(('SFR', f'Pack/Bundle {tenor}',
            f'SFR {tenor} from {LABELS[exp]}',
            f'SFR{tenor}{exp} Comdty', exp, ''))

# SFI bundles
for exp in EXPIRIES[:16]:
    for tenor in ('1Y','2Y','3Y'):
        all_rows['SFI'].append(('SFI', f'Bundle {tenor}',
            f'SFI {tenor} from {LABELS[exp]}',
            f'SFI{tenor}{exp} Comdty', exp, ''))

# ER basis
for exp in EXPIRIES[:16]:
    all_rows['ER'].append(('ER', 'Basis ESTR/EUR',
        f'ESTR vs Euribor {LABELS[exp]}',
        f'TKYER{exp} Comdty', exp, ''))

universe = []
for p in PRODUCTS:
    universe.extend(all_rows[p])

print(f'Total rows: {len(universe)}')
for p, rows in all_rows.items():
    print(f'  {p}: {len(rows)}')


# ── Write workbook ────────────────────────────────────────────────────────────
out = '/home/user/Future-curve-pipeline/Workbook/Most.Traded.Universe.xlsx'
wb  = xlsxwriter.Workbook(out, {'strings_to_formulas': False})

HEADERS = ['Product','Type','Description','Bloomberg Ticker',
           'Legs','Notes','Volume','Bid','Ask','Last Price']
COL_W   = [8, 16, 36, 26, 16, 28, 10, 9, 9, 11]

def add_universe_sheet(wb, title, rows, dark_hex, light_hex):
    """Write a sheet with headers + data rows + BDP formulas."""
    ws = wb.add_worksheet(title)

    # Formats
    hdr_fmt = wb.add_format({
        'bold': True, 'font_color': 'white',
        'bg_color': dark_hex, 'border': 1,
        'align': 'center', 'valign': 'vcenter',
        'font_size': 10,
    })
    num_fmt = wb.add_format({
        'bg_color': light_hex, 'border': 1,
        'align': 'right', 'valign': 'vcenter',
        'num_format': '#,##0',
    })
    txt_fmt = wb.add_format({
        'bg_color': light_hex, 'border': 1,
        'valign': 'vcenter',
    })

    # Headers
    for col, (h, w) in enumerate(zip(HEADERS, COL_W)):
        ws.write(0, col, h, hdr_fmt)
        ws.set_column(col, col, w)
    ws.set_row(0, 20)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(rows), len(HEADERS)-1)

    # Data rows
    for ri, row in enumerate(rows, 1):
        product = row[0]
        p_light = PRODUCTS[product]['light']
        p_dark  = PRODUCTS[product]['dark']

        row_txt = wb.add_format({
            'bg_color': p_light, 'border': 1, 'valign': 'vcenter'})
        row_num = wb.add_format({
            'bg_color': p_light, 'border': 1,
            'align': 'right', 'valign': 'vcenter',
            'num_format': '#,##0'})

        for ci, val in enumerate(row):
            ws.write(ri, ci, val, row_txt)

        ticker_cell = xlsxwriter.utility.xl_rowcol_to_cell(ri, 3)  # col D
        for ci, field in zip(range(6,10), ('VOLUME','PX_BID','PX_ASK','PX_LAST')):
            ws.write_formula(ri, ci, f'=BDP({ticker_cell},"{field}")', row_num)

    return ws


# ── Universe tab ─────────────────────────────────────────────────────────────
add_universe_sheet(wb, 'Universe', universe, '#1F497D', '#EBF1F8')


# ── Per-product ranked tabs ───────────────────────────────────────────────────
for product, rows in all_rows.items():
    dark  = PRODUCTS[product]['dark']
    light = PRODUCTS[product]['light']
    add_universe_sheet(wb, f'{product}', rows, dark, light)


# ── Notes tab ────────────────────────────────────────────────────────────────
ws_n = wb.add_worksheet('Notes')
ws_n.hide_gridlines(2)
ws_n.set_column(0, 0, 2)
ws_n.set_column(1, 1, 72)

title_fmt = wb.add_format({'bold':True,'font_size':14,'font_color':'#1F497D'})
head_fmt  = wb.add_format({'bold':True,'font_size':11,'font_color':'#1F497D'})
body_fmt  = wb.add_format({'font_size':10})

notes = [
    ('title', 'MOST TRADED UNIVERSE — SETUP & NOTES'),
    ('', ''),
    ('head', 'SETUP'),
    ('body', '1.  Open in Excel with Bloomberg Add-in active (Bloomberg must be running).'),
    ('body', '2.  BDP formulas refresh automatically when Bloomberg connects.'),
    ('body', '3.  Force refresh: Data → Refresh All, or press F9.'),
    ('body', '4.  To rank by volume: select any cell in Volume col → Data → Sort Z→A.'),
    ('', ''),
    ('head', 'TABS'),
    ('body', 'Universe  — All 5 products combined. Use auto-filter to slice by Product/Type.'),
    ('body', 'SFR       — SOFR structures only.'),
    ('body', 'SFI       — SONIA structures only.'),
    ('body', 'ER        — Euribor structures only.'),
    ('body', 'IR        — Australian Bank Bills only.'),
    ('body', 'BA        — Canadian BAs only.'),
    ('body', 'Sort any tab by Volume (col G) descending to see the most traded structures.'),
    ('', ''),
    ('head', 'TICKER FORMAT — CONFIRMED'),
    ('body', 'Spreads:  SFR/SFI/IR/BA: {prefix}{near}{far}  e.g. SFRM6U6'),
    ('body', '          ER: ER{near}ER{far}                  e.g. ERM6ERU6'),
    ('body', 'Flies:    B{prefix}{near_wing}{far_wing}        e.g. BSFRM6Z6'),
    ('body', '          (belly is implied; only near/far wings in ticker)'),
    ('body', 'Condors:  C{prefix}{first}{last}               e.g. CSFRM6H7'),
    ('body', '          SFR only — not listed for other products'),
    ('', ''),
    ('head', 'STRUCTURES'),
    ('body', 'Spreads:  3m / 6m / 9m / 12m / 18m / 24m / 36m tenors'),
    ('body', 'Flies:    equal-wing, spacing 3m / 6m / 9m / 12m'),
    ('body', 'Condors:  equal-wing, spacing 3m / 6m  (SFR only)'),
    ('body', 'Packs:    SFR 1Y/2Y/3Y/4Y,  SFI 1Y/2Y/3Y'),
    ('body', 'Basis:    ESTR vs Euribor (TKYER prefix, ER tab)'),
    ('body', 'BA fly:   BBA prefix assumed — verify BBAM6Z6 in Bloomberg'),
    ('', ''),
    ('body', f'Total tickers: {len(universe)}'),
]

for ri, (style, text) in enumerate(notes):
    if style == 'title':
        ws_n.write(ri, 1, text, title_fmt)
        ws_n.set_row(ri, 22)
    elif style == 'head':
        ws_n.write(ri, 1, text, head_fmt)
        ws_n.set_row(ri, 16)
    elif style == 'body':
        ws_n.write(ri, 1, text, body_fmt)
    else:
        ws_n.write(ri, 1, '', body_fmt)

# Move Notes to front by writing it last and reordering isn't possible in xlsxwriter
# — Notes will be rightmost tab; acceptable

wb.close()
print(f'Saved: {out}')
