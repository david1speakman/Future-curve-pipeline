"""
Most Traded Universe - Excel Generator
Creates Most.Traded.Universe.xlsx with all spread/fly/condor permutations
for SFR, SFI, ER, IR, BA futures, with Bloomberg BDP volume formulas.

Ticker conventions (confirmed from Curve sheet + user):
  Spreads:
    SFR/SFI/IR/BA: {prefix}{near}{far}   e.g. SFRM6U6
    ER:            ER{near}ER{far}        e.g. ERM6ERU6
  Flies:
    B{fly_prefix}{near_wing}{far_wing}    e.g. BSFRM6Z6 (belly implied)
    Prefixes: SFR→BSFR, SFI→BSFI, ER→BER, IR→BIR
  Condors:
    C{prefix}{first}{last}               e.g. CSFRM6M7
    USD products only (SFR confirmed; SFI/ER/IR/BA omitted)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Expiry codes (20 quarterly contracts, Jun-26 → Mar-31) ──────────────────
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

# ── Product config ───────────────────────────────────────────────────────────
# condor=None means no exchange-listed condor tickers for that product
PRODUCTS = {
    'SFR': {
        'spread': lambda a, b: f'SFR{a}{b} Comdty',
        'fly':    lambda a, c: f'BSFR{a}{c} Comdty',    # near wing + far wing
        'condor': lambda a, d: f'CSFR{a}{d} Comdty',    # first + last leg
        'dark': '1F497D', 'light': 'C5D9F1',
    },
    'SFI': {
        'spread': lambda a, b: f'SFI{a}{b} Comdty',
        'fly':    lambda a, c: f'BSFI{a}{c} Comdty',
        'condor': None,
        'dark': '974706', 'light': 'FCE4D6',
    },
    'ER': {
        'spread': lambda a, b: f'ER{a}ER{b} Comdty',
        'fly':    lambda a, c: f'BER{a}{c} Comdty',
        'condor': None,
        'dark': '375623', 'light': 'D9EAD3',
    },
    'IR': {
        'spread': lambda a, b: f'IR{a}{b} Comdty',
        'fly':    lambda a, c: f'BIR{a}{c} Comdty',
        'condor': None,
        'dark': '5B0A9F', 'light': 'E8DAEF',
    },
    'BA': {
        'spread': lambda a, b: f'BA{a}{b} Comdty',
        'fly':    lambda a, c: f'BBA{a}{c} Comdty',    # BBA prefix assumed
        'condor': None,
        'dark': 'C00000', 'light': 'FFE0E0',
    },
}

# ── Universe generation ──────────────────────────────────────────────────────
# Each row: [product, structure_type, description, ticker, legs, notes]

SPREAD_GAPS = [(1,3),(2,6),(3,9),(4,12),(6,18),(8,24),(12,36)]
FLY_STEPS   = [(1,3),(2,6),(3,9),(4,12)]   # n quarters per wing
CONDOR_STEPS = [(1,3),(2,6)]

all_rows = {}   # keyed by product

for product, cfg in PRODUCTS.items():
    rows = []

    # ── Spreads ──────────────────────────────────────────────────────────────
    for gap, months in SPREAD_GAPS:
        for i in range(N):
            j = i + gap
            if j >= N:
                continue
            a, b = EXPIRIES[i], EXPIRIES[j]
            rows.append([
                product,
                f'Spread {months}m',
                f'{LABELS[a]} / {LABELS[b]}',
                cfg['spread'](a, b),
                f'{a}-{b}',
                '',
            ])

    # ── Flies (equal-wing: near=a, belly=b, far=c; ticker uses a and c only) ─
    for n, months in FLY_STEPS:
        note_ba = 'BBA prefix unconfirmed' if product == 'BA' else ''
        note_cond = 'Condor ticker format unconfirmed' if product == 'BA' else ''
        for i in range(N - 2 * n):
            a, b, c = EXPIRIES[i], EXPIRIES[i + n], EXPIRIES[i + 2 * n]
            rows.append([
                product,
                f'Fly {months}m',
                f'{LABELS[a]} / {LABELS[b]} / {LABELS[c]}',
                cfg['fly'](a, c),          # correct: near + far only
                f'{a}-{b}-{c}',
                note_ba,
            ])

    # ── Condors (SFR only; C{prefix}{first}{last} confirmed by user) ─────────
    if cfg['condor'] is not None:
        for n, months in CONDOR_STEPS:
            for i in range(N - 3 * n):
                a, b, c, d = (EXPIRIES[i], EXPIRIES[i + n],
                              EXPIRIES[i + 2 * n], EXPIRIES[i + 3 * n])
                rows.append([
                    product,
                    f'Condor {months}m',
                    f'{LABELS[a]} / {LABELS[b]} / {LABELS[c]} / {LABELS[d]}',
                    cfg['condor'](a, d),    # first leg + last leg only
                    f'{a}-{b}-{c}-{d}',
                    '',
                ])

    all_rows[product] = rows

# ── SFR Packs & Bundles ──────────────────────────────────────────────────────
for exp in EXPIRIES[:16]:
    for tenor in ('1Y', '2Y', '3Y', '4Y'):
        all_rows['SFR'].append([
            'SFR', f'Pack/Bundle {tenor}',
            f'SFR {tenor} from {LABELS[exp]}',
            f'SFR{tenor}{exp} Comdty', exp, '',
        ])

# ── SFI Bundles ──────────────────────────────────────────────────────────────
for exp in EXPIRIES[:16]:
    for tenor in ('1Y', '2Y', '3Y'):
        all_rows['SFI'].append([
            'SFI', f'Bundle {tenor}',
            f'SFI {tenor} from {LABELS[exp]}',
            f'SFI{tenor}{exp} Comdty', exp, '',
        ])

# ── ER Intercommodity: ESTR vs Euribor ───────────────────────────────────────
for exp in EXPIRIES[:16]:
    all_rows['ER'].append([
        'ER', 'Basis ESTR/EUR',
        f'ESTR vs Euribor {LABELS[exp]}',
        f'TKYER{exp} Comdty', exp, '',
    ])

# Flat list for Universe tab
universe_rows = []
for product in PRODUCTS:
    universe_rows.extend(all_rows[product])

total = len(universe_rows)
print(f'Total universe rows: {total}')
for p, rows in all_rows.items():
    print(f'  {p}: {len(rows)} rows')


# ── Workbook helpers ─────────────────────────────────────────────────────────
wb = openpyxl.Workbook()

def hdr_cell(ws, row, col, value, bg='1F497D'):
    c = ws.cell(row, col, value)
    c.font = Font(bold=True, color='FFFFFF', size=10)
    c.fill = PatternFill('solid', fgColor=bg)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    return c

def thin_border():
    s = Side(style='thin', color='BFBFBF')
    return Border(left=s, right=s, top=s, bottom=s)

HEADERS = ['Product', 'Type', 'Description', 'Bloomberg Ticker',
           'Legs', 'Notes', 'Volume', 'Bid', 'Ask', 'Last Price']
COL_WIDTHS = [8, 16, 36, 26, 16, 30, 10, 9, 9, 11]


# ══════════════════════════════════════════════════════════════════════════════
# Tab: Universe  (all products combined)
# ══════════════════════════════════════════════════════════════════════════════
ws_u = wb.active
ws_u.title = 'Universe'

for col, h in enumerate(HEADERS, 1):
    hdr_cell(ws_u, 1, col, h)
ws_u.row_dimensions[1].height = 22

for ri, row_data in enumerate(universe_rows, 2):
    product = row_data[0]
    light = PRODUCTS[product]['light']
    for ci, val in enumerate(row_data, 1):
        c = ws_u.cell(ri, ci, val)
        c.fill = PatternFill('solid', fgColor=light)
        c.alignment = Alignment(vertical='center')
        c.border = thin_border()
    for ci, field in zip(range(7, 11), ('VOLUME', 'PX_BID', 'PX_ASK', 'PX_LAST')):
        c = ws_u.cell(ri, ci, f'=BDP(D{ri},"{field}")')
        c.fill = PatternFill('solid', fgColor=light)
        c.alignment = Alignment(horizontal='right', vertical='center')
        c.border = thin_border()

for col, width in zip(range(1, 11), COL_WIDTHS):
    ws_u.column_dimensions[get_column_letter(col)].width = width
ws_u.freeze_panes = 'A2'
ws_u.auto_filter.ref = f'A1:J{total + 1}'


# ══════════════════════════════════════════════════════════════════════════════
# Per-product ranked tabs  (one tab per currency)
# ══════════════════════════════════════════════════════════════════════════════
last_univ_row = total + 1   # +1 for header

for product in PRODUCTS:
    dark  = PRODUCTS[product]['dark']
    light = PRODUCTS[product]['light']
    ws = wb.create_sheet(f'{product} Ranked')

    for col, h in enumerate(HEADERS, 1):
        hdr_cell(ws, 1, col, h, bg=dark)
    ws.row_dimensions[1].height = 22

    # SORT(FILTER(...)) — shows only this product rows with volume > 0, sorted desc
    ws['A2'] = (
        f'=SORT('
        f'FILTER(Universe!A2:J{last_univ_row},'
        f'(Universe!A2:A{last_univ_row}="{product}")*'
        f'ISNUMBER(Universe!G2:G{last_univ_row})*'
        f'(Universe!G2:G{last_univ_row}>0)),'
        f'7,-1)'
    )

    for col, width in zip(range(1, 11), COL_WIDTHS):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = 'A2'

    # Instruction note in M1
    note = ws.cell(1, 12,
        f'{product}: sorted by volume, zero-vol rows hidden. '
        f'Requires Excel 365 SORT()+FILTER().')
    note.font = Font(italic=True, color='808080', size=9)


# ══════════════════════════════════════════════════════════════════════════════
# Tab: Notes
# ══════════════════════════════════════════════════════════════════════════════
ws_n = wb.create_sheet('Notes', 0)
ws_n.sheet_view.showGridLines = False
ws_n.column_dimensions['A'].width = 2
ws_n.column_dimensions['B'].width = 72

notes_lines = [
    ('title', 'MOST TRADED UNIVERSE — SETUP & NOTES'),
    ('', ''),
    ('head', 'SETUP'),
    ('body', '1.  Open in Excel with Bloomberg Add-in active (Bloomberg must be running).'),
    ('body', '2.  Allow BDP formulas to recalculate — may take 2-4 minutes for ~1,100 tickers.'),
    ('body', '3.  Refresh all: Data → Refresh All, or press F9.'),
    ('', ''),
    ('head', 'TABS'),
    ('body', 'Universe      — All permutations across all 5 products. Use auto-filter (row 1).'),
    ('body', 'SFR Ranked    — SOFR structures sorted by volume, zero-vol hidden.'),
    ('body', 'SFI Ranked    — SONIA structures sorted by volume.'),
    ('body', 'ER Ranked     — Euribor structures sorted by volume.'),
    ('body', 'IR Ranked     — Australian Bank Bills sorted by volume.'),
    ('body', 'BA Ranked     — Canadian BAs sorted by volume.'),
    ('body', 'All Ranked tabs use SORT()+FILTER() — requires Excel 365.'),
    ('', ''),
    ('head', 'PRODUCTS'),
    ('body', 'SFR  — 3-Month SOFR (CME)'),
    ('body', 'SFI  — 3-Month SONIA (ICE LIFFE)'),
    ('body', 'ER   — 3-Month Euribor (ICE LIFFE) + ESTR/Euribor basis (TKYER prefix)'),
    ('body', 'IR   — 3-Month Bank Bills, Australia (ASX)'),
    ('body', 'BA   — 3-Month Bankers Acceptances, Canada (MX)'),
    ('', ''),
    ('head', 'STRUCTURES INCLUDED'),
    ('body', 'Spreads      — all pairs, tenors: 3m/6m/9m/12m/18m/24m/36m'),
    ('body', 'Flies        — equal-wing butterflies, wing spacing: 3m/6m/9m/12m'),
    ('body', '               Ticker: B{prefix}{near_wing}{far_wing}  e.g. BSFRM6Z6 = Jun/Sep/Dec fly'),
    ('body', 'Condors      — equal-wing, spacing: 3m/6m (ticker format unconfirmed)'),
    ('body', 'Packs        — SFR 1Y/2Y/3Y/4Y, SFI 1Y/2Y/3Y'),
    ('body', 'Basis        — ESTR vs Euribor per expiry (ER only, TKYER prefix)'),
    ('', ''),
    ('head', 'TICKER FORMAT NOTES'),
    ('body', 'Spreads:  confirmed (SFRM6U6, ERM6ERU6, SFIM6U6, IRM6U6, BAM6U6)'),
    ('body', 'Flies:    confirmed SFR/SFI/ER/IR — B{prefix}{near_wing}{far_wing}'),
    ('body', '          e.g. BSFRM6Z6 = Jun/Sep/Dec fly (belly Sep implied)'),
    ('body', '          BA: BBA prefix assumed — verify BBAM6Z6 in Bloomberg'),
    ('body', 'Condors:  confirmed SFR only — C{prefix}{first}{last}'),
    ('body', '          e.g. CSFRM6M7 = Jun/Sep/Dec/Mar condor'),
    ('body', '          Not available for SFI/ER/IR/BA'),
    ('', ''),
    ('body', f'Total tickers in universe: {total}'),
]

for ri, (style, text) in enumerate(notes_lines, 1):
    c = ws_n.cell(ri, 2, text)
    if style == 'title':
        c.font = Font(bold=True, size=14, color='1F497D')
        ws_n.row_dimensions[ri].height = 22
    elif style == 'head':
        c.font = Font(bold=True, size=11, color='1F497D')
        ws_n.row_dimensions[ri].height = 16
    else:
        c.font = Font(size=10)


# ── Save ─────────────────────────────────────────────────────────────────────
out_path = '/home/user/Future-curve-pipeline/Workbook/Most.Traded.Universe.xlsx'
wb.save(out_path)
print(f'Saved: {out_path}')
