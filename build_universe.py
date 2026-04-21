"""
Most Traded Universe - Excel Generator
Creates Most.Traded.Universe.xlsx with all spread/fly/condor permutations
for SFR, SFI, ER, IR, BA futures, with Bloomberg BDP volume formulas.
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

# ── Product colours (dark header / light row) ────────────────────────────────
COLORS = {
    'SFR': ('1F497D', 'C5D9F1'),   # blue
    'SFI': ('974706', 'FCE4D6'),   # orange
    'ER':  ('375623', 'D9EAD3'),   # green
    'IR':  ('5B0A9F', 'E8DAEF'),   # purple
    'BA':  ('C00000', 'FFE0E0'),   # red
}

# ── Ticker builders ──────────────────────────────────────────────────────────
def make_ticker(product, legs):
    """
    Constructs Bloomberg strategy ticker.
    SFR/SFI/IR/BA: prefix + legs concatenated (e.g. SFRM6U6)
    ER:            ER interleaved  (e.g. ERM6ERU6ERZ6)
    """
    if product == 'ER':
        return 'ER' + 'ER'.join(legs) + ' Comdty'
    return product + ''.join(legs) + ' Comdty'


# ── Universe generation ──────────────────────────────────────────────────────
# Each row: [product, structure_type, description, ticker, legs, notes]
rows = []

SPREAD_GAPS = {1: 3, 2: 6, 3: 9, 4: 12, 6: 18, 8: 24, 12: 36}
FLY_STEPS   = [(1, 3), (2, 6), (3, 9), (4, 12)]
CONDOR_STEPS = [(1, 3), (2, 6)]

for product in ['SFR', 'SFI', 'ER', 'IR', 'BA']:

    # ── Spreads ──────────────────────────────────────────────────────────────
    for gap, months in SPREAD_GAPS.items():
        for i in range(N):
            j = i + gap
            if j >= N:
                continue
            a, b = EXPIRIES[i], EXPIRIES[j]
            rows.append([
                product,
                f'Spread {months}m',
                f'{LABELS[a]} / {LABELS[b]}',
                make_ticker(product, [a, b]),
                f'{a}-{b}',
                '',
            ])

    # ── Flies (equal-wing) ───────────────────────────────────────────────────
    for n, months in FLY_STEPS:
        for i in range(N - 2 * n):
            a, b, c = EXPIRIES[i], EXPIRIES[i + n], EXPIRIES[i + 2 * n]
            note = ('Fly ticker format unconfirmed for this product'
                    if product in ('IR', 'BA') else '')
            rows.append([
                product,
                f'Fly {months}m',
                f'{LABELS[a]} / {LABELS[b]} / {LABELS[c]}',
                make_ticker(product, [a, b, c]),
                f'{a}-{b}-{c}',
                note,
            ])

    # ── Condors (equal-wing) ─────────────────────────────────────────────────
    for n, months in CONDOR_STEPS:
        for i in range(N - 3 * n):
            a, b, c, d = (EXPIRIES[i], EXPIRIES[i + n],
                          EXPIRIES[i + 2 * n], EXPIRIES[i + 3 * n])
            rows.append([
                product,
                f'Condor {months}m',
                f'{LABELS[a]} / {LABELS[b]} / {LABELS[c]} / {LABELS[d]}',
                make_ticker(product, [a, b, c, d]),
                f'{a}-{b}-{c}-{d}',
                'Condor ticker format unconfirmed',
            ])

# ── SFR Packs & Bundles ──────────────────────────────────────────────────────
for exp in EXPIRIES[:16]:
    for tenor in ('1Y', '2Y', '3Y', '4Y'):
        rows.append([
            'SFR',
            f'Pack/Bundle {tenor}',
            f'SFR {tenor} from {LABELS[exp]}',
            f'SFR{tenor}{exp} Comdty',
            exp,
            '',
        ])

# ── SFI Bundles ──────────────────────────────────────────────────────────────
for exp in EXPIRIES[:16]:
    for tenor in ('1Y', '2Y', '3Y'):
        rows.append([
            'SFI',
            f'Bundle {tenor}',
            f'SFI {tenor} from {LABELS[exp]}',
            f'SFI{tenor}{exp} Comdty',
            exp,
            '',
        ])

# ── ER Intercommodity: ESTR vs Euribor (TKYER prefix) ───────────────────────
for exp in EXPIRIES[:16]:
    rows.append([
        'ER',
        'Basis ESTR/EUR',
        f'ESTR vs Euribor {LABELS[exp]}',
        f'TKYER{exp} Comdty',
        exp,
        '',
    ])

print(f'Total universe rows: {len(rows)}')


# ── Workbook construction ────────────────────────────────────────────────────
wb = openpyxl.Workbook()

# ── Helpers ──────────────────────────────────────────────────────────────────
def hdr_cell(ws, row, col, value, bg='1F497D'):
    c = ws.cell(row, col, value)
    c.font = Font(bold=True, color='FFFFFF', size=10)
    c.fill = PatternFill('solid', fgColor=bg)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    return c

def thin_border():
    s = Side(style='thin', color='BFBFBF')
    return Border(left=s, right=s, top=s, bottom=s)


# ══════════════════════════════════════════════════════════════════════════════
# Tab 1: Universe
# ══════════════════════════════════════════════════════════════════════════════
ws_u = wb.active
ws_u.title = 'Universe'

HEADERS = ['Product', 'Type', 'Description', 'Bloomberg Ticker',
           'Legs', 'Notes', 'Volume', 'Bid', 'Ask', 'Last Price']

for col, h in enumerate(HEADERS, 1):
    hdr_cell(ws_u, 1, col, h)

ws_u.row_dimensions[1].height = 22

for ri, row_data in enumerate(rows, 2):
    product = row_data[0]
    _, light = COLORS[product]

    for ci, val in enumerate(row_data, 1):
        c = ws_u.cell(ri, ci, val)
        c.fill = PatternFill('solid', fgColor=light)
        c.alignment = Alignment(vertical='center', wrap_text=False)
        c.border = thin_border()

    # BDP formulas — columns G/H/I/J (7-10)
    for ci, field in zip(range(7, 11), ('VOLUME', 'PX_BID', 'PX_ASK', 'PX_LAST')):
        c = ws_u.cell(ri, ci, f'=BDP(D{ri},"{field}")')
        c.fill = PatternFill('solid', fgColor=light)
        c.alignment = Alignment(horizontal='right', vertical='center')
        c.border = thin_border()

# Column widths
for col, width in zip(range(1, 11), [8, 16, 36, 30, 16, 36, 10, 9, 9, 11]):
    ws_u.column_dimensions[get_column_letter(col)].width = width

ws_u.freeze_panes = 'A2'

# Auto-filter
ws_u.auto_filter.ref = f'A1:J{len(rows) + 1}'


# ══════════════════════════════════════════════════════════════════════════════
# Tab 2: Rankings  (SORT spill — requires Excel 365)
# ══════════════════════════════════════════════════════════════════════════════
ws_r = wb.create_sheet('Rankings')

for col, h in enumerate(HEADERS, 1):
    hdr_cell(ws_r, 1, col, h)
ws_r.row_dimensions[1].height = 22

# FILTER to remove zero/N/A volume rows, then SORT descending by volume
last_data_row = len(rows) + 1
ws_r['A2'] = (
    f'=SORT('
    f'FILTER(Universe!A2:J{last_data_row},'
    f'ISNUMBER(Universe!G2:G{last_data_row})*'
    f'(Universe!G2:G{last_data_row}>0)),'
    f'7,-1)'
)

# Column widths (same as Universe)
for col, width in zip(range(1, 11), [8, 16, 36, 30, 16, 36, 10, 9, 9, 11]):
    ws_r.column_dimensions[get_column_letter(col)].width = width

ws_r.freeze_panes = 'A2'

# Note about conditional formatting (can't apply to spill range in openpyxl)
note = ws_r.cell(1, 12, '← Spill formula: sorted by Volume, zero-volume rows hidden')
note.font = Font(italic=True, color='808080', size=9)


# ══════════════════════════════════════════════════════════════════════════════
# Tab 3: By_Product  (one ranking per product using separate SORT blocks)
# ══════════════════════════════════════════════════════════════════════════════
ws_p = wb.create_sheet('By_Product')

col_offset = 1
gap_cols   = 1  # blank column between product blocks

for product in ['SFR', 'SFI', 'ER', 'IR', 'BA']:
    dark, light = COLORS[product]

    # Product header
    hdr = ws_p.cell(1, col_offset, product)
    hdr.font = Font(bold=True, size=12, color='FFFFFF')
    hdr.fill = PatternFill('solid', fgColor=dark)
    hdr.alignment = Alignment(horizontal='center', vertical='center')
    ws_p.merge_cells(
        start_row=1, start_column=col_offset,
        end_row=1,   end_column=col_offset + 3
    )

    # Sub-headers: Type | Description | Ticker | Volume
    for ci, h in enumerate(['Type', 'Description', 'Ticker', 'Volume'], col_offset):
        c = ws_p.cell(2, ci, h)
        c.font = Font(bold=True, color='FFFFFF', size=9)
        c.fill = PatternFill('solid', fgColor=dark)
        c.alignment = Alignment(horizontal='center', vertical='center')

    # FILTER+SORT formula scoped to this product
    ws_p.cell(
        3, col_offset,
        f'=SORT('
        f'FILTER('
        f'CHOOSE({{1,2,3,4}},Universe!B2:B{last_data_row},Universe!C2:C{last_data_row},'
        f'Universe!D2:D{last_data_row},Universe!G2:G{last_data_row}),'
        f'(Universe!A2:A{last_data_row}="{product}")*'
        f'ISNUMBER(Universe!G2:G{last_data_row})*'
        f'(Universe!G2:G{last_data_row}>0)),'
        f'4,-1)'
    )

    # Column widths for this block
    for ci, w in zip(range(col_offset, col_offset + 4), [16, 28, 24, 10]):
        ws_p.column_dimensions[get_column_letter(ci)].width = w

    col_offset += 4 + gap_cols

ws_p.row_dimensions[1].height = 20
ws_p.row_dimensions[2].height = 18
ws_p.freeze_panes = 'A3'


# ══════════════════════════════════════════════════════════════════════════════
# Tab 0: Notes (inserted at front)
# ══════════════════════════════════════════════════════════════════════════════
ws_n = wb.create_sheet('Notes', 0)
ws_n.sheet_view.showGridLines = False
ws_n.column_dimensions['A'].width = 2
ws_n.column_dimensions['B'].width = 70

notes = [
    ('', ''),
    ('', 'MOST TRADED UNIVERSE — SETUP & NOTES'),
    ('', ''),
    ('', 'SETUP'),
    ('', '1.  Open in Excel with Bloomberg Add-in active (Bloomberg must be running).'),
    ('', '2.  Allow BDP formulas to recalculate — this may take 1-3 minutes for ~1,100 tickers.'),
    ('', '3.  Refresh: Data → Refresh All, or F9.'),
    ('', ''),
    ('', 'TABS'),
    ('', 'Universe    — Full list of all permutations with live BDP volume/bid/ask/last.'),
    ('', '               Use the Auto-Filter (row 1) to filter by Product or Type.'),
    ('', 'Rankings    — All tickers sorted by Volume descending, zero-volume rows removed.'),
    ('', '               Uses SORT()+FILTER() — requires Excel 365.'),
    ('', 'By_Product  — Side-by-side top-volume tables per product.'),
    ('', ''),
    ('', 'PRODUCTS'),
    ('', 'SFR  — 3-Month SOFR (CME)'),
    ('', 'SFI  — 3-Month SONIA (ICE LIFFE)'),
    ('', 'ER   — 3-Month Euribor (ICE LIFFE) + ESTR/Euribor basis (TKYER prefix)'),
    ('', 'IR   — 3-Month Bank Bills, Australia (ASX)'),
    ('', 'BA   — 3-Month Bankers Acceptances, Canada (MX)'),
    ('', ''),
    ('', 'STRUCTURES INCLUDED'),
    ('', 'Spreads   — all pairs, tenors: 3m/6m/9m/12m/18m/24m/36m'),
    ('', 'Flies     — equal-wing butterflies, wing spacing: 3m/6m/9m/12m'),
    ('', 'Condors   — equal-wing, spacing: 3m/6m'),
    ('', 'Packs     — SFR 1Y/2Y/3Y/4Y, SFI 1Y/2Y/3Y'),
    ('', 'Basis     — ESTR vs Euribor per expiry (ER only)'),
    ('', ''),
    ('', 'TICKER FORMAT NOTES'),
    ('', 'Spreads:  confirmed from existing flow data (e.g. SFRM6U6, ERM6ERU6, SFIM6U6)'),
    ('', 'Flies:    inferred convention (e.g. SFRM6U6Z6) — confirm first occurrence in BBG'),
    ('', 'Condors:  inferred convention (e.g. SFRM6U6Z6H7) — confirm in BBG'),
    ('', 'IR/BA flies & condors: especially uncertain — flag in Notes column'),
    ('', ''),
    ('', f'Total tickers in universe: {len(rows)}'),
]

for ri, (_, text) in enumerate(notes, 1):
    c = ws_n.cell(ri, 2, text)
    if 'SETUP' == text or 'TABS' == text or 'PRODUCTS' == text or \
       'STRUCTURES INCLUDED' == text or 'TICKER FORMAT NOTES' == text:
        c.font = Font(bold=True, size=11, color='1F497D')
    elif text.startswith('MOST TRADED'):
        c.font = Font(bold=True, size=14, color='1F497D')
    else:
        c.font = Font(size=10)

ws_n.row_dimensions[2].height = 24


# ── Save ─────────────────────────────────────────────────────────────────────
out_path = '/home/user/Future-curve-pipeline/Workbook/Most.Traded.Universe.xlsx'
wb.save(out_path)
print(f'Saved: {out_path}')
