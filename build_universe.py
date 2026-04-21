"""
Most Traded Universe - Excel Generator
Creates Most.Traded.Universe.xlsx with all spread/fly/condor permutations
for SFR, SFI, ER, IR, COR futures, with Bloomberg BDP volume formulas
and per-currency Top 15 display tabs with sparklines.

Ticker conventions (confirmed):
  Spreads:  SFR/SFI/COR: {prefix}{near}{far}       e.g. SFRM6U6 Comdty
            ER/IR: {prefix}{near}{prefix}{far}       e.g. ERM6ERU6, IRM6IRU6
  Flies:    B{prefix}{near_wing}{far_wing}           e.g. BSFRM6Z6 Comdty
  Condors:  C{prefix}{first}{last}                   e.g. CSFRM6H7 Comdty (SFR only)
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
        'spread': lambda a, b: f'IR{a}IR{b} Comdty',   # interleaved like ER
        'fly':    lambda a, c: f'BIR{a}{c} Comdty',
        'condor': None,
        'dark': '#5B0A9F', 'light': '#E8DAEF',
    },
    'COR': {
        'spread': lambda a, b: f'COR{a}{b} Comdty',    # Canadian CORRA (was BA)
        'fly':    lambda a, c: f'BCOR{a}{c} Comdty',
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
        for i in range(N - 2*n):
            a, b, c = EXPIRIES[i], EXPIRIES[i+n], EXPIRIES[i+2*n]
            rows.append((product, f'Fly {months}m',
                         f'{LABELS[a]} / {LABELS[b]} / {LABELS[c]}',
                         cfg['fly'](a, c), f'{a}-{b}-{c}', ''))

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

# ER basis (ESTR vs Euribor)
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

    hdr_fmt = wb.add_format({
        'bold': True, 'font_color': 'white',
        'bg_color': dark_hex, 'border': 1,
        'align': 'center', 'valign': 'vcenter', 'font_size': 10,
    })
    num_fmt = wb.add_format({
        'bg_color': light_hex, 'border': 1,
        'align': 'right', 'valign': 'vcenter', 'num_format': '#,##0',
    })
    txt_fmt = wb.add_format({
        'bg_color': light_hex, 'border': 1, 'valign': 'vcenter',
    })

    for col, (h, w) in enumerate(zip(HEADERS, COL_W)):
        ws.write(0, col, h, hdr_fmt)
        ws.set_column(col, col, w)
    ws.set_row(0, 20)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(rows), len(HEADERS)-1)

    for ri, row in enumerate(rows, 1):
        product = row[0]
        p_light = PRODUCTS[product]['light']

        row_txt = wb.add_format({
            'bg_color': p_light, 'border': 1, 'valign': 'vcenter'})
        row_num = wb.add_format({
            'bg_color': p_light, 'border': 1,
            'align': 'right', 'valign': 'vcenter', 'num_format': '#,##0'})

        for ci, val in enumerate(row):
            ws.write(ri, ci, val, row_txt)

        ticker_cell = xlsxwriter.utility.xl_rowcol_to_cell(ri, 3)  # col D
        for ci, field in zip(range(6, 10), ('VOLUME','PX_BID','PX_ASK','PX_LAST')):
            ws.write_formula(ri, ci, f'=BDP({ticker_cell},"{field}")', row_num)

    return ws


# ── Display tab (Top 10 by volume, matching Fut.flow.curve secondary table) ───
def add_display_tab(wb, product, prod_tab_name, dark_hex, light_hex):
    """Top 10 display tab: Instrument|Last|Chg|Volume|Bid|Offer|Zscore|Sparkline|ROLL|Vol20d|V/Avg|Alert."""
    tab_name = f'{product} Top10'
    ws = wb.add_worksheet(tab_name)

    TITLE_ROW  = 0
    HDR_ROW    = 1
    DATA_START = 2
    NROWS      = 10

    # ── Visible column indices ─────────────────────────────────────────────
    CA = 0   # A: full Bloomberg ticker (hidden — drives BDP/BDH)
    CB = 1   # B: Instrument display name
    CC = 2   # C: Last
    CD = 3   # D: Chg  (= Last - BDH yesterday)
    CE = 4   # E: Volume
    CF = 5   # F: Bid
    CG = 6   # G: Offer
    CH = 7   # H: Zscore 20d
    CI = 8   # I: 20d moves (sparkline)
    CJ = 9   # J: ROLL
    CK = 10  # K: Volume 20d Av.
    CL = 11  # L: V/Avg
    CM = 12  # M: Alert (BUY / SELL signal)
    # ── Stat + history columns — narrow but NOT hidden so sparklines render ──
    CAV  = 13  # N: Av. 20d
    CSD  = 14  # O: StDev 20d
    CHI0 = 15  # P: BDH history start  (20 cols: P=15 … AI=34)
    CHI1 = 34  # AI: BDH history end

    # ── Column widths ──────────────────────────────────────────────────────
    ws.set_column(CA,  CA,  None, None, {'hidden': True})  # A: hidden ticker
    ws.set_column(CB,  CB,  14)
    ws.set_column(CC,  CC,  10)
    ws.set_column(CD,  CD,  10)
    ws.set_column(CE,  CE,  10)
    ws.set_column(CF,  CF,   9)
    ws.set_column(CG,  CG,   9)
    ws.set_column(CH,  CH,  11)
    ws.set_column(CI,  CI,  22)
    ws.set_column(CJ,  CJ,  10)
    ws.set_column(CK,  CK,  14)
    ws.set_column(CL,  CL,   9)
    ws.set_column(CM,  CM,   8)
    ws.set_column(CAV, CHI1, 3)   # narrow but visible — Excel sparklines skip hidden cols

    # ── Formats ───────────────────────────────────────────────────────────
    title_fmt = wb.add_format({
        'bold': True, 'font_size': 12, 'font_color': 'white',
        'bg_color': dark_hex, 'align': 'left', 'valign': 'vcenter',
    })
    hdr_fmt = wb.add_format({
        'bold': True, 'font_color': 'white', 'bg_color': '#2F2F2F',
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 10,
    })
    date_hdr_fmt = wb.add_format({'num_format': 'dd-mmm', 'font_size': 7})

    cf_grn = wb.add_format({'bg_color': '#00B050', 'font_color': 'white',
                             'bold': True, 'border': 1, 'align': 'right',
                             'num_format': '0.00', 'font_size': 10})
    cf_red = wb.add_format({'bg_color': '#FF0000', 'font_color': 'white',
                             'bold': True, 'border': 1, 'align': 'right',
                             'num_format': '0.00', 'font_size': 10})
    alert_buy_fmt  = wb.add_format({'bg_color': '#00B050', 'font_color': 'white',
                                    'bold': True, 'border': 1, 'align': 'center',
                                    'font_size': 10})
    alert_sell_fmt = wb.add_format({'bg_color': '#FF0000', 'font_color': 'white',
                                    'bold': True, 'border': 1, 'align': 'center',
                                    'font_size': 10})

    ROW_BG = ['#FFFFFF', '#F5F5F5']
    fmts = {}
    for bg in ROW_BG:
        fmts[bg] = {
            'txt':  wb.add_format({'bg_color': bg, 'border': 1,
                                   'valign': 'vcenter', 'font_size': 10}),
            'bold': wb.add_format({'bg_color': bg, 'border': 1, 'bold': True,
                                   'valign': 'vcenter', 'font_size': 10}),
            'num':  wb.add_format({'bg_color': bg, 'border': 1, 'align': 'right',
                                   'valign': 'vcenter', 'num_format': '#,##0',
                                   'font_size': 10}),
            'px3':  wb.add_format({'bg_color': bg, 'border': 1, 'align': 'right',
                                   'valign': 'vcenter', 'num_format': '0.000',
                                   'font_size': 10}),
            'px4':  wb.add_format({'bg_color': bg, 'border': 1, 'align': 'right',
                                   'valign': 'vcenter', 'num_format': '0.0000',
                                   'font_size': 10}),
            'zsc':  wb.add_format({'bg_color': bg, 'border': 1, 'align': 'right',
                                   'valign': 'vcenter', 'num_format': '0.00',
                                   'font_size': 10}),
            'rat':  wb.add_format({'bg_color': bg, 'border': 1, 'align': 'right',
                                   'valign': 'vcenter', 'num_format': '0.0"x"',
                                   'font_size': 10}),
            'spk':  wb.add_format({'bg_color': bg, 'border': 1, 'valign': 'vcenter'}),
            'alrt': wb.add_format({'bg_color': bg, 'border': 1, 'align': 'center',
                                   'valign': 'vcenter', 'font_size': 10}),
        }

    note_fmt = wb.add_format({'italic': True, 'font_size': 9, 'font_color': '#808080'})

    # ── Row 0: product-coloured title bar ─────────────────────────────────
    ws.merge_range(TITLE_ROW, CA, TITLE_ROW, CM,
                   f'{product} — Most Traded (Top 10)', title_fmt)
    ws.set_row(TITLE_ROW, 18)

    # ── Row 1: header ─────────────────────────────────────────────────────
    for ci, label in {CB: 'Instrument', CC: 'Last', CD: 'Chg',   CE: 'Volume',
                      CF: 'Bid',        CG: 'Offer', CH: 'Zscore 20d',
                      CI: '20d moves',  CJ: 'ROLL',  CK: 'Volume 20d Av.',
                      CL: 'V/Avg',     CM: 'Alert'}.items():
        ws.write(HDR_ROW, ci, label, hdr_fmt)
    ws.set_row(HDR_ROW, 18)

    for offset in range(20):
        ws.write_formula(HDR_ROW, CHI0 + offset,
                         f'=WORKDAY(TODAY(),-{20 - offset})', date_hdr_fmt)

    ws.freeze_panes(DATA_START, 0)

    # ── Data rows 2..11 ───────────────────────────────────────────────────
    for rank in range(NROWS):
        row = DATA_START + rank
        r1  = row + 1
        bg  = ROW_BG[rank % 2]
        f   = fmts[bg]

        hi0 = xlsxwriter.utility.xl_rowcol_to_cell(row, CHI0)
        hi1 = xlsxwriter.utility.xl_rowcol_to_cell(row, CHI1)

        ws.write_formula(row, CA,
            f'=IFERROR(INDEX({prod_tab_name}!$D:$D,'
            f'MATCH(LARGE({prod_tab_name}!$G:$G,{rank+1}),'
            f'{prod_tab_name}!$G:$G,0)),"")')

        a = f'A{r1}'

        ws.write_formula(row, CB,
            f'=IFERROR(SUBSTITUTE({a}," Comdty",""),"")', f['bold'])

        ws.write_formula(row, CC, f'=IFERROR(BDP({a},"PX_LAST"),"")',        f['px3'])
        ws.write_formula(row, CD, f'=IFERROR(C{r1}-{hi1},"")',               f['px4'])
        ws.write_formula(row, CE, f'=IFERROR(BDP({a},"VOLUME"),"")',          f['num'])
        ws.write_formula(row, CF, f'=IFERROR(BDP({a},"PX_BID"),"")',          f['px3'])
        ws.write_formula(row, CG, f'=IFERROR(BDP({a},"PX_ASK"),"")',          f['px3'])
        ws.write_formula(row, CJ,
            f'=IFERROR(VALUE(BDP({a},"ROLL_DOWN_VALUE")),"")',                 f['px3'])
        ws.write_formula(row, CK, f'=IFERROR(BDP({a},"VOLUME_AVG_20D"),"")', f['num'])
        ws.write_formula(row, CL, f'=IFERROR(E{r1}/K{r1},"")',               f['rat'])

        # Alert: BUY = negative z-score with positive roll; SELL = positive z with negative roll
        ws.write_formula(row, CM,
            f'=IFERROR(IF(AND(H{r1}<0,J{r1}>0),"BUY",'
            f'IF(AND(H{r1}>0,J{r1}<0),"SELL","")),"")',                       f['alrt'])

        # BDH 20-day price history (P=15 … AI=34) — narrow cols, not hidden
        for offset in range(20):
            days_back = 20 - offset
            ws.write_formula(row, CHI0 + offset,
                f'=IFERROR(BDH({a},"PX_LAST",'
                f'WORKDAY(TODAY(),-{days_back}),'
                f'WORKDAY(TODAY(),-{days_back})),"")')

        ws.write_formula(row, CAV, f'=IFERROR(AVERAGE({hi0}:{hi1}),"")')
        ws.write_formula(row, CSD, f'=IFERROR(STDEV({hi0}:{hi1}),"")')

        av = xlsxwriter.utility.xl_rowcol_to_cell(row, CAV)
        sd = xlsxwriter.utility.xl_rowcol_to_cell(row, CSD)
        ws.write_formula(row, CH,
            f'=IFERROR(IF({sd}=0,"",(C{r1}-{av})/{sd}),"")', f['zsc'])

        ws.write_blank(row, CI, f['spk'])
        ws.add_sparkline(row, CI, {
            'range':        f"'{tab_name}'!{hi0}:{hi1}",
            'type':         'line',
            'weight':       1.5,
            'series_color': dark_hex,
            'high_point':   True,
            'low_point':    True,
            'high_color':   '#00B050',
            'low_color':    '#FF0000',
        })

        ws.set_row(row, 18)

    # ── Z-score conditional format: two-tier ──────────────────────────────
    r0, r1_ = DATA_START, DATA_START + NROWS - 1
    ws.conditional_format(r0, CH, r1_, CH, {
        'type': 'cell', 'criteria': '>=', 'value': 1, 'format': cf_grn})
    ws.conditional_format(r0, CH, r1_, CH, {
        'type': 'cell', 'criteria': '<=', 'value': -1, 'format': cf_red})
    ws.conditional_format(r0, CH, r1_, CH, {
        'type':     '3_color_scale',
        'min_type': 'num', 'min_value': -1, 'min_color': '#FF6666',
        'mid_type': 'num', 'mid_value':  0, 'mid_color': '#FFFF00',
        'max_type': 'num', 'max_value':  1, 'max_color': '#92D050',
    })

    # ── Alert conditional format ───────────────────────────────────────────
    ws.conditional_format(r0, CM, r1_, CM, {
        'type': 'text', 'criteria': 'containing', 'value': 'BUY',
        'format': alert_buy_fmt})
    ws.conditional_format(r0, CM, r1_, CM, {
        'type': 'text', 'criteria': 'containing', 'value': 'SELL',
        'format': alert_sell_fmt})

    ws.write(DATA_START + NROWS + 1, CB,
             f'Sort [{prod_tab_name}] by Volume (col G) desc to rank.  '
             f'Alert: BUY = z<0 & roll>0 (cheap & carry+);  SELL = z>0 & roll<0.', note_fmt)

    return ws


# ── Universe tab ─────────────────────────────────────────────────────────────
add_universe_sheet(wb, 'Universe', universe, '#1F497D', '#EBF1F8')

# ── Per-product: data tab then Top15 display tab, interleaved ────────────────
for product, rows in all_rows.items():
    add_universe_sheet(wb, product, rows,
                       PRODUCTS[product]['dark'], PRODUCTS[product]['light'])
    add_display_tab(wb, product, product,
                    PRODUCTS[product]['dark'], PRODUCTS[product]['light'])


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
    ('body', '4.  To rank by volume on a data tab: select any cell in Volume col → Data → Sort Z→A.'),
    ('body', '5.  The Top15 display tabs auto-update once you sort the matching data tab by Volume.'),
    ('', ''),
    ('head', 'TABS'),
    ('body', 'Universe      — All 5 products combined. Use auto-filter to slice by Product/Type.'),
    ('body', 'SFR           — SOFR structures only (data).'),
    ('body', 'SFI           — SONIA structures only (data).'),
    ('body', 'ER            — Euribor structures only (data).'),
    ('body', 'IR            — Australian Bank Bills only (data).'),
    ('body', 'COR           — Canadian CORRA only (data).'),
    ('body', 'SFR Top10     — Most traded SOFR structures with sparklines, z-scores + alerts.'),
    ('body', 'SFI Top10     — Most traded SONIA structures with sparklines, z-scores + alerts.'),
    ('body', 'ER Top10      — Most traded Euribor structures with sparklines, z-scores + alerts.'),
    ('body', 'IR Top10      — Most traded AUD Bank Bill structures with sparklines, z-scores + alerts.'),
    ('body', 'COR Top10     — Most traded CORRA structures with sparklines, z-scores + alerts.'),
    ('', ''),
    ('head', 'TICKER FORMAT — CONFIRMED'),
    ('body', 'Spreads:  SFR/SFI/COR: {prefix}{near}{far}      e.g. SFRM6U6 Comdty'),
    ('body', '          ER/IR: {prefix}{near}{prefix}{far}     e.g. ERM6ERU6, IRM6IRU6 Comdty'),
    ('body', 'Flies:    B{prefix}{near_wing}{far_wing}         e.g. BSFRM6Z6 Comdty'),
    ('body', '          (belly is implied; only near/far wings in ticker)'),
    ('body', 'Condors:  C{prefix}{first}{last}                 e.g. CSFRM6H7 Comdty'),
    ('body', '          SFR only — not listed for other products'),
    ('', ''),
    ('head', 'STRUCTURES'),
    ('body', 'Spreads:  3m / 6m / 9m / 12m / 18m / 24m / 36m tenors'),
    ('body', 'Flies:    equal-wing, spacing 3m / 6m / 9m / 12m'),
    ('body', 'Condors:  equal-wing, spacing 3m / 6m  (SFR only)'),
    ('body', 'Packs:    SFR 1Y/2Y/3Y/4Y'),
    ('body', 'Bundles:  SFI 1Y/2Y/3Y'),
    ('body', 'Basis:    ESTR vs Euribor (TKYER prefix, ER tab)'),
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

wb.close()
print(f'Saved: {out}')
