#!/usr/bin/env python3
"""
London Morning Email Pipeline
==============================
Pulls Top 10 most-traded rate structures for selected currencies,
renders each as a formatted table image (matching the Excel Top 10 tabs),
and assembles an HTML email draft with embedded images.

Usage:
    python morning_email.py SFR ER SFI          # specific currencies
    python morning_email.py --all               # SFR + SFI + ER + IR + COR
    python morning_email.py SFR --no-bloomberg  # demo/test mode, no Bloomberg needed
    python morning_email.py SFR --out ./drafts  # custom output folder
    python morning_email.py SFR --ib-paste      # try to auto-paste into Bloomberg IB

Outputs (./output/ by default):
    morning_YYYYMMDD.html       open in Outlook or any browser to copy-paste
    morning_YYYYMMDD_ib.txt     plain-text table for Bloomberg IB
    SFR_top10.png / ER_top10.png / ...  individual images

Bloomberg IB notes: see paste_to_bloomberg_ib() at the bottom of this file.
"""

import argparse
import base64
import io
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# ── Product config ────────────────────────────────────────────────────────────
PRODUCTS = {
    'SFR': {'dark': '#1F497D', 'light': '#C5D9F1', 'label': 'SOFR (CME)'},
    'SFI': {'dark': '#974706', 'light': '#FCE4D6', 'label': 'SONIA (ICE)'},
    'ER':  {'dark': '#375623', 'light': '#D9EAD3', 'label': 'Euribor (ICE)'},
    'IR':  {'dark': '#5B0A9F', 'light': '#E8DAEF', 'label': 'BBSW / AUD (ASX)'},
    'COR': {'dark': '#C00000', 'light': '#FFE0E0', 'label': 'CORRA (TMX)'},
}

EXPIRIES = ['M6','U6','Z6','H7','M7','U7','Z7','H8',
            'M8','U8','Z8','H9','M9','U9','Z9','H0',
            'M0','U0','Z0','H1']
N = len(EXPIRIES)
SPREAD_GAPS  = [(1,3),(2,6),(3,9),(4,12),(6,18),(8,24),(12,36)]
FLY_STEPS    = [(1,3),(2,6),(3,9),(4,12)]
CONDOR_STEPS = [(1,3),(2,6)]

# Table column layout (must match the Excel Top 10 tabs)
COLS  = ['Instrument', 'Last', 'Chg', 'Volume', 'Bid', 'Offer',
         'Z-score\n20d', '20d moves', 'ROLL', 'Vol 20d\nAv.', 'V/Avg', 'Alert']
COL_W = [2.8, 0.90, 0.90, 1.10, 0.90, 0.90, 0.95, 1.80, 0.90, 1.20, 0.70, 0.70]
SPK_IDX = 7   # which column index is the sparkline

# ── Value formatting helpers ──────────────────────────────────────────────────
def _fmt(val, kind):
    if val is None or val == '':
        return ''
    try:
        if kind == 'px3':  return f'{float(val):.3f}'
        if kind == 'px4':  return f'{float(val):+.4f}'
        if kind == 'vol':  return f'{int(val):,}'
        if kind == 'zsc':  return f'{float(val):.2f}'
        if kind == 'rat':  return f'{float(val):.1f}x'
    except Exception:
        return ''
    return str(val)

def _lerp_hex(c1, c2, t):
    """Interpolate between two hex colours; t=0 → c1, t=1 → c2."""
    t = max(0.0, min(1.0, float(t)))
    r1,g1,b1 = int(c1[1:3],16), int(c1[3:5],16), int(c1[5:7],16)
    r2,g2,b2 = int(c2[1:3],16), int(c2[3:5],16), int(c2[5:7],16)
    return '#{:02X}{:02X}{:02X}'.format(
        round(r1+t*(r2-r1)), round(g1+t*(g2-g1)), round(b1+t*(b2-b1)))

def _zscore_style(z):
    """Returns (bg_hex, font_color) matching the Excel two-tier conditional format."""
    if z is None:
        return None, 'black'
    if z >= 1:
        return '#00B050', 'white'
    if z <= -1:
        return '#FF0000', 'white'
    if z >= 0:
        return _lerp_hex('#FFFF00', '#92D050', z), 'black'
    return _lerp_hex('#FF6666', '#FFFF00', z + 1), 'black'

def _alert_style(alert):
    if alert == 'BUY':  return '#00B050', 'white'
    if alert == 'SELL': return '#FF0000', 'white'
    return None, 'black'


# ── Ticker generation (mirrors build_universe.py) ─────────────────────────────
def get_tickers(product):
    if product in ('SFR', 'SFI', 'COR'):
        spread = lambda a, b: f'{product}{a}{b} Comdty'
        fly    = lambda a, c: f'B{product}{a}{c} Comdty'
        condor = (lambda a, d: f'C{product}{a}{d} Comdty') if product == 'SFR' else None
    elif product in ('ER', 'IR'):
        spread = lambda a, b: f'{product}{a}{product}{b} Comdty'
        fly    = lambda a, c: f'B{product}{a}{c} Comdty'
        condor = None
    else:
        return []

    out = []
    for g, _ in SPREAD_GAPS:
        for i in range(N):
            if i + g < N:
                out.append(spread(EXPIRIES[i], EXPIRIES[i + g]))
    for n, _ in FLY_STEPS:
        for i in range(N - 2*n):
            out.append(fly(EXPIRIES[i], EXPIRIES[i + 2*n]))
    if condor:
        for n, _ in CONDOR_STEPS:
            for i in range(N - 3*n):
                out.append(condor(EXPIRIES[i], EXPIRIES[i + 3*n]))
    return out


# ── Bloomberg fetch layer ─────────────────────────────────────────────────────
def _bbg_session():
    import blpapi
    o = blpapi.SessionOptions()
    o.setServerHost('localhost')
    o.setServerPort(8194)
    s = blpapi.Session(o)
    if not s.start() or not s.openService('//blp/refdata'):
        raise RuntimeError('Cannot connect to Bloomberg — is the terminal running?')
    return s

def _bbg_bdp(tickers, fields):
    """Bulk BDP.  Returns {ticker: {field: value}}."""
    import blpapi
    s = _bbg_session()
    svc = s.getService('//blp/refdata')
    req = svc.createRequest('ReferenceDataRequest')
    for t in tickers: req.append('securities', t)
    for f in fields:   req.append('fields', f)
    s.sendRequest(req)

    out = {t: {} for t in tickers}
    while True:
        ev = s.nextEvent(5000)
        for msg in ev:
            if msg.hasElement('securityData'):
                sd = msg.getElement('securityData')
                for i in range(sd.numValues()):
                    sec = sd.getValue(i)
                    t   = sec.getElementAsString('security')
                    fd  = sec.getElement('fieldData')
                    for f in fields:
                        if fd.hasElement(f):
                            try:    out[t][f] = fd.getElementAsFloat(f)
                            except: out[t][f] = None
        if ev.eventType() == blpapi.Event.RESPONSE:
            break
    s.stop()
    return out

def _bbg_bdh(tickers, start, end):
    """BDH PX_LAST oldest→newest.  Returns {ticker: [price, ...]}."""
    import blpapi
    s = _bbg_session()
    svc = s.getService('//blp/refdata')
    req = svc.createRequest('HistoricalDataRequest')
    for t in tickers: req.append('securities', t)
    req.append('fields', 'PX_LAST')
    req.set('startDate', start.strftime('%Y%m%d'))
    req.set('endDate',   end.strftime('%Y%m%d'))
    req.set('periodicitySelection', 'DAILY')
    req.set('nonTradingDayFillOption', 'NIL_VALUE')
    s.sendRequest(req)

    out = {t: [] for t in tickers}
    while True:
        ev = s.nextEvent(5000)
        for msg in ev:
            if msg.hasElement('securityData'):
                sec = msg.getElement('securityData')
                t   = sec.getElementAsString('security')
                fd  = sec.getElement('fieldData')
                for i in range(fd.numValues()):
                    row = fd.getValue(i)
                    try:    out[t].append(row.getElementAsFloat('PX_LAST'))
                    except: pass
        if ev.eventType() == blpapi.Event.RESPONSE:
            break
    s.stop()
    return out

def _build_rows(tickers, detail, hist):
    rows = []
    for t in tickers:
        d      = detail[t]
        prices = [p for p in (hist.get(t, [])[-20:]) if p is not None]
        last   = d.get('PX_LAST')
        vol20d = d.get('VOLUME_AVG_20D')
        roll   = d.get('ROLL_DOWN_VALUE')
        vol    = d.get('VOLUME')
        chg = zscore = v_avg = None

        if prices and last is not None and len(prices) >= 2:
            chg = last - prices[-1]
            av = np.mean(prices)
            sd = np.std(prices, ddof=1)
            if sd > 0:
                zscore = (last - av) / sd
        if vol and vol20d and vol20d > 0:
            v_avg = vol / vol20d

        alert = ''
        if zscore is not None and roll is not None and roll != 0:
            if zscore < 0 and roll > 0:   alert = 'BUY'
            elif zscore > 0 and roll < 0: alert = 'SELL'

        rows.append(dict(
            ticker=t, instrument=t.replace(' Comdty', ''),
            last=last, chg=chg, volume=vol,
            bid=d.get('PX_BID'), offer=d.get('PX_ASK'),
            zscore=zscore, history=prices,
            roll=roll, vol20d=vol20d, v_avg=v_avg, alert=alert,
        ))
    return rows

def fetch_top10_bloomberg(product, n=10):
    tickers = get_tickers(product)
    print(f'    Ranking {len(tickers)} tickers by volume...', end=' ', flush=True)
    vol = _bbg_bdp(tickers, ['VOLUME'])
    ranked = sorted(tickers, key=lambda t: vol[t].get('VOLUME') or 0, reverse=True)[:n]

    fields = ['PX_LAST', 'PX_BID', 'PX_ASK', 'VOLUME', 'VOLUME_AVG_20D', 'ROLL_DOWN_VALUE']
    detail = _bbg_bdp(ranked, fields)
    hist   = _bbg_bdh(ranked,
                      date.today() - timedelta(days=35),
                      date.today() - timedelta(days=1))
    return _build_rows(ranked, detail, hist)


# ── Excel fetch via xlwings (reads Bloomberg-populated values from open workbook) ──
def fetch_top10_excel(product, excel_path=None):
    """
    Read Top 10 data from the Most.Traded.Universe.xlsx workbook.
    Bloomberg must have already populated the formulas (file open in Excel).
    Columns read: A ticker, B instrument, C last, D chg, E vol, F bid, G offer,
                  H zscore, J roll, K vol20d, L v_avg, M alert, P:AI history (20d).
    """
    try:
        import xlwings as xw
    except ImportError:
        raise RuntimeError('xlwings not installed — run: pip install xlwings')

    tab = f'{product} Top10'

    # Attach to already-open workbook, or open it
    wb = None
    for book in xw.books:
        if 'Most.Traded.Universe' in book.name:
            wb = book
            break

    if wb is None:
        if excel_path is None:
            excel_path = Path(__file__).parent / 'Workbook' / 'Most.Traded.Universe.xlsx'
        print(f'\n    Opening {Path(excel_path).name} — wait for Bloomberg to load, '
              f'then press Enter...', flush=True)
        wb = xw.Book(str(excel_path))
        input()   # give the user time to let Bloomberg populate

    ws = wb.sheets[tab]
    rows = []

    for r in range(3, 13):   # Excel rows 3-12 (data rows 1-10)
        instr = ws[f'B{r}'].value
        if not instr:
            continue

        def _v(cell):
            v = ws[cell].value
            return float(v) if isinstance(v, (int, float)) else None

        # 20-day history from cols P(16) to AI(35), 1-indexed
        raw_hist = ws.range(f'P{r}:AI{r}').value or []
        history  = [float(v) for v in raw_hist
                    if v is not None and isinstance(v, (int, float))]

        alert = str(ws[f'M{r}'].value or '')

        rows.append(dict(
            ticker=str(ws[f'A{r}'].value or f'{instr} Comdty'),
            instrument=str(instr),
            last=_v(f'C{r}'),  chg=_v(f'D{r}'),    volume=_v(f'E{r}'),
            bid=_v(f'F{r}'),   offer=_v(f'G{r}'),   zscore=_v(f'H{r}'),
            history=history,
            roll=_v(f'J{r}'),  vol20d=_v(f'K{r}'),  v_avg=_v(f'L{r}'),
            alert=alert,
        ))

    if not rows:
        raise RuntimeError(
            f'No data found in "{tab}" — make sure Bloomberg has populated the sheet '
            f'and the product tab has been sorted by Volume.')
    return rows


# ── Mock / demo data (no Bloomberg required) ──────────────────────────────────
_MOCK_INSTRUMENTS = {
    'SFR': ['SFRM6U6','SFRZ6H7','SFRM6Z6','SFRU6H7','SFRM7U7',
            'BSFRM6Z6','BSFRU6H7','BSFRM7Z7','CSFRM6H7','SFR2YM6'],
    'SFI': ['SFIM6U6','SFIZ6H7','SFIM6Z6','SFIU6H7','SFIM7U7',
            'BSFIM6Z6','BSFIU6H7','BSFIM7Z7','BSFIZ6H8','SFI1YM6'],
    'ER':  ['ERM6ERU6','ERZ6ERH7','ERM6ERZ6','ERU6ERH7','ERM7ERU7',
            'BERM6Z6','BERU6H7','BERM7Z7','BERZ6H8','TKYER M6'],
    'IR':  ['IRM6IRU6','IRZ6IRH7','IRM6IRZ6','IRU6IRH7','IRM7IRU7',
            'BIRM6Z6','BIRU6H7','BIRM7Z7','BIRZ6H8','BIRM6Z7'],
    'COR': ['CORM6U6','CORZ6H7','CORM6Z6','CORU6H7','CORM7U7',
            'BCORM6Z6','BCORU6H7','BCORM7Z7','BCORZ6H8','BCORM6Z7'],
}

def fetch_top10_mock(product, n=10):
    rng = np.random.default_rng(hash(product) & 0xFFFFFFFF)
    instruments = _MOCK_INSTRUMENTS.get(product, [f'{product}{i}' for i in range(n)])[:n]
    rows = []
    for instr in instruments:
        is_spread = not (instr.startswith('B') or instr.startswith('C') or 'TKY' in instr)
        base  = float(rng.uniform(-0.12, 0.04) if is_spread else rng.uniform(-0.06, 0.06))
        hist  = [base + float(rng.normal(0, 0.013)) for _ in range(20)]
        last  = hist[-1]
        vol   = int(rng.integers(4000, 28000))
        vol20d = max(1, int(vol * float(rng.uniform(0.6, 1.6))))
        roll  = None if rng.random() < 0.25 else float(rng.uniform(-0.015, 0.045))
        av = np.mean(hist); sd = np.std(hist, ddof=1)
        zscore = (last - av) / sd if sd else None
        alert = ''
        if zscore is not None and roll is not None:
            if zscore < 0 and roll > 0:   alert = 'BUY'
            elif zscore > 0 and roll < 0: alert = 'SELL'
        rows.append(dict(
            ticker=f'{instr} Comdty', instrument=instr,
            last=last, chg=last - hist[-2], volume=vol,
            bid=last - 0.005, offer=last + 0.005,
            zscore=zscore, history=hist,
            roll=roll, vol20d=vol20d, v_avg=vol / vol20d, alert=alert,
        ))
    rows.sort(key=lambda r: r['volume'], reverse=True)
    return rows[:n]


# ── Table image renderer ──────────────────────────────────────────────────────
def render_table_png(product, rows, dpi=120):
    cfg    = PRODUCTS[product]
    dark   = cfg['dark']

    TITLE_H = 0.38
    HDR_H   = 0.34
    ROW_H   = 0.31
    TOTAL_W = sum(COL_W)
    TOTAL_H = TITLE_H + HDR_H + len(rows) * ROW_H

    fig = plt.figure(figsize=(TOTAL_W, TOTAL_H), dpi=dpi)
    ax  = fig.add_axes([0, 0, 1, 1])          # axes fills entire figure
    ax.set_xlim(0, TOTAL_W)
    ax.set_ylim(0, TOTAL_H)
    ax.axis('off')

    def rect(x, y, w, h, fc, ec='#CCCCCC', lw=0.3):
        ax.add_patch(mpatches.Rectangle(
            (x, y), w, h, facecolor=fc, edgecolor=ec, linewidth=lw, zorder=1))

    def cell_text(x, y, w, h, s, fs=8.0, fc='black', bold=False, ha='center'):
        if not s:
            return
        xoff = w * (0.06 if ha == 'left' else 0.94 if ha == 'right' else 0.50)
        ax.text(x + xoff, y + h * 0.5, s,
                fontsize=fs, color=fc,
                fontweight='bold' if bold else 'normal',
                ha=ha, va='center', clip_on=True, zorder=2)

    # ── Title ─────────────────────────────────────────────────────────────────
    rect(0, TOTAL_H - TITLE_H, TOTAL_W, TITLE_H, dark, dark)
    ax.text(0.12, TOTAL_H - TITLE_H * 0.5,
            f"{product} — Most Traded (Top 10)  ·  {cfg['label']}",
            fontsize=10.5, fontweight='bold', color='white', va='center', zorder=2)

    # COEX logo — right-aligned in title bar (accepts any coex*.png, case-insensitive)
    _candidates = sorted(Path(__file__).parent.glob('[Cc]oex*.png'))
    logo_path = _candidates[0] if _candidates else None
    if logo_path and logo_path.exists():
        logo_img = plt.imread(str(logo_path))
        logo_h_in = TITLE_H * 0.70          # logo takes 70% of title bar height
        logo_ar   = logo_img.shape[1] / logo_img.shape[0]   # width / height
        logo_w_in = logo_h_in * logo_ar
        margin    = TITLE_H * 0.15
        fx0 = (TOTAL_W - logo_w_in - margin) / TOTAL_W
        fy0 = (TOTAL_H - TITLE_H + margin)   / TOTAL_H
        fw  = logo_w_in / TOTAL_W
        fh  = logo_h_in / TOTAL_H
        ax_logo = fig.add_axes([fx0, fy0, fw, fh])
        ax_logo.imshow(logo_img, aspect='auto')
        ax_logo.axis('off')
    else:
        # Fallback: text brand mark when no logo file is present
        ax.text(TOTAL_W - 0.12, TOTAL_H - TITLE_H * 0.5, 'COEX',
                fontsize=9, fontweight='bold', color='white',
                alpha=0.75, ha='right', va='center', zorder=2)

    # ── Column headers ────────────────────────────────────────────────────────
    y_hdr = TOTAL_H - TITLE_H - HDR_H
    x = 0
    for label, w in zip(COLS, COL_W):
        rect(x, y_hdr, w, HDR_H, '#2F2F2F', '#2F2F2F')
        cell_text(x, y_hdr, w, HDR_H, label, fs=7.5, fc='white', bold=True)
        x += w

    # ── Data rows ─────────────────────────────────────────────────────────────
    BGS = ['#FFFFFF', '#F5F5F5']

    for ri, row in enumerate(rows):
        y  = TOTAL_H - TITLE_H - HDR_H - (ri + 1) * ROW_H
        bg = BGS[ri % 2]

        z_bg, z_fc = _zscore_style(row['zscore'])
        if z_bg is None: z_bg = bg
        a_bg, a_fc = _alert_style(row['alert'])
        if a_bg is None: a_bg = bg
        z_bold = abs(row['zscore'] or 0) >= 1

        cell_defs = [
            # (text,                        bg,   ha,       bold,                fc     )
            (row['instrument'],              bg,   'left',   False,               'black'),
            (_fmt(row['last'],   'px3'),     bg,   'right',  False,               'black'),
            (_fmt(row['chg'],    'px4'),     bg,   'right',  False,               'black'),
            (_fmt(row['volume'], 'vol'),     bg,   'right',  False,               'black'),
            (_fmt(row['bid'],    'px3'),     bg,   'right',  False,               'black'),
            (_fmt(row['offer'],  'px3'),     bg,   'right',  False,               'black'),
            (_fmt(row['zscore'], 'zsc'),     z_bg, 'right',  z_bold,              z_fc   ),
            (None,                           bg,   'center', False,               'black'),  # sparkline
            (_fmt(row['roll'],   'px3'),     bg,   'right',  False,               'black'),
            (_fmt(row['vol20d'], 'vol'),     bg,   'right',  False,               'black'),
            (_fmt(row['v_avg'],  'rat'),     bg,   'right',  False,               'black'),
            (row['alert'],                   a_bg, 'center', bool(row['alert']),  a_fc   ),
        ]

        x = 0
        for ci, (text, cell_bg, ha, bold, fc) in enumerate(cell_defs):
            w = COL_W[ci]
            rect(x, y, w, ROW_H, cell_bg)

            if ci == SPK_IDX:
                # Embedded mini sparkline chart
                hist = [p for p in row['history'] if p is not None]
                if hist:
                    pad_x = w * 0.05
                    pad_y = ROW_H * 0.12
                    # Convert data coords → figure fraction (axes fills whole figure)
                    fx0 = (x + pad_x) / TOTAL_W
                    fy0 = (y + pad_y) / TOTAL_H
                    fw  = (w - 2 * pad_x) / TOTAL_W
                    fh  = (ROW_H - 2 * pad_y) / TOTAL_H
                    ax_s = fig.add_axes([fx0, fy0, fw, fh])
                    xs = list(range(len(hist)))
                    ax_s.plot(xs, hist, color=dark, linewidth=1.1,
                              solid_capstyle='round')
                    if len(hist) > 1:
                        hi_i = int(np.argmax(hist))
                        lo_i = int(np.argmin(hist))
                        ax_s.scatter([hi_i], [hist[hi_i]],
                                     color='#00B050', s=14, zorder=5)
                        ax_s.scatter([lo_i], [hist[lo_i]],
                                     color='#FF0000', s=14, zorder=5)
                    ax_s.set_xticks([]); ax_s.set_yticks([])
                    ax_s.set_facecolor('none')
                    for sp in ax_s.spines.values():
                        sp.set_visible(False)
            else:
                cell_text(x, y, w, ROW_H, text or '',
                          fs=8.0, fc=fc, bold=bold, ha=ha)
            x += w

    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=dpi)
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ── Bloomberg IB text table ───────────────────────────────────────────────────
def format_ib_text(product, rows):
    cfg = PRODUCTS[product]
    sep = '─' * 84
    lines = [
        sep,
        f"  {product} — Most Traded Top 10  ·  {cfg['label']}  ·  {date.today():%d %b %Y}",
        sep,
        f"  {'Instrument':<16} {'Last':>8} {'Chg':>9} {'Volume':>8} "
        f"{'Z-score':>8} {'ROLL':>7} {'V/Avg':>6}  Signal",
        f"  {'─'*16} {'─'*8} {'─'*9} {'─'*8} {'─'*8} {'─'*7} {'─'*6}  {'─'*6}",
    ]
    for r in rows:
        sig = f'  *** {r["alert"]} ***' if r['alert'] else ''
        lines.append(
            f"  {r['instrument']:<16} "
            f"{_fmt(r['last'],   'px3'):>8} "
            f"{_fmt(r['chg'],    'px4'):>9} "
            f"{_fmt(r['volume'], 'vol'):>8} "
            f"{_fmt(r['zscore'], 'zsc'):>8} "
            f"{_fmt(r['roll'],   'px3'):>7} "
            f"{_fmt(r['v_avg'],  'rat'):>6}"
            f"{sig}"
        )
    lines.append(sep + '\n')
    return '\n'.join(lines)


# ── HTML email builder ────────────────────────────────────────────────────────
_HTML = """\
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body   {{ font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #222; margin: 20px; }}
  .hdr   {{ background:#1F497D; color:white; padding:10px 16px;
            font-size:13pt; font-weight:bold; margin-bottom:18px; border-radius:3px; }}
  .sect  {{ margin-bottom:20px; }}
  img    {{ max-width:100%; border:1px solid #DDD; border-radius:2px; }}
  .foot  {{ font-size:8.5pt; color:#999; margin-top:24px;
            border-top:1px solid #EEE; padding-top:8px; }}
  b.buy  {{ color:#00B050; }}
  b.sell {{ color:#CC0000; }}
</style>
</head>
<body>
<div class="hdr">London Morning — Rate Structures Snapshot &nbsp;·&nbsp; {date_str}</div>
<p>Good morning,</p>
<p>Below are the top 10 most-active structures by volume for each product,
   ranked by today's volume.
   Columns: last price, change on day, volume, bid/offer, 20-day z-score, sparkline,
   roll-down, 20-day average volume, and volume ratio.<br>
   Alert signals:&nbsp; <b class="buy">BUY</b> = z-score negative &amp; roll positive
   &nbsp;|&nbsp; <b class="sell">SELL</b> = z-score positive &amp; roll negative.</p>

{sections}

<div class="foot">
  Generated {datetime_str} UTC &nbsp;·&nbsp; Most.Traded.Universe pipeline &nbsp;·&nbsp; Source: Bloomberg
</div>
</body>
</html>
"""

_SECTION = """\
<div class="sect">
  <img src="data:image/png;base64,{b64}" alt="{product} Top 10">
</div>
"""

def build_email_html(images: dict) -> str:
    now = datetime.now(timezone.utc)
    sections = '\n'.join(
        _SECTION.format(b64=base64.b64encode(png).decode(), product=prod)
        for prod, png in images.items()
    )
    return _HTML.format(
        date_str=now.strftime('%d %b %Y'),
        datetime_str=now.strftime('%d %b %Y  %H:%M'),
        sections=sections,
    )


# ── Bloomberg IB automation ───────────────────────────────────────────────────
# Bloomberg Instant Messaging has no public API.  Options:
#
# 1. DEFAULT — manual paste:
#    Script saves morning_YYYYMMDD_ib.txt.  Open Bloomberg IB (MSG <GO>),
#    open the target conversation, Ctrl+V.
#
# 2. --ib-paste flag — pyautogui automation (Windows only):
#    Finds the Bloomberg IB window by title, brings it to foreground, pastes.
#    Requirements: pip install pyautogui pyperclip pywin32
#    Pre-conditions:
#      • Bloomberg terminal must be running
#      • IB window (MSG <GO>) must be open with the target conversation visible
#
# 3. blpapi EMSX — NOT applicable; EMSX handles order messages, not IB chat.
#    There is no blpapi endpoint for sending Instant Bloomberg messages.

def paste_to_bloomberg_ib(text: str) -> bool:
    if sys.platform != 'win32':
        print('  Bloomberg IB paste is Windows-only.')
        print('  Copy the _ib.txt file contents manually into Bloomberg IB (MSG <GO>).')
        return False
    try:
        import pyperclip
        import pyautogui
        import time
        import win32gui
    except ImportError:
        print('  Install required packages: pip install pyautogui pyperclip pywin32')
        return False

    pyperclip.copy(text)

    ib_hwnd = None
    def _cb(hwnd, _):
        nonlocal ib_hwnd
        t = win32gui.GetWindowText(hwnd)
        if any(k in t for k in ('Bloomberg Instant', 'IB -', 'MSG ')):
            ib_hwnd = hwnd
    win32gui.EnumWindows(_cb, None)

    if ib_hwnd is None:
        print('  Bloomberg IB window not found.')
        print('  Open IB (MSG <GO>), select the target conversation, then re-run with --ib-paste.')
        return False

    win32gui.SetForegroundWindow(ib_hwnd)
    time.sleep(0.4)
    pyautogui.hotkey('ctrl', 'v')
    print('  Pasted to Bloomberg IB.  Review and press Enter to send.')
    return True


# ── CLI ───────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description='London Morning Email — rate structure Top 10 images')
    parser.add_argument('products', nargs='*', metavar='PRODUCT',
                        help='One or more of: SFR SFI ER IR COR')
    parser.add_argument('--all',          action='store_true',
                        help='Run all five currencies')
    parser.add_argument('--no-bloomberg',   action='store_true',
                        help='Use demo data — no Bloomberg or Excel required')
    parser.add_argument('--from-bloomberg', action='store_true',
                        help='Pull data direct from Bloomberg API (requires blpapi)')
    parser.add_argument('--ib-paste',       action='store_true',
                        help='Paste text table into Bloomberg IB window (Windows only)')
    parser.add_argument('--out',          default='./output', metavar='DIR',
                        help='Output folder (default: ./output)')
    parser.add_argument('--dpi',          type=int, default=120,
                        help='Image resolution in DPI (default: 120)')
    parser.add_argument('--n',            type=int, default=10,
                        help='Top N instruments per currency (default: 10)')
    args = parser.parse_args()

    if args.all:
        selected = list(PRODUCTS.keys())
    elif args.products:
        selected = [p.upper() for p in args.products]
        bad = [p for p in selected if p not in PRODUCTS]
        if bad:
            print(f'Unknown products: {bad}  Valid: {list(PRODUCTS.keys())}')
            sys.exit(1)
    else:
        parser.print_help()
        sys.exit(0)

    if args.no_bloomberg:
        fetch_fn = fetch_top10_mock
    elif args.from_bloomberg:
        fetch_fn = fetch_top10_bloomberg
    else:
        fetch_fn = fetch_top10_excel   # default: read from open Excel workbook
    out_dir  = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp    = date.today().strftime('%Y%m%d')

    images      = {}
    ib_sections = []

    print(f'\nLondon Morning Pipeline  —  {date.today():%d %b %Y}')
    print(f'Products : {" ".join(selected)}')
    mode = 'demo (mock data)' if args.no_bloomberg else \
           'Bloomberg API (blpapi)' if args.from_bloomberg else \
           'Excel workbook (xlwings)'
    print(f'Mode     : {mode}')
    print(f'Output   : {out_dir.resolve()}\n')

    for prod in selected:
        print(f'[{prod}] fetching...', end=' ', flush=True)
        try:
            rows = fetch_fn(prod, n=args.n)
        except Exception as e:
            print(f'ERROR — {e}')
            continue

        print(f'{len(rows)} rows  →  rendering...', end=' ', flush=True)
        png = render_table_png(prod, rows, dpi=args.dpi)

        png_path = out_dir / f'{prod}_top10.png'
        png_path.write_bytes(png)
        images[prod] = png
        ib_sections.append(format_ib_text(prod, rows))
        print(f'saved {png_path.name}')

    if not images:
        print('No data produced — exiting.')
        sys.exit(1)

    # HTML email draft
    html_path = out_dir / f'morning_{stamp}.html'
    html_path.write_text(build_email_html(images), encoding='utf-8')
    print(f'\nEmail draft  :  {html_path}')

    # Bloomberg IB text
    ib_text  = '\n'.join(ib_sections)
    ib_path  = out_dir / f'morning_{stamp}_ib.txt'
    ib_path.write_text(ib_text, encoding='utf-8')
    print(f'IB text      :  {ib_path}')

    if args.ib_paste:
        print('\nPasting to Bloomberg IB...')
        paste_to_bloomberg_ib(ib_text)
    else:
        print('\nTo send via Bloomberg IB:')
        print(f'  1. Open Bloomberg IB  (MSG <GO>)  and select your conversation')
        print(f'  2. Copy {ib_path.name} and Ctrl+V  — or re-run with --ib-paste')


if __name__ == '__main__':
    main()
