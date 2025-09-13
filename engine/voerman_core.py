import os
import pandas as pd

def _validity():
    try:
        return int(os.environ.get("VALIDITY_DAYS","14"))
    except Exception:
        return 14

def _load_rates():
    path = os.environ.get("PRICING_EXCEL_PATH","tarieven.xlsx")
    if not os.path.exists(path):
        return None
    try:
        df = pd.read_excel(path)
    except Exception:
        # try common sheet name
        df = pd.read_excel(path, sheet_name=0)
    # normalize columns
    cols = {c.lower().strip():c for c in df.columns}
    # expected: mode, item, unit, rate
    def pick(*names):
        for n in names:
            if n in cols: return cols[n]
        return None
    return df.rename(columns={
        pick('mode','modus'): 'mode',
        pick('item','description','descr'): 'item',
        pick('unit','per'): 'unit',
        pick('rate','price','tarief','amount'): 'rate'
    })

def build_lines(req):
    df = _load_rates()
    lines = []
    if df is None or 'mode' not in df or 'rate' not in df:
        return lines
    # choose first mode
    mode = (req.modes or ['LCL'])[0] if isinstance(req.modes, list) else getattr(req,'mode','LCL')
    dff = df
    try:
        dff = df[df['mode'].str.upper()==str(mode).upper()]
    except Exception:
        pass
    vols = getattr(req,'volumes',[]) or []
    total_m3 = 0.0
    for v in vols:
        try:
            if (getattr(v,'unit','m3') or 'm3').lower() in ('m3','cbm'):
                total_m3 += float(getattr(v,'value',0) or 0)
        except Exception:
            pass
    # freight per m3
    try:
        row = dff[dff['unit'].str.lower().isin(['m3','cbm'])].iloc[0]
        rate = float(row['rate'])
        amt = round(rate * (total_m3 or 1.0), 2)
        lines.append({'descr': f"Freight ({mode})", 'qty': f"{total_m3:.2f} cbm", 'rate': f"€ {rate:.2f}/cbm", 'amount': amt})
    except Exception:
        pass
    # flat items (origin/dest/handling)
    try:
        flats = dff[dff['unit'].str.lower().isin(['flat','fixed','eenmalig'])]
        for _, r in flats.iterrows():
            lines.append({'descr': str(r.get('item') or 'Flat'), 'qty': '1', 'rate': f"€ {float(r['rate']):.2f}", 'amount': float(r['rate'])})
    except Exception:
        pass
    return lines

def totals_from_lines(lines):
    buy = sum(float(x.get('amount',0) or 0) for x in (lines or []))
    sell = round(buy, 2)
    return {'buy_total': buy, 'sell_total': sell}

def render_pdf(req, lines, path):
    # simple reportlab render; main app already has a fallback
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
        c = canvas.Canvas(path, pagesize=A4)
        c.drawString(72, 800, "Quote")
        y = 760
        for li in lines:
            c.drawString(72, y, f"{li.get('descr','')}  {li.get('qty','')}  {li.get('rate','')}  = € {float(li.get('amount',0.0)):.2f}"); y -= 18
        total = sum(float(li.get('amount',0) or 0) for li in lines)
        c.drawString(72, y-10, f"Total: € {total:.2f}")
        c.showPage(); c.save()
        return path
    except Exception:
        open(path,'wb').close(); return path
