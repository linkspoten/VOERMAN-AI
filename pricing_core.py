# pricing_core.py
from __future__ import annotations
import os, uuid
from typing import Any, Dict, List
from dotenv import load_dotenv
load_dotenv(override=False)

def _env(k: str, d: str="") -> str: return os.getenv(k, d)
def _ensure_out() -> str:
    out = _env("OUT_DIR","out") or "out"; os.makedirs(out, exist_ok=True); return out

def _label(req: Any) -> str:
    def _fmt(p):
        if not p: return "-"
        city = getattr(p,"city",None) or getattr(p,"POL",None) or getattr(p,"IATA",None) or ""
        cty  = getattr(p,"country",None) or ""
        s = ", ".join([x for x in [city, cty] if x])
        return s or getattr(p,"POL",None) or getattr(p,"IATA",None) or "-"
    mode = (getattr(req,"modes",["LCL"]) or ["LCL"])[0] if isinstance(getattr(req,"modes",None), list) else getattr(req,"mode","LCL")
    return f"{mode} – {_fmt(getattr(req,'origin',None))} → {_fmt(getattr(req,'destination',None))}"

def _placeholder_lines_and_totals(req: Any):
    vols = getattr(req,"volumes",[]) or []
    lines: List[Dict[str,Any]] = []; total = 0.0
    for v in vols:
        unit = getattr(v,"unit",None) or (isinstance(v,dict) and v.get("unit")) or "m3"
        val  = getattr(v,"value",None) or (isinstance(v,dict) and v.get("value")) or 0.0
        try: val = float(val)
        except Exception: val = 0.0
        if str(unit).lower() in ("m3","cbm"):
            rate=95.0; amt=round(val*rate,2)
            lines.append({"descr":"Freight (LCL) indicative","qty":f"{val:.2f} cbm","rate":f"€ {rate:.2f}/cbm","amount":amt}); total+=amt
        elif str(unit).lower() in ("kg","kilogram"):
            rate=1.25; amt=round(val*rate,2)
            lines.append({"descr":"Air freight indicative","qty":f"{val:.0f} kg","rate":f"€ {rate:.2f}/kg","amount":amt}); total+=amt
    if not lines:
        lines=[{"descr":"Base service (indicative)","qty":"1","rate":"€ 250.00","amount":250.00}]; total=250.00
    buy = float(round(max(0.0, total*0.85), 2)); sell = float(round(total,2))
    return lines, buy, sell

def _services_from_req(req: Any) -> list[str]:
    s = getattr(req, "services", None) or []
    if isinstance(s, (list, tuple)): return [str(x).lower() for x in s]
    return ["origin","freight","destination"]

def generate_quote(req: Any) -> List[Dict[str, Any]]:
    """
    ALTIJD Studio voor de PDF. Voor prijs/regels gebruiken we een simpele
    berekening zodat mail + preview altijd werken.
    Als Studio geen PDF kan leveren, wordt een fallback-PDF gemaakt.
    """
    out_dir = _ensure_out()
    label = _label(req)

    # 1) simpele regels + totalen (voor mail/preview én voor fallback)
    lines, buy, sell = _placeholder_lines_and_totals(req)

    # 2) Studio aanroepen voor de PDF (altijd)
    try:
        total_cbm = 0.0
        for v in getattr(req,"volumes",[]) or []:
            try:
                unit = (getattr(v,"unit",None) or (isinstance(v,dict) and v.get("unit")) or "m3").lower()
                val = float(getattr(v,"value",None) or (isinstance(v,dict) and v.get("value")) or 0.0)
                if unit in ("m3","cbm"): total_cbm += val
            except Exception: pass

        from studio_adapter import generate_pdf_with_studio
        brand = _env("BRAND","Voerman")
        services = _services_from_req(req)
        origin_label = getattr(getattr(req,"origin",None), "city", None) or getattr(getattr(req,"origin",None), "POL", None) or "-"
        dest_label   = getattr(getattr(req,"destination",None), "city", None) or getattr(getattr(req,"destination",None), "POD", None) or "-"
        mode = (getattr(req,"modes",['LCL']) or ['LCL'])[0] if isinstance(getattr(req,"modes",None), list) else getattr(req,"mode","LCL")

        pdf_path = generate_pdf_with_studio(
            brand=brand, services=services, mode=mode, total_cbm=total_cbm,
            origin_label=origin_label, dest_label=dest_label, req_label=label,
            priced_lines=lines
        )
    except Exception:
        pdf_path = os.path.join(out_dir, f"quote_{uuid.uuid4().hex[:8]}.pdf")
        open(pdf_path,"wb").close()

    return [{
        "label": label,
        "buy_total": buy,
        "sell_total": sell,
        "validity": f"{int(_env('VALIDITY_DAYS','14'))} dagen",
        "pdf_path": pdf_path,
        "assumptions": [],
        "mode": (getattr(req,'modes', ['LCL']) or ['LCL'])[0] if isinstance(getattr(req,'modes',None), list) else getattr(req,'mode','LCL')
    }]
