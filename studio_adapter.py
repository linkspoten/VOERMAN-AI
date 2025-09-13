# studio_adapter.py
from __future__ import annotations
import importlib.util, os, uuid, traceback
from typing import Any, Dict, List

OUT_DIR = os.environ.get("OUT_DIR", "out")
os.makedirs(OUT_DIR, exist_ok=True)

def _import_studio():
    path = os.environ.get("STUDIO_PATH", "Voerman_Quote_Studio_MQ26_P0PATCH.py")
    if not os.path.exists(path):
        for alt in [os.path.join(os.getcwd(), path), os.path.join(os.path.dirname(__file__), path)]:
            if os.path.exists(alt):
                path = alt; break
    if not os.path.exists(path):
        return None
    spec = importlib.util.spec_from_file_location("voerman_studio", path)
    mod = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(mod)  # type: ignore
    return mod

def _fallback_pdf(req_label: str, lines: List[Dict[str, Any]], brand: str) -> str:
    pdf = os.path.join(OUT_DIR, f"quote_{uuid.uuid4().hex[:8]}.pdf")
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
        c = canvas.Canvas(pdf, pagesize=A4)
        y = 800
        c.setFont("Helvetica-Bold", 14); c.drawString(72, y, f"{brand} – Offerte"); y -= 26
        c.setFont("Helvetica", 11); c.drawString(72, y, req_label); y -= 24
        total = 0.0
        for li in (lines or []):
            amt = float(li.get("amount", 0) or 0); total += amt
            c.drawString(72, y, f"- {li.get('descr','')}   {li.get('qty','')}   {li.get('rate','')}   = € {amt:,.2f}".replace(",", "X").replace(".", ",").replace("X","."))
            y -= 16
        y -= 6; c.setFont("Helvetica-Bold", 12)
        c.drawString(72, y, f"Totaal: € {total:,.2f}".replace(",", "X").replace(".", ",").replace("X","."))
        c.showPage(); c.save()
    except Exception:
        open(pdf, "wb").close()
    return pdf

def generate_pdf_with_studio(
    *, brand: str, services: List[str], mode: str, total_cbm: float,
    origin_label: str, dest_label: str, req_label: str,
    priced_lines: List[Dict[str, Any]]
) -> str:
    """
    1) Probeer Studio.api_generate_pdf(...)
    2) Zo niet: fallback-PDF.
    """
    try:
        studio = _import_studio()
        if studio and hasattr(studio, "api_generate_pdf"):
            out_path = os.path.join(OUT_DIR, f"quote_{uuid.uuid4().hex[:8]}.pdf")
            pdf = studio.api_generate_pdf(  # type: ignore[attr-defined]
                brand=brand, services=services, mode=mode, total_cbm=total_cbm,
                origin=origin_label, destination=dest_label, out_path=out_path,
                charges_rows=priced_lines, client_name=None, show_rates=True, show_vat=False
            )
            if isinstance(pdf, str) and os.path.exists(pdf) and os.path.getsize(pdf) > 0:
                return pdf
    except Exception:
        traceback.print_exc()

    return _fallback_pdf(req_label, priced_lines, brand or "Voerman")
