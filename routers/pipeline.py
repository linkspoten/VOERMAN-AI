from fastapi import APIRouter
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import os, time
import storage, pricing_core
from extractor import extract_from_unified
from email_service import render_preview
from models_contracts import QuoteOption

router = APIRouter()

class GenerateBody(BaseModel):
    message_id: str
    language: Optional[str] = None
    customer_name: Optional[str] = None

@router.post("/pipeline/generate")
def generate(b: GenerateBody):
    msg = storage.get_message(b.message_id)
    if not msg:
        return {"error":"message not found"}
    res = extract_from_unified(msg)
    qr = res.request
    if b.language:
        qr.language = b.language
    options: List[Dict[str, Any]] = []
    for m in (qr.modes or ["LCL"]):
        qr_single = qr.copy(update={"modes":[m]})
        for o in pricing_core.generate_quote(qr_single):
            o["mode"] = m
            options.append(o)
    html = render_preview(qr.language, options, b.customer_name or msg['sender'].get('email'), res.clarifying_questions, 'Met vriendelijke groet,\nVoerman Team', 'q_'+b.message_id)
    out_dir = os.environ.get('OUT_DIR','out'); os.makedirs(out_dir, exist_ok=True)
    filename = f"email_preview_{b.message_id}.html"
    html_path = os.path.join(out_dir, filename)
    web_path = f"/out/{filename}"
    with open(html_path,'w',encoding='utf-8') as f: f.write(html)
    atts = [o.get('pdf_path') for o in options if o.get('pdf_path')]
    return {"options": options, "html_path": html_path, "web_path": web_path, "html": html, "attachments": atts, "clarifying_questions": res.clarifying_questions}

class SendBody(BaseModel):
    to: str
    language: str = 'nl'
    options: List[QuoteOption]
    customer_name: Optional[str] = None
    quote_id: Optional[str] = None
    subject: Optional[str] = None

@router.post("/pipeline/send")
def send(b: SendBody):
    html = render_preview(b.language, [o.dict() for o in b.options], b.customer_name, [], 'Met vriendelijke groet,\nVoerman Team', b.quote_id or 'q_demo')
    atts = [o.pdf_path for o in b.options if o.pdf_path]
    if not atts:
        # as a fallback, attach any PDFs under out/quote_*.pdf
        out_dir = os.environ.get('OUT_DIR','out')
        try:
            for name in os.listdir(out_dir):
                if name.lower().startswith('quote_') and name.lower().endswith('.pdf'):
                    atts.append(os.path.join(out_dir, name))
        except Exception:
            pass
    subject = b.subject or (f"Offerte – {b.options[0].label}" if b.options else "Offerte")
    from email_service import send_via_smtp
    res = send_via_smtp(b.to, subject, html, atts)
    return {"ok": res.get("ok", False), "info": res.get("info",""), "attachments": atts}


class SendRawBody(BaseModel):
    to: str
    subject: str
    html: str
    attachments: list[str] = []

@router.post("/pipeline/send_raw")
def send_raw(b: SendRawBody):
    from email_service import send_raw_html
    res = send_raw_html(b.to, b.subject, b.html, b.attachments or [])
    return {"ok": res.get("ok", False), "info": res.get("info",""), "attachments": b.attachments}
