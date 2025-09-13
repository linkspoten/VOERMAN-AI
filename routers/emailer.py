from fastapi import APIRouter
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
from models_contracts import QuoteOption
import os
from email_service import render_preview
router = APIRouter()
class EmailPreviewBody(BaseModel):
    language: str = 'nl'
    options: List[QuoteOption]
    customer_name: Optional[str] = None
    clarifying_questions: List[Dict[str, Any]] = []
    signature_block: Optional[str] = 'Met vriendelijke groet,\nVoerman Team'
    quote_id: Optional[str] = 'q_demo'
@router.post("/preview")
def preview(b: EmailPreviewBody):
    html = render_preview(b.language, [o.dict() for o in b.options], b.customer_name, b.clarifying_questions, b.signature_block, b.quote_id)
    out_dir = os.environ.get('OUT_DIR','out'); os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, 'email_preview.html'); open(path,'w',encoding='utf-8').write(html)
    return {"html_path": path, "length": len(html)}

class EmailSendBody(BaseModel):
    to: str
    language: str = 'nl'
    options: List[QuoteOption]
    customer_name: Optional[str] = None
    clarifying_questions: List[Dict[str, Any]] = []
    signature_block: Optional[str] = 'Met vriendelijke groet,\nVoerman Team'
    quote_id: Optional[str] = 'q_demo'
    subject: Optional[str] = None

@router.post("/send")
def send_email(b: EmailSendBody):
    # Render preview first
    html = render_preview(b.language, [o.dict() for o in b.options], b.customer_name, b.clarifying_questions, b.signature_block, b.quote_id)
    # Gather attachments from options
    atts = [o.pdf_path for o in b.options if o.pdf_path]
    # Subject fallback
    subject = b.subject or (f"Offerte – {b.options[0].label}" if b.options else "Offerte")
    # Send
    from email_service import send_via_smtp
    res = send_via_smtp(b.to, subject, html, atts)
    return {"ok": res.get("ok", False), "info": res.get("info",""), "attachments": atts}
