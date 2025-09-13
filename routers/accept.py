from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse
import storage, os, csv
from email_service import verify_token
router = APIRouter()
@router.get("/accept", response_class=HTMLResponse)
async def accept(request: Request, token: str):
    qid = verify_token(token)
    if not qid: return HTMLResponse("<h1>Invalid or expired link</h1>", status_code=400)
    storage.set_quote_status(qid, 'accepted'); storage.log_event('quote.accepted', {'quote_id': qid})
    out_dir = os.environ.get('OUT_DIR','out'); os.makedirs(out_dir, exist_ok=True)
    csv_path = os.path.join(out_dir, f'handoff_{qid}.csv')
    with open(csv_path,'w',newline='',encoding='utf-8') as f:
        w = csv.writer(f); w.writerow(['quote_id','status']); w.writerow([qid,'accepted'])
    return HTMLResponse(f"<h1>Offerte geaccepteerd</h1><p>Quote ID: {qid}</p>")
