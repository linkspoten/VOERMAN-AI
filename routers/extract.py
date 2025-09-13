from fastapi import APIRouter
from pydantic import BaseModel
import storage
from extractor import extract_from_unified
router = APIRouter()
class ExtractBody(BaseModel):
    message_id: str
@router.post("")
def extract(b: ExtractBody):
    msg = storage.get_message(b.message_id)
    if not msg: return {"error":"message not found"}
    res = extract_from_unified(msg)
    return {"request": res.request.dict(), "clarifying_questions": res.clarifying_questions}
