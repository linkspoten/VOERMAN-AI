from fastapi import APIRouter
from pydantic import BaseModel
import storage, time, uuid
router = APIRouter()
class IngestBody(BaseModel):
    source: str = "web"
    sender_email: str = "customer@example.com"
    subject: str = "Offerte aanvraag"
    body: str
    language: str = "nl"
@router.post("/test")
def ingest_test(b: IngestBody):
    mid = "umsg_" + uuid.uuid4().hex[:10]
    msg = {'id': mid,'source': b.source,'sender': {'email': b.sender_email},'subject': b.subject,'body': b.body,'attachments': [],'language': b.language,'timestamp': time.strftime('%Y-%m-%dT%H:%M:%SZ'),'thread_id': None,'message_id': None}
    storage.insert_message(msg); return msg
