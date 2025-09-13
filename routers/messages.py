from fastapi import APIRouter
import storage, time, uuid

router = APIRouter()

@router.get("/messages")
def list_messages(limit: int = 100, offset: int = 0):
    # list ordered by timestamp or rowid fallback
    with storage._conn() as c:
        try:
            rows = c.execute("SELECT id, source, sender_email, subject, ts FROM messages ORDER BY ts DESC LIMIT ? OFFSET ?", (limit, offset)).fetchall()
        except Exception:
            rows = c.execute("SELECT id, source, sender_email, subject, ts FROM messages ORDER BY rowid DESC LIMIT ? OFFSET ?", (limit, offset)).fetchall()
    return [{"id": r[0], "source": r[1], "sender_email": r[2], "subject": r[3], "timestamp": r[4]} for r in rows]

@router.get("/messages/{mid}")
def get_message(mid: str):
    msg = storage.get_message(mid)
    return msg or {"error":"not found"}

@router.post("/messages/seed")
def seed_message():
    mid = "umsg_" + uuid.uuid4().hex[:8]
    m = {
        "id": mid, "source": "web",
        "sender": {"email": "seed@example.com"},
        "subject": "Seed offerteaanvraag",
        "body": "Test: 10 m3 en 500 kg van Amsterdam naar Montreal.",
        "attachments": [], "language": "nl",
        "timestamp": time.strftime('%Y-%m-%dT%H:%M:%SZ'),
        "thread_id": None, "message_id": None
    }
    storage.insert_message(m)
    return {"ok": True, "id": mid}


@router.delete("/messages/{mid}")
def delete_message(mid: str):
    try:
        with storage._conn() as c:
            c.execute("DELETE FROM attachments WHERE message_id=?", (mid,))
            c.execute("DELETE FROM messages WHERE id=?", (mid,))
            c.commit()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}
