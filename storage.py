import sqlite3, json, os, time, uuid

# Always keep DB next to this file by default
_default_path = os.environ.get("DB_PATH", "voerman.db")
if not os.path.isabs(_default_path):
    BASE_DIR = os.path.dirname(__file__)
    DB_PATH = os.path.join(BASE_DIR, _default_path)
else:
    DB_PATH = _default_path

def _conn():
    # Thread-safe enough for FastAPI threadpool
    return sqlite3.connect(DB_PATH)

def init_db():
    os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)
    with _conn() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS messages(
            id TEXT PRIMARY KEY,
            source TEXT,
            sender_email TEXT,
            subject TEXT,
            body_text TEXT,
            body_html TEXT,
            language TEXT,
            ts TEXT,
            thread_id TEXT,
            message_id TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS attachments(
            id TEXT PRIMARY KEY,
            message_id TEXT,
            uri TEXT,
            filename TEXT,
            mimetype TEXT,
            size INTEGER
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS quotes(
            id TEXT PRIMARY KEY,
            source_message_id TEXT,
            status TEXT,
            currency TEXT,
            created_at TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS quote_options(
            id TEXT PRIMARY KEY,
            quote_id TEXT,
            mode TEXT,
            service TEXT,
            label TEXT,
            buy_total REAL,
            sell_total REAL,
            validity TEXT,
            pdf_path TEXT,
            review_required INTEGER DEFAULT 0
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS events(
            id TEXT PRIMARY KEY,
            type TEXT,
            payload_json TEXT,
            created_at TEXT
        )""")
        c.commit()

def insert_message(m):
    sender_email = None
    if isinstance(m.get('sender'), dict):
        sender_email = m['sender'].get('email')
    sender_email = sender_email or m.get('sender_email')
    with _conn() as c:
        c.execute("INSERT OR REPLACE INTO messages VALUES(?,?,?,?,?,?,?,?,?,?)",
                  (m['id'], m.get('source'), sender_email, m.get('subject'),
                   m.get('body') or m.get('body_text') or '', m.get('body_html'),
                   m.get('language','nl'), m.get('timestamp'), m.get('thread_id'), m.get('message_id')))
        for a in m.get('attachments',[]) or []:
            aid = a.get('id') or str(uuid.uuid4())
            c.execute("INSERT OR REPLACE INTO attachments VALUES(?,?,?,?,?,?)",
                      (aid, m['id'], a.get('uri'), a.get('filename'), a.get('mimetype'), a.get('size') or 0))
        c.commit()

def get_message(mid):
    with _conn() as c:
        row = c.execute("SELECT id, source, sender_email, subject, body_text, body_html, language, ts, thread_id, message_id FROM messages WHERE id=?", (mid,)).fetchone()
        if not row:
            return None
        arows = c.execute("SELECT id, uri, filename, mimetype, size FROM attachments WHERE message_id=?", (mid,)).fetchall()
    return {
        'id': row[0],
        'source': row[1],
        'sender': {'email': row[2]},
        'subject': row[3],
        'body': row[4] or row[5] or '',
        'attachments': [{'id': r[0], 'uri': r[1], 'filename': r[2], 'mimetype': r[3], 'size': r[4]} for r in arows],
        'language': row[6],
        'timestamp': row[7],
        'thread_id': row[8],
        'message_id': row[9]
    }

def new_quote(source_message_id, currency='EUR'):
    qid = 'q_' + uuid.uuid4().hex[:10]
    with _conn() as c:
        c.execute("INSERT INTO quotes VALUES(?,?,?,?,?)", (qid, source_message_id, 'draft', currency, time.strftime('%Y-%m-%dT%H:%M:%SZ')))
        c.commit()
    return qid

def add_option(quote_id, opt: dict):
    oid = 'opt_' + uuid.uuid4().hex[:10]
    with _conn() as c:
        c.execute("INSERT INTO quote_options(id, quote_id, mode, service, label, buy_total, sell_total, validity, pdf_path, review_required) VALUES(?,?,?,?,?,?,?,?,?,?)",
                  (oid, quote_id, opt.get('mode'), opt.get('service'), opt.get('label'), float(opt.get('buy_total',0.0)), float(opt.get('sell_total',0.0)), opt.get('validity'), opt.get('pdf_path'), int(opt.get('review_required',0))))
        c.commit()
    return oid

def set_quote_status(qid, status):
    with _conn() as c:
        c.execute("UPDATE quotes SET status=? WHERE id=?", (status, qid))
        c.commit()

def log_event(type_, payload):
    eid = 'evt_' + uuid.uuid4().hex[:10]
    with _conn() as c:
        c.execute("INSERT INTO events VALUES(?,?,?,?)", (eid, type_, json.dumps(payload), time.strftime('%Y-%m-%dT%H:%M:%SZ')))
        c.commit()
    return eid
