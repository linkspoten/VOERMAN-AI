import os, smtplib, mimetypes, hmac, hashlib, time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from jinja2 import Environment, FileSystemLoader, select_autoescape

def _env(k, d=""): return os.getenv(k, d)

def _jinja_env():
    tpl_dir = os.path.join(os.path.dirname(__file__), "templates")
    return Environment(loader=FileSystemLoader(tpl_dir), autoescape=select_autoescape(['html','xml']))

def render_preview(language: str, options, customer_name, questions, signoff, quote_id):
    env = _jinja_env()
    name = 'quote_nl.j2' if (language or 'nl').lower().startswith('nl') else 'quote_en.j2'
    try:
        tpl = env.get_template(name)
    except Exception:
        tpl = env.from_string("""
<html><body style='font-family:Arial,sans-serif'>
<p>Beste {{customer_name or 'relatie'}},</p>
{% if options %}
<p>Hierbij onze offerte:</p>
<ul>
{% for o in options %}<li>{{o.label}} – € {{ "%.2f"|format(o.sell_total) }}</li>{% endfor %}
</ul>
{% endif %}
{% if questions %}
<p>Open punten:</p>
<ul>
{% for q in questions %}<li>{{q.question}}</li>{% endfor %}
</ul>
{% endif %}
<p>Met vriendelijke groet,<br/>Voerman Team</p>
</body></html>
""")
    html = tpl.render(language=language, options=options, customer_name=customer_name, questions=questions, signoff=signoff, quote_id=quote_id)
    return html

def send_via_smtp(to_addr: str, subject: str, html: str, attachments=None):
    host = _env('SMTP_HOST'); port = int(_env('SMTP_PORT','587')); user = _env('SMTP_USER'); pwd = _env('SMTP_PASS')
    from_addr = _env('FROM_EMAIL', user or 'no-reply@example.com')
    if not host or not user or not pwd:
        return {'ok': False, 'info': 'SMTP not configured in .env'}
    msg = MIMEMultipart()
    msg['From'] = from_addr; msg['To'] = to_addr; msg['Subject'] = subject
    msg.attach(MIMEText(html, 'html', 'utf-8'))
    for path in (attachments or []):
        try:
            ctype, encoding = mimetypes.guess_type(path)
            maintype, subtype = (ctype or 'application/octet-stream').split('/',1)
            with open(path, 'rb') as f:
                part = MIMEBase(maintype, subtype); part.set_payload(f.read()); encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(path))
            msg.attach(part)
        except Exception:
            pass
    try:
        with smtplib.SMTP(host, port) as s:
            s.starttls()
            s.login(user, pwd)
            s.sendmail(from_addr, [to_addr], msg.as_string())
        return {'ok': True, 'info': 'sent'}
    except Exception as e:
        return {'ok': False, 'info': str(e)}

# --- Token helpers used by routers/accept.py -------------------------------

def _sign(data: str, ttl: int = 7*24*3600) -> str:
    """Create a simple HMAC token with expiry."""
    secret = _env('QUOTE_ACCEPT_SECRET', 'changeme')
    exp = int(time.time()) + int(ttl)
    payload = f"{data}.{exp}".encode('utf-8')
    sig = hmac.new(secret.encode('utf-8'), payload, hashlib.sha256).hexdigest()
    return f"{data}.{exp}.{sig}"

def verify_token(token: str) -> bool:
    """Verify HMAC token created by _sign().
    Expected format: <data>.<exp>.<sig>
    """
    try:
        data, exp_str, sig = token.rsplit('.', 2)
        exp = int(exp_str)
        if exp < int(time.time()):
            return False
        secret = _env('QUOTE_ACCEPT_SECRET', 'changeme')
        payload = f"{data}.{exp}".encode('utf-8')
        expected = hmac.new(secret.encode('utf-8'), payload, hashlib.sha256).hexdigest()
        return hmac.compare_digest(expected, sig)
    except Exception:
        return False
