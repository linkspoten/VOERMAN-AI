import os, json, time
import storage
from extractor import extract_from_unified
from models_contracts import QuoteRequest
import pricing_core
from email_service import render_preview, _sign_token, verify_token

os.environ.setdefault('OUT_DIR','out')
if not os.path.exists('out'):
    os.makedirs('out', exist_ok=True)

print('[1/6] Init DB...')
storage.init_db()

print('[2/6] Ingest message...')
msg = {'id':'umsg_demo','source':'web','sender':{'email':'customer@example.com'},'subject':'Verhuizing Amsterdam naar Montreal','body':'Hi, graag prijs voor 12.5 m3 en 800 kg.','attachments':[],'language':'nl','timestamp': time.strftime('%Y-%m-%dT%H:%M:%SZ'),'thread_id': None,'message_id': None}
storage.insert_message(msg)
print('   message_id:', msg['id'])

print('[3/6] Extract...')
res = extract_from_unified(msg)
qr = res.request

print('[4/6] Pricing...')
options = []
for m in qr.modes:
    qr_single = qr.copy(update={'modes':[m]})
    opt_list = pricing_core.generate_quote(qr_single)
    for o in opt_list:
        o['mode'] = m
        options.append(o)
print('   options:', json.dumps(options, indent=2, ensure_ascii=False))

print('[5/6] Email preview...')
qid = 'q_demo'
html = render_preview(qr.language, options, 'Klant', res.clarifying_questions, 'Met vriendelijke groet,\nVoerman Team', qid)
with open('out/email_preview.html','w',encoding='utf-8') as f:
    f.write(html)
print('   wrote out/email_preview.html')

print('[6/6] Accept token...')
token = _sign_token(qid)
assert verify_token(token) == qid
with open('out/handoff_simulated.txt','w',encoding='utf-8') as f:
    f.write('ACCEPTED '+qid+' '+token)
print('OK — check out/ for PDF + preview')
