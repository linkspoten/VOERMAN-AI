# Voerman E2E — py38 + Email Send

## Snel starten
- `voerman_one.bat` → installeert deps + demo.
- `voerman_one.bat serve` → start API.

## Email testen (/email/send)
1) Zet SMTP-variabelen in `.env` (Gmail: gebruik een app password):
```
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=yourgmail@gmail.com
SMTP_PASS=your-app-password
FROM_EMAIL=quotes@example.com
```
2) Flow API:
- `POST /ingest/test` → noteer `id`
- `POST /extract` met `{ "message_id": "<id>" }`
- `POST /quote` met de `QuoteRequest` uit extract
- `POST /email/send` met body:
```json
{
  "to": "jouwadres@example.com",
  "language": "nl",
  "options": [ ... de optie(s) uit /quote ... ],
  "customer_name": "Klantnaam",
  "quote_id": "q_demo",
  "subject": "Offerte – Amsterdam → Montreal (LCL)"
}
```
De PDF's in `pdf_path` worden automatisch als bijlage toegevoegd.


## Dashboard gebruiken
1) Dubbelklik `voerman_one.bat` → de server start en opent **/dashboard**.
2) Links kun je **testmails importeren** en de inbox verversen/aanvinken.
3) Rechts zie je de mail, klik **Genereer AI response + PDF** → preview + opties verschijnen.
4) Vul ontvanger in en klik **Verzend e‑mail** (vereist SMTP in `.env`).

Tip: `voerman_one.bat demo` draait de oude offline demo (zonder UI) en zet artefacten in `out/`.


### Voerman engine koppelen
In `.env` kun je dit zetten:
```
ENGINE_MODE=real
PRICING_TOOL_PATH=Voerman_Quote_Studio_MQ26_P0PATCH.py
PRICING_EXCEL_PATH=tarieven.xlsx
```
Sla je Excel als `.xlsx`. Voor demo: `ENGINE_MODE=placeholder`.
