import os, sys, importlib.util, pathlib
BASE_DIR = os.path.dirname(__file__)
# Make sure local dir is on sys.path
if BASE_DIR not in sys.path: sys.path.insert(0, BASE_DIR)
if os.getcwd() not in sys.path: sys.path.insert(0, os.getcwd())

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from dotenv import load_dotenv

def _import_local(modname: str, filename: str):
    path = os.path.join(BASE_DIR, filename)
    if not os.path.exists(path):
        raise ImportError(f"{filename} not found next to app.py")
    spec = importlib.util.spec_from_file_location(modname, path)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    sys.modules[modname] = module
    return module

# Import storage robustly
try:
    import storage  # type: ignore
except Exception as _e:
    storage = _import_local("storage", "storage.py")  # type: ignore

load_dotenv(override=False)
storage.init_db()

app = FastAPI(title="Voerman Dashboard API")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

# Routers
try:
    from routers import ingest, extract, pricing, emailer, accept, messages, pipeline
except Exception as _e:
    # Fallback: load routers individually by path
    ingest = _import_local("routers.ingest", os.path.join("routers","ingest.py"))
    extract = _import_local("routers.extract", os.path.join("routers","extract.py"))
    pricing = _import_local("routers.pricing", os.path.join("routers","pricing.py"))
    emailer = _import_local("routers.emailer", os.path.join("routers","emailer.py"))
    accept = _import_local("routers.accept", os.path.join("routers","accept.py"))
    messages = _import_local("routers.messages", os.path.join("routers","messages.py"))
    pipeline = _import_local("routers.pipeline", os.path.join("routers","pipeline.py"))

app.include_router(ingest.router, prefix="/ingest", tags=["ingest"])  # /ingest/test
app.include_router(extract.router, prefix="/extract", tags=["extract"]) # /extract
app.include_router(pricing.router, prefix="", tags=["pricing"])         # /quote
app.include_router(emailer.router, prefix="/email", tags=["email"])     # /email/preview, /email/send
app.include_router(messages.router, prefix="", tags=["messages"])       # /messages, /messages/{id}
app.include_router(pipeline.router, prefix="", tags=["pipeline"])       # /pipeline/generate, /pipeline/send

# Static: serve /out for previews
OUT_DIR = os.environ.get("OUT_DIR","out")
os.makedirs(OUT_DIR, exist_ok=True)
app.mount("/out", StaticFiles(directory=OUT_DIR), name="out")


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard():
    path = os.path.join(BASE_DIR, "templates", "dashboard.html")
    with open(path, 'r', encoding='utf-8') as f:
        return HTMLResponse(f.read())

@app.get("/health")
def health(): return {"ok": True}
