import tkinter as tk

# -*- coding: utf-8 -*-
"""
Quotation GUI – Origin / Freight / Destination + FCL Freight (lanes) + ROAD freight
v4.4 – Adds ROAD freight support:
- Domestic (NL↔NL): single consolidated line "Domestic door-to-door services (ROAD)",
  priced as: Origin(ROAD to Nootdorp) + Destination(ROAD to Nootdorp). No separate freight line.
- International (outside NL): three lines (if selected): Origin services (to Nootdorp),
  Road freight (origin↔destination) with km * €/km (Combined or Direct), Destination services (from Nootdorp).
- UI: ROAD options panel with type (Combined/Direct) and €/km editable (defaults 1.10 / 2.10).
Other parts kept as in v4.3. AI parser unchanged.
"""
import os
import copy, sys, json, traceback, requests, math, re, time
from dataclasses import dataclass
from typing import Optional, Tuple, List, Dict
from datetime import datetime

try:
    import pandas as pd
    PANDAS_OK = True
except Exception:
    PANDAS_OK = False

# --- ReportLab (PDF) ---
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, Image
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    REPORTLAB_OK = True
    from reportlab.lib.utils import ImageReader
except Exception:
    REPORTLAB_OK = False

from tkinter import ttk, messagebox, filedialog
from tkinter import scrolledtext

# ---- Minimal pure-Tkinter ToggleSwitch (no external deps) ------------------
class ToggleSwitch(ttk.Frame):
    def __init__(self, master, *, variable=None, command=None, text="",
                 width=40, height=20, on_color="#2e7d32",
                 off_color="#c9c9c9", knob_color="#ffffff", **kw):
        super().__init__(master, **kw)
        self.var = variable or tk.BooleanVar(value=False)
        self.command = command
        self.on_color, self.off_color, self.knob_color = on_color, off_color, knob_color
        self._sw_w, self._sw_h = int(width), int(height)
        self._pad = 2
        self.columnconfigure(1, weight=1)

        self._canvas = tk.Canvas(self, width=self._sw_w, height=self._sw_h,
                                 highlightthickness=0, bd=0, bg=self._bg())
        self._canvas.grid(row=0, column=0, padx=(0,4))
        self._label = ttk.Label(self, text=text)
        self._label.grid(row=0, column=1, sticky="w", padx=(2,0))

        self.configure(takefocus=1)
        for w in (self, self._canvas, self._label):
            w.bind("<Button-1>", self._toggle)
        self.bind("<space>", self._toggle)
        self.bind("<Button-1>", lambda e: self.focus_set(), add="+")
        self._canvas.bind("<Button-1>", lambda e: self.focus_set(), add="+")
        self._label.bind("<Button-1>", lambda e: self.focus_set(), add="+")
        self._canvas.bind("<Configure>", lambda e: self._draw())
        self.var.trace_add("write", lambda *_: self._draw())
        self._draw()

    def _bg(self):
        try:
            return self.master.cget("background")
        except Exception:
            return self.winfo_toplevel().cget("background")

    def _draw(self):
        c = self._canvas; c.delete("all")
        w, h, pad = self._sw_w, self._sw_h, self._pad
        r = h / 2
        on = bool(self.var.get())
        track = self.on_color if on else self.off_color

        # Track pill
        c.create_oval(0, 0, h, h, fill=track, outline=track)
        c.create_oval(w-h, 0, w, h, fill=track, outline=track)
        c.create_rectangle(r, 0, w-r, h, fill=track, outline=track)

        # Knob
        dia = h - 2*pad
        x0 = (w - h + pad) if on else pad
        c.create_oval(x0, pad, x0 + dia, pad + dia,
                      fill=self.knob_color, outline="#cfcfcf")

        try:
            has_focus = (self.focus_get() is self)
        except Exception:
            has_focus = False
        if has_focus:
            c.create_rectangle(1, 1, w-1, h-1, outline="#7aa7ff", width=1, dash=(2,2))
        c.configure(cursor="hand2")

    def _toggle(self, *_):
        self.var.set(not self.var.get())
        if callable(self.command):
            try:
                self.command()
            except TypeError:
                self.command(self.var.get())

    def get(self): return bool(self.var.get())
    def set(self, val): self.var.set(bool(val))
# ---------------------------------------------------------------------------


# Optional modern ttk theme (sv-ttk)
try:
    import sv_ttk
    SVTTK_AVAILABLE = True
except Exception:
    SVTTK_AVAILABLE = False

# --- Optional modern theming (ttkbootstrap) ---
try:
    import ttkbootstrap as tb
    BOOTSTRAP_AVAILABLE = True
except Exception:
    BOOTSTRAP_AVAILABLE = False

# --- Geo ---
from geopy.geocoders import Nominatim
from geopy.distance import geodesic

# --- utils ---
from xml.sax.saxutils import escape

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass



def _safe_read_excel(path):
    if not (globals().get("PANDAS_OK", False)):
        try:
            from tkinter import messagebox
            messagebox.showerror("Missing dependency", "Pandas is niet geïnstalleerd. Voer uit:\npy -m pip install pandas openpyxl")
        except Exception:
            pass
        return None
    try:
        return pd.read_excel(path)
    except Exception as e:
        try:
            from tkinter import messagebox
            messagebox.showerror("Excel fout", f"Kon Excel niet lezen:\n{e}")
        except Exception:
            pass
        return None


# --- EU country set for VAT logic ---
EU_COUNTRIES = {
    "AT","BE","BG","HR","CY","CZ","DK","EE","FI","FR","DE","GR","HU",
    "IE","IT","LV","LT","LU","MT","NL","PL","PT","RO","SK","SI","ES","SE"
}
def should_show_vat(*, is_private: bool, is_agent: bool | None, checkbox: bool, origin_iso: str, dest_iso: str) -> bool:
    """Centrale VAT-beslisboom. Alleen tonen wanneer user het wil (checkbox)
    én private klant én intra-EU. `is_agent` wordt genegeerd als `is_private=True`,
    maar kan later gebruikt worden voor aanvullende uitzonderingen."""
    try:
        o = (origin_iso or "").strip().upper()
        d = (dest_iso or "").strip().upper()
    except Exception:
        o = ""; d = ""
    intra_eu = (o in EU_COUNTRIES) and (d in EU_COUNTRIES)
    return bool(checkbox and is_private and intra_eu)


# ---- Country name <-> ISO2 helpers for UI (full world list) ----
try:
    import pycountry
    ISO2_TO_NAME = {c.alpha_2.upper(): c.name for c in pycountry.countries}
    COUNTRY_NAMES = sorted({c.name for c in pycountry.countries})
except Exception:
    ISO2_TO_NAME = {
        "NL": "Netherlands", "BE": "Belgium", "DE": "Germany", "FR": "France",
        "GB": "United Kingdom", "IE": "Ireland", "LU": "Luxembourg",
        "ES": "Spain", "PT": "Portugal", "IT": "Italy", "SE": "Sweden",
        "NO": "Norway", "DK": "Denmark", "FI": "Finland", "PL": "Poland",
        "CZ": "Czechia", "AT": "Austria", "CH": "Switzerland",
        "US": "United States", "CA": "Canada", "CN": "China", "JP": "Japan",
        "AU": "Australia", "NZ": "New Zealand", "TR": "Türkiye", "AE": "United Arab Emirates",
    }
    COUNTRY_NAMES = sorted(ISO2_TO_NAME.values())
NAME_TO_ISO2 = {v: k for k, v in ISO2_TO_NAME.items()}

def code_to_name(code: str) -> str:
    return ISO2_TO_NAME.get((code or "").upper(), code or "")

def name_to_code(name: str) -> str:
    return NAME_TO_ISO2.get(name, (name or "")[:2].upper())
# ---- end helpers ----


LOG_PATH = os.path.join(os.path.abspath(os.path.dirname(sys.argv[0] or __file__)), "ai_debug.log")

def log(msg: str):
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().isoformat(timespec='seconds')}] {msg}\n")
    except Exception:
        pass

def _sdk_versions():
    v = {}
    try:
        import openai as _o
        v["openai"] = getattr(_o, "__version__", "unknown")
    except Exception:
        v["openai"] = "not-installed"
    return v

def _load_new_client():
    try:
        from openai import OpenAI  # v1.x
        return OpenAI()
    except Exception as e:
        log(f"New SDK client failed: {e}")
        return None

def _has_legacy():
    try:
        import openai as _o
        return hasattr(_o, "ChatCompletion")
    except Exception:
        return False

# =================== CONFIG ===================
WAREHOUSE_OPERATION = "Nootdorp"
WAREHOUSE_LOCATION  = "Nootdorp, Netherlands"

EXCEL_PAD        = "tarieven.xlsx"

# Default output directory for PDFs
try:
    APP_DIR = os.path.abspath(os.path.dirname(sys.argv[0] or __file__))
except Exception:
    APP_DIR = os.getcwd()
OUTPUT_DIR = os.path.join(APP_DIR, "output")
SERVICES_SHEET   = 0
FCL_SHEET        = ""
DEST_ONLY_SHEET = "DestOnlyCharges"

BEDRIJFSNAAM  = "Your Company B.V."
BEDRIJF_ADRES = "Example Street 1, 1234 AB City"
BEDRIJF_EMAIL = "info@yourcompany.com"
BEDRIJF_TEL   = "+31 (0)12 345 6789"

# --- Branding presets ---
BRANDS = {
    "Voerman": {
        "name": "Voerman International B.V.",
        "addr": "Reflectiestraat 2, 2631 RV Nootdorp, NL",
        "email": "info@voerman.com",
        "tel": "+31 (0)70 301 7700",
        "logo": ""  # optioneel: pad naar Voerman-logo (PNG/JPG)
    },
    "Transpack": {
        "name": "Transpack B.V.",
        "addr": "Reflectiestraat 2, 2631 RV Nootdorp, NL",
        "email": "info@transpack.nl",
        "tel": "+31 (0)70 301 7800",
        "logo": ""  # optioneel: pad naar Transpack-logo (PNG/JPG)
    },
}

COLS = {
    "operation": "Operation",
    "type": "Type",
    "mode": "Mode",
    "d_start": "Distance start",
    "d_end": "Distance end",
    "v_min": "Min. Value",
    "v_max": "Max. Value",
    "rate_per_cbm": "Flexibel( rate per cbm)",
    "flat": "Flat rate in EUR",
    "rate_type": "Rate type",
    "port": "PORT CODE",
}

FCL_LANE_COLS = {
    "opol":  ["Origin port", "Origin", "POL", "Load port"],
    "opolc": ["Origin port code", "Origin Code", "POL code", "POL Code", "Origin Code (UN/LOCODE)"],
    "dpod":  ["Destination port", "Destination", "POD", "Discharge port"],
    "dpodc": ["Destination port code", "Destination Code", "POD code", "POD Code", "Destination Code (UN/LOCODE)"],
    "r20":   ["20ft", "20 ft", "20'", "20-FT", "20F", "20FT"],
    "r40":   ["40ft", "40 ft", "40'", "40-FT", "40F", "40FT"],
    "r40hq": ["40ft HQ", "40 HQ", "40HC", "40'HC", "40HQ", "40 H Q"],

    "lcl": ["LCL (per cbm)", "LCL per cbm", "LCL/cbm", "LCL", "LCL per m3", "LCL (per m3)", "LCL price", "LCL rate"],}

VALID_MODES = ("FCL", "LCL", "AIR", "ROAD", "GROUPAGE")

DEST_ONLY_COLS = {
    "country": ["Country", "Land", "Destination Country", "Country Name", "Dest Country", "Bestemming land"],
    "mode":    ["Mode", "Modality"],
    "rate_20": ["Rate_20FT", "20ft", "20 ft", "DTHC 20", "DTHC 20FT"],
    "rate_40": ["Rate_40FT", "40ft", "40 ft", "DTHC 40", "DTHC 40FT"],
    "rate_40hq":["Rate_40HQ", "40ft HQ", "40 HQ", "DTHC 40HQ", "DTHC 40 HC"],
    "rate_lcl":["Rate_LCL_per_cbm", "LCL (per cbm)", "LCL per cbm", "NVOCC per cbm", "LCL per m3"],
    "rate_air":["Rate_AIR_per_kg", "AIR per kg", "ATHC per kg", "Air per kg", "Air (per kg)"],
    "charge_name": ["Charge", "Charge name", "Naam", "Omschrijving"],
}

ORS_API_KEY = os.getenv("ORS_API_KEY")

FCL_CAPACITY = {"20FT": 30.0, "40FT": 60.0, "40HQ": 67.0}
PRETTY_TYPE = {"20FT": "20 ft container", "40FT": "40 ft container", "40HQ": "40 ft HQ container"}

# ROAD defaults
DEFAULT_ROAD_RATE_COMBINED = 1.10  # €/km
DEFAULT_ROAD_RATE_DIRECT   = 2.10  # €/km

# =================== HELPERS ===================
def eur(n: float) -> str:
    s = f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"€ {s}"

def _sanitize_filename(name: str) -> str:
    import re
    s = re.sub(r"[^A-Za-z0-9 _\-\.]+" , "", name or "")
    s = s.strip()
    return s or "output"

def _logo_flowable(logo_path: str, max_w_mm=60, max_h_mm=24):
    try:
        ir = ImageReader(logo_path)
        iw, ih = ir.getSize()
        max_w = max_w_mm * mm
        max_h = max_h_mm * mm
        scale = min(max_w / float(iw), max_h / float(ih))
        return Image(logo_path, width=float(iw)*scale, height=float(ih)*scale)
    except Exception:
        return None

def fmt_qty(qty, unit: str | None = None) -> str:
    if qty is None: return ""
    return f"{qty:g} {unit}" if unit else f"{qty:g}"

def _val(x, default="-"):
    try:
        if x is None: return default
        if isinstance(x, float) and pd.isna(x): return default
        if x == "": return default
        return x
    except Exception:
        return default

def is_unlocode(s: str) -> bool:
    if not isinstance(s, str): return False
    return bool(re.fullmatch(r"[A-Z]{2}[A-Z0-9]{3}", s.strip().upper()))

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip()).lower()

def detect_cols(df: pd.DataFrame, candidates: Dict[str, List[str]]) -> Dict[str, str]:
    norm_map = {_norm(c): c for c in df.columns}
    out = {}
    for key, opts in candidates.items():
        found = None
        norm_opts = [_norm(o) for o in opts]
        for o in norm_opts:
            if o in norm_map: found = norm_map[o]; break
        if not found:
            for col_norm, col_real in norm_map.items():
                if any(o in col_norm or col_norm in o for o in norm_opts):
                    found = col_real; break
        if not found:
            raise ValueError(f"Column for '{key}' not found. Expected one of: {opts}")
        out[key] = found
    return out

# =================== EMAIL PARSER (rules) ===================
import re

def parse_rfq_text(text: str) -> dict:
    """
    Rule-based RFQ parser (no AI).
    - Detects mode, services, origin/destination, volume (m3 or cf), POL/POD.
    - CF → m3 conversion; ranges like '187–200 cf' use the upper bound.
    """
    out = {}
    if not text or not text.strip():
        return out

    t = text.replace("\r\n", "\n")
    t_clean = re.sub(r"[ \t]+", " ", t)
    t_low = t_clean.lower()

    # ---------------------------
    # 1) MODE
    # ---------------------------
    if re.search(r"\blcl\b", t_low):
        out["mode"] = "LCL"
    elif re.search(r"\bfcl\b|\bfull\s*container\b", t_low):
        out["mode"] = "FCL"
    elif re.search(r"\bair\b|\bairfreight|\bair freight", t_low):
        out["mode"] = "AIR"
    elif re.search(r"\broad\b|\btruck(ing)?\b", t_low):
        out["mode"] = "ROAD"
    elif re.search(r"\bgroupage\b", t_low):
        out["mode"] = "GROUPAGE"
    elif re.search(r"\btransport\s*:\s*sea\b|\bby sea\b|\bocean freight\b|\bsea freight\b", t_low):
        pass

    # ---------------------------
    # 2) SERVICES (heuristics)
    # ---------------------------
    services = set()
    if re.search(r"\bdestination service\b|\bdthc\b|\bpoe\b|\bdelivery address\b|\bfinal address\b|\bdeliver(y)?\b", t_low):
        services.add("destination")
    if re.search(r"\borigin service\b|\borigin address\b|\bpick\s*up\b|\bpickup\b|\bcollection\b|\bpack(ing)?\b", t_low):
        services.add("origin")
    if re.search(r"\bfreight\b|\bvia\s*lcl\b|\bport to port\b|\bocean freight\b|\bsea freight\b|\broad freight\b", t_low):
        services.add("freight")
    if re.search(r"\bdoor\s*to\s*door\b|\bfrom\b.+\bto\b", t_low):
        services.update({"origin", "freight", "destination"})
    if services:
        out["services"] = sorted(list(services))

    # ---------------------------
    # 3) ORIGIN / DESTINATION
    # ---------------------------
    def grab_after(label_regex: str):
        m = re.search(label_regex, t_clean, re.I)
        if not m:
            return None
        s = m.group(1).strip()
        s = re.split(r"\s[–-]\s", s, maxsplit=1)[0].strip()
        return s

    dest = (grab_after(r"destination(?: address)?\s*:\s*(.+)") or
            grab_after(r"delivery address\s*:\s*(.+)"))
    if dest:
        out["destination_location"] = dest

    origin = (grab_after(r"origin(?: address)?\s*:\s*(.+)") or
              grab_after(r"pick[\s-]*up\s*:\s*(.+)") or
              grab_after(r"collection\s*:\s*(.+)"))
    if origin:
        out["origin_location"] = origin

    # ---------------------------
    # 4) POL / POD
    # ---------------------------
    m_pol = re.search(r"\bpol\b\s*[:\-]\s*([A-Z0-9]{3,6})", t, re.I)
    m_pod = re.search(r"\bpod\b\s*[:\-]\s*([A-Z0-9]{3,6})", t, re.I)
    if m_pol: out["pol"] = m_pol.group(1).upper()
    if m_pod: out["pod"] = m_pod.group(1).upper()

    # ---------------------------
    # 5) VOLUME
    # ---------------------------
    volumes_m3 = []
    for m in re.finditer(r"(\d+(?:[.,]\d+)?)\s*(?:m3|m\xb3|m\^3|cbm|cubic\s*meters?)\b", t_low, re.I):
        try:
            v = float(m.group(1).replace(",", ".")); volumes_m3.append(v)
        except Exception: pass

    CF_TO_M3 = 0.0283168
    for m in re.finditer(r"(\d+(?:[.,]\d+)?)\s*(?:–|-|to|upto|up\s*to)\s*(\d+(?:[.,]\d+)?)\s*(?:cf|cft|ft3|ft\^3|cu\.?\s*ft|cubic\s*feet)\b", t_low, re.I):
        try:
            upper = float(m.group(2).replace(",", ".")); volumes_m3.append(upper * CF_TO_M3)
        except Exception: pass
    for m in re.finditer(r"(\d+(?:[.,]\d+)?)\s*(?:cf|cft|ft3|ft\^3|cu\.?\s*ft|cubic\s*feet)\b", t_low, re.I):
        try:
            v = float(m.group(1).replace(",", ".")); volumes_m3.append(v * CF_TO_M3)
        except Exception: pass

    if volumes_m3:
        out["volume_cbm"] = max(volumes_m3)

    return out


# =================== AI EMAIL PARSER with fallbacks ===================
def ai_parse_rfq_text(text: str, model: Optional[str] = None) -> dict:
    if not text or not text.strip():
        return {}
    key = os.getenv("OPENAI_API_KEY")
    if not key:
        log("AI parse aborted: no OPENAI_API_KEY")
        return {}

    model = model or os.getenv("OPENAI_MODEL", "gpt-4o-mini")
    log(f"AI parse start – model={model}, sdk={_sdk_versions()}")

    sys_prompt = (
        "You extract shipping RFQ details from emails and output JSON. "
        "Use these keys when possible: "
        "mode (FCL/LCL/AIR/ROAD/GROUPAGE), services [origin|freight|destination], "
        "origin_location, destination_location, pol, pod, volume_cbm (number), "
        "containers {20FT,40FT,40HQ}, container_choice (20FT|40FT|40HQ|auto|unknown), "
        "incoterm, reference, contact {company,name,email,phone}, confidence (0..1). "
        "If value is unknown, omit the key."
    )

    # 1) New SDK + JSON Schema
    try:
        client = _load_new_client()
        if client is not None:
            json_schema = {
                "name": "rfq_extract",
                "schema": {
                    "type": "object",
                    "properties": {
                        "mode": {"type":"string","enum":["FCL","LCL","AIR","ROAD","GROUPAGE"]},
                        "services": {"type":"array","items":{"type":"string","enum":["origin","freight","destination"]}},
                        "origin_location": {"type":"string"},
                        "destination_location": {"type":"string"},
                        "pol": {"type":"string"},
                        "pod": {"type":"string"},
                        "volume_cbm": {"type":"number"},
                        "containers": {
                            "type":"object",
                            "properties": {
                                "20FT":{"type":"integer","minimum":0},
                                "40FT":{"type":"integer","minimum":0},
                                "40HQ":{"type":"integer","minimum":0}
                            },
                            "additionalProperties": False
                        },
                        "container_choice": {"type":"string","enum":["20FT","40FT","40HQ","auto","unknown"]},
                        "incoterm": {"type":"string"},
                        "reference": {"type":"string"},
                        "contact": {
                            "type":"object",
                            "properties": {
                                "company":{"type":"string"},
                                "name":{"type":"string"},
                                "email":{"type":"string"},
                                "phone":{"type":"string"}
                            },
                            "additionalProperties": False
                        },
                        "confidence": {"type":"number","minimum":0,"maximum":1}
                    },
                    "required": [],
                    "additionalProperties": False
                },
                "strict": True
            }
            r = client.chat.completions.create(
                model=model,
                temperature=0,
                messages=[
                    {"role":"system","content":sys_prompt},
                    {"role":"user","content":text}
                ],
                response_format={
                    "type":"json_schema",
                    "json_schema": json_schema
                }
            )
            content = r.choices[0].message.content
            data = json.loads(content) if isinstance(content, str) else {}
            if isinstance(data, dict) and data:
                log("AI parse: new SDK + schema success")
                return data
            else:
                log("AI parse: new SDK + schema returned empty or non-dict")
    except Exception as e:
        log(f"New SDK + schema failed: {e}")

    # 2) New SDK + JSON object mode
    try:
        if client is None:
            client = _load_new_client()
        if client is not None:
            r = client.chat.completions.create(
                model=model,
                temperature=0,
                messages=[
                    {"role":"system","content":sys_prompt + " Return ONLY valid JSON object."},
                    {"role":"user","content":text}
                ],
                response_format={"type":"json_object"}
            )
            content = r.choices[0].message.content
            data = json.loads(content) if isinstance(content, str) else {}
            if isinstance(data, dict) and data:
                log("AI parse: new SDK + json_object success")
                return data
            else:
                log("AI parse: new SDK + json_object returned empty")
    except Exception as e:
        log(f"New SDK + json_object failed: {e}")

    # 3) Legacy SDK (v0.x)
    if _has_legacy():
        try:
            import openai
            openai.api_key = key
            r = openai.ChatCompletion.create(
                model=model,
                temperature=0,
                messages=[
                    {"role":"system","content":sys_prompt + " Return ONLY valid JSON object."},
                    {"role":"user","content":text}
                ]
            )
            content = r["choices"][0]["message"]["content"]
            m = re.search(r"\{.*\}", content, re.S)
            content_json = m.group(0) if m else content
            data = json.loads(content_json)
            if isinstance(data, dict) and data:
                log("AI parse: legacy SDK success")
                return data
        except Exception as e:
            log(f"Legacy SDK failed: {e}")

    log("AI parse failed – returning {}")
    return {}

# =================== .env helper ===================
def write_env_file(env_path: str, key_value: str, model_value: str = "gpt-4o-mini") -> None:
    lines = []
    if os.path.isfile(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
    def set_or_add(name, value):
        nonlocal lines
        pattern = re.compile(rf"^{re.escape(name)}\s*=")
        found = False
        new_lines = []
        for ln in lines:
            if pattern.match(ln):
                new_lines.append(f"{name}={value}")
                found = True
            else:
                new_lines.append(ln)
        if not found:
            new_lines.append(f"{name}={value}")
        lines = new_lines
    set_or_add("OPENAI_API_KEY", key_value.strip())
    set_or_add("OPENAI_MODEL", model_value.strip() or "gpt-4o-mini")
    with open(env_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

# =================== DATAFUNCTIES ===================
@dataclass
class LineParams:
    label_for_pdf: str
    type_label: str
    location_for_operation: str
    mode: str
    volume_cbm: float
    distance_km: float

def lees_services_sheet(pad: str, sheet) -> pd.DataFrame:
    df = pd.read_excel(pad, sheet_name=sheet, header=0)
    missing = [v for v in COLS.values() if v not in df.columns]
    if missing: raise ValueError(f"Missing columns in services sheet: {missing}")
    df = df[list(COLS.values())].copy()
    for c in [COLS["d_start"], COLS["d_end"], COLS["v_min"], COLS["v_max"], COLS["rate_per_cbm"], COLS["flat"]]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in [COLS["operation"], COLS["type"], COLS["mode"], COLS["rate_type"], COLS["port"]]:
        df[c] = df[c].astype(str).str.strip()
    return df

def geocode_raw(addr: str):
    geolocator = Nominatim(user_agent="quotation_gui_v44")
    return geolocator.geocode(addr, addressdetails=True)

def geocode(addr: str) -> Tuple[float, float]:
    loc = geocode_raw(addr)
    if not loc: raise ValueError(f"Address not found: {addr}")
    return (loc.latitude, loc.longitude)

def geocode_country(addr: str) -> Tuple[Tuple[float,float], str, str]:
    """Return ((lat,lon), country_code_lower, country_name)."""
    loc = geocode_raw(addr)
    if not loc: raise ValueError(f"Address not found: {addr}")
    cc = ""
    name = ""
    try:
        addr_dict = loc.raw.get("address", {}) if hasattr(loc, "raw") else {}
        cc = (addr_dict.get("country_code") or "").lower()
        name = addr_dict.get("country") or ""
    except Exception:
        pass
    if not cc and hasattr(loc, "address"):
        if "Netherlands" in loc.address or "Nederland" in loc.address: cc = "nl"; name = "Netherlands"
    return (loc.latitude, loc.longitude), cc, name

def ors_route_distance_km(a: Tuple[float,float], b: Tuple[float,float]) -> Optional[float]:
    if not ORS_API_KEY: return None
    url = "https://api.openrouteservice.org/v2/directions/driving-car"
    headers = {"Authorization": ORS_API_KEY, "Content-Type": "application/json"}
    body = {"coordinates": [[a[1], a[0]], [b[1], b[0]]]}
    try:
        r = requests.post(url, headers=headers, data=json.dumps(body), timeout=20)
        if r.status_code != 200: return None
        meters = r.json()["features"][0]["properties"]["segments"][0]["distance"]
        return meters / 1000.0
    except Exception:
        return None

def afstand_km_via_warehouse(free_addr: str, warehouse_addr: str) -> Tuple[float, str]:
    a = geocode(free_addr); b = geocode(warehouse_addr)
    km = ors_route_distance_km(a, b)
    return (geodesic(a, b).km * 1.25, "Geodesic × 1.25 (approx.)") if km is None else (km, "OpenRouteService (route)")

def road_distance_between_addrs(origin_addr: str, dest_addr: str) -> Tuple[float, str]:
    a = geocode(origin_addr); b = geocode(dest_addr)
    km = ors_route_distance_km(a, b)
    return (geodesic(a, b).km * 1.25, "Geodesic × 1.25 (approx.)") if km is None else (km, "OpenRouteService (route)")

def match_service_rij(df: pd.DataFrame, p: LineParams) -> pd.Series:
    heeft_op = (df[COLS["operation"]].str.lower() == p.location_for_operation.lower().strip()).any()
    base = (
        (df[COLS["type"]].str.lower() == p.type_label.lower()) &
        (df[COLS["mode"]].str.lower() == p.mode.lower()) &
        (df[COLS["d_start"]] <= p.distance_km) &
        (p.distance_km < df[COLS["d_end"]])
    )
    if heeft_op:
        base &= (df[COLS["operation"]].str.lower() == p.location_for_operation.lower().strip())
    v_max = df[COLS["v_max"]].fillna(float("inf"))
    m = base & (df[COLS["v_min"]] <= p.volume_cbm) & (p.volume_cbm < v_max)
    cand = df.loc[m]
    # FLAT preference at boundary: if volume equals a FLAT row's v_max, pick that FLAT
    try:
        __vol = float(p.volume_cbm)
    except Exception:
        __vol = None
    if __vol is not None and not cand.empty:
        vmax_num = pd.to_numeric(cand[COLS['v_max']], errors='coerce')
        flat_mask = (cand[COLS['rate_type']].str.strip().str.upper()=="FLAT") & (vmax_num == __vol)
        if flat_mask.any():
            return cand.loc[flat_mask].iloc[0]
    if cand.empty:
        flat = df.loc[base & (df[COLS["rate_type"]].str.upper() == "FLAT")]
        if not flat.empty: return flat.iloc[0]
        if heeft_op:
            base2 = (
                (df[COLS["type"]].str.lower() == p.type_label.lower()) &
                (df[COLS["mode"]].str.lower() == p.mode.lower()) &
                (df[COLS["d_start"]] <= p.distance_km) &
                (p.distance_km < df[COLS["d_end"]])
            )
            m2 = base2 & (df[COLS["v_min"]] <= p.volume_cbm) & (p.volume_cbm < v_max)
            cand2 = df.loc[m2]
            if not cand2.empty: return cand2.iloc[0]
        raise LookupError("No matching rate row found for the given parameters.")
    return cand.iloc[0]


def match_service_rij_strict_op(df: pd.DataFrame, p: LineParams) -> pd.Series:
    """
    Enforce exact Operation match for non-NL places. Ignores distance band and picks a single row that matches
    Type + Mode + Operation, with a preference for volume-appropriate rows or FLAT rate rows.
    Raises a LookupError with a clear message if the place name is not present in the tarieven sheet.
    """
    op_col = COLS["operation"]; typ_col = COLS["type"]; mode_col = COLS["mode"]
    vmin_col = COLS["v_min"]; vmax_col = COLS["v_max"]; rate_type_col = COLS["rate_type"]
    place = str(p.location_for_operation).strip()
    # candidates that match exact Operation (case-insensitive), Type, Mode
    m = (
        df[op_col].str.strip().str.casefold() == place.casefold()
    ) & (
        df[typ_col].str.strip().str.casefold() == p.type_label.strip().casefold()
    ) & (
        df[mode_col].str.strip().str.casefold() == p.mode.strip().casefold()
    )
    cand = df.loc[m]
    # FLAT preference at boundary: if volume equals a FLAT row's v_max, pick that FLAT
    try:
        __vol = float(p.volume_cbm)
    except Exception:
        __vol = None
    if __vol is not None and not cand.empty:
        vmax_num = pd.to_numeric(cand[COLS['v_max']], errors='coerce')
        flat_mask = (cand[COLS['rate_type']].str.strip().str.upper()=="FLAT") & (vmax_num == __vol)
        if flat_mask.any():
            return cand.loc[flat_mask].iloc[0].copy()
    if cand.empty:
        raise LookupError(f"Plaats '{place}' niet gevonden in de tarieven sheet voor Type='{p.type_label}' en Mode='{p.mode}'.")

    # Prefer rows where volume fits in v_min..v_max. If none, fall back to FLAT. Else just the first row.
    cand[vmin_col] = pd.to_numeric(cand[vmin_col], errors="coerce")
    cand[vmax_col] = pd.to_numeric(cand[vmax_col], errors="coerce")
    v_max = cand[vmax_col].fillna(float("inf"))
    vol_fit = (cand[vmin_col] <= p.volume_cbm) & (p.volume_cbm < v_max)
    if vol_fit.any():
        try:
            __vol = float(p.volume_cbm)
        except Exception:
            __vol = None
        if __vol is not None:
            vmax_num = pd.to_numeric(cand[vmax_col], errors='coerce')
            flat_mask = (cand[rate_type_col].str.strip().str.upper()=="FLAT") & (vmax_num == __vol)
            if flat_mask.any():
                return cand.loc[flat_mask].iloc[0]
        return cand.loc[vol_fit].iloc[0]
    flat_rows = cand[cand[rate_type_col].str.strip().str.upper() == "FLAT"]
    if not flat_rows.empty:
        return flat_rows.iloc[0]
    return cand.iloc[0]
def calc_service_prijs(row: pd.Series, volume_cbm: float) -> float:
    """Return price for origin/destination services; enforce min volume as flat."""
    rate_type = str(row.get(COLS['rate_type'], '')).strip().upper()
    # Flat value
    flat_val = None
    try:
        flat_val = float(row.get(COLS['flat'], ''))
    except Exception:
        pass
    # Minimum volume threshold
    vmin = None
    try:
        vmin = float(row.get(COLS['v_min'], ''))
    except Exception:
        pass
    # Per-cbm rate
    per_cbm = None
    try:
        per_cbm = float(row.get(COLS['rate_per_cbm'], ''))
    except Exception:
        pass
    # Logic
    if rate_type == 'FLAT' and flat_val is not None:
        return round(flat_val, 2)
    if vmin is not None and volume_cbm is not None and float(volume_cbm) <= vmin and flat_val is not None and flat_val > 0:
        return round(flat_val, 2)
    if per_cbm is not None and volume_cbm is not None:
        return round(per_cbm * float(volume_cbm), 2)
    return round(float(flat_val or 0.0), 2)

def lees_fcl_lanes(pad: str, sheet_name: str) -> pd.DataFrame:
    last_err = None; df = None; cols = None
    for hdr in range(0, 6):
        try:
            df_try = pd.read_excel(pad, sheet_name=sheet_name, header=hdr)
            cols_try = detect_cols(df_try, FCL_LANE_COLS)
            df, cols = df_try, cols_try; break
        except Exception as e:
            last_err = e; continue
    if df is None:
        raise last_err if last_err else ValueError("Could not detect FCL lane columns.")

    def num(x):
        if isinstance(x, str): x = x.replace(".", "").replace(",", ".")
        return pd.to_numeric(x, errors="coerce")

    out = pd.DataFrame({
        "OPORT": df[cols["opol"]].astype(str).str.strip(),
        "OPORT_CODE": df[cols["opolc"]].astype(str).str.strip().str.upper(),
        "DPORT": df[cols["dpod"]].astype(str).str.strip(),
        "DPORT_CODE": df[cols["dpodc"]].astype(str).str.strip().str.upper(),
        "RATE_20FT": num(df[cols["r20"]]),
        "RATE_40FT": num(df[cols["r40"]]),
        "RATE_40HQ": num(df[cols["r40hq"]]),
        "RATE_LCL_CBM": (num(df[cols["lcl"]]) if "lcl" in cols and cols["lcl"] in df.columns else pd.Series([pd.NA]*len(df)))
    })

    name2code = {}
    for _, r in out.iterrows():
        if is_unlocode(r["OPORT_CODE"]): name2code[r["OPORT"]] = r["OPORT_CODE"]
        if is_unlocode(r["DPORT_CODE"]): name2code[r["DPORT"]] = r["DPORT_CODE"]

    def map_name_to_code(name):
        if name in name2code: return name2code[name]
        for k, v in name2code.items():
            if str(k).lower() == str(name).lower(): return v
        return None

    bad_o = ~out["OPORT_CODE"].apply(is_unlocode)
    out.loc[bad_o, "OPORT_CODE"] = out.loc[bad_o, "OPORT"].map(map_name_to_code)

    bad_d = ~out["DPORT_CODE"].apply(is_unlocode)
    out.loc[bad_d, "DPORT_CODE"] = out.loc[bad_d, "DPORT"].map(map_name_to_code)

    return out

def auto_find_fcl_lanes_sheet(excel_path: str) -> str:
    xls = pd.ExcelFile(excel_path)
    best = None
    for name in xls.sheet_names:
        for hdr in range(0, 6):
            try:
                df_try = pd.read_excel(excel_path, sheet_name=name, header=hdr)
                cols = detect_cols(df_try, FCL_LANE_COLS); score = len(cols)
            except Exception:
                continue
            pref = 1 if any(k in name.upper() for k in ["FCL","LANE","FREIGHT","SEA"]) else 0
            key = (score, pref, -hdr)
            if (best is None) or (key > best[0]): best = (key, name, hdr)
    if not best:
        raise ValueError("Could not auto-detect the FCL lanes sheet.")
    return best[1]

def build_name_to_code(df_lanes: pd.DataFrame) -> dict:
    m = {}
    for _, r in df_lanes.iterrows():
        if is_unlocode(r["OPORT_CODE"]): m[str(r["OPORT"]).strip()] = r["OPORT_CODE"]
        if is_unlocode(r["DPORT_CODE"]): m[str(r["DPORT"]).strip()] = r["DPORT_CODE"]
    return m

def build_code_to_name(df_lanes: pd.DataFrame) -> dict:
    m = {}
    for _, r in df_lanes.iterrows():
        oc, dc = str(r.get("OPORT_CODE","")).strip().upper(), str(r.get("DPORT_CODE","")).strip().upper()
        on, dn = str(r.get("OPORT","")).strip(), str(r.get("DPORT","")).strip()
        if oc and on and oc not in m: m[oc] = on
        if dc and dn and dc not in m: m[dc] = dn
    return m

def resolve_port_input(inp: str, df_lanes: pd.DataFrame) -> str:
    s = (inp or "").strip()
    if is_unlocode(s): return s.upper()
    name2code = build_name_to_code(df_lanes)
    if s in name2code: return name2code[s]
    for k, v in name2code.items():
        if k.lower() == s.lower(): return v
    raise ValueError(f"Could not resolve UN/LOCODE for '{inp}'.")

def fcl_rates_for_lane(df_lanes: pd.DataFrame, o_code: str, d_code: str) -> Dict[str, float]:
    o_code = o_code.strip().upper(); d_code = d_code.strip().upper()
    sub = df_lanes[(df_lanes["OPORT_CODE"] == o_code) & (df_lanes["DPORT_CODE"] == d_code)]
    if sub.empty:
        sub = df_lanes[(df_lanes["OPORT_CODE"] == d_code) & (df_lanes["DPORT_CODE"] == o_code)]
    if sub.empty: raise LookupError(f"No FCL lane for {o_code} → {d_code}.")
    r = sub.iloc[0]; rates = {}
    if pd.notna(r["RATE_20FT"]): rates["20FT"] = float(r["RATE_20FT"])
    if pd.notna(r["RATE_40FT"]): rates["40FT"] = float(r["RATE_40FT"])
    if pd.notna(r["RATE_40HQ"]): rates["40HQ"] = float(r["RATE_40HQ"])
    if not rates: raise LookupError("No container rates filled for this lane.")
    return rates

def choose_fcl_combo(volume_cbm: float, rates: Dict[str, float]) -> Dict[str, int]:
    cap = FCL_CAPACITY; types = [t for t in ["40HQ","40FT","20FT"] if t in rates]
    if not types: raise LookupError("No usable FCL rates.")
    best_cost = math.inf; best = {"20FT":0,"40FT":0,"40HQ":0}
    max_hq = math.ceil(volume_cbm / cap.get("40HQ", 1e9)) + 3 if "40HQ" in types else 0
    for n_hq in range(0, max_hq+1):
        vol_after_hq = max(0.0, volume_cbm - n_hq*cap.get("40HQ",0))
        max_40 = math.ceil(vol_after_hq / cap.get("40FT", 1e9)) + 3 if "40FT" in types else 0
        for n_40 in range(0, max_40+1):
            vol_after_40 = max(0.0, vol_after_hq - n_40*cap.get("40FT",0))
            if "20FT" in types:
                n_20 = math.ceil(vol_after_40 / cap.get("20FT",1e9)) if vol_after_40>0 else 0
            else:
                n_20 = 0
                if vol_after_40>0: continue
            total_cap = n_hq*cap.get("40HQ",0)+n_40*cap.get("40FT",0)+n_20*cap.get("20FT",0)
            if total_cap < volume_cbm - 1e-9: continue
            cost = n_hq*rates.get("40HQ",0.0)+n_40*rates.get("40FT",0.0)+n_20*rates.get("20FT",0.0)
            if (cost < best_cost - 1e-6) or (abs(cost-best_cost)<=1e-6 and (n_hq+n_40+n_20) < (best["40HQ"]+best["40FT"]+best["20FT"])):
                best_cost = cost; best = {"20FT":n_20,"40FT":n_40,"40HQ":n_hq}
    if math.isinf(best_cost): raise LookupError("No container combination covers this volume.")
    return best


# =================== AIR lanes ===================
AIR_SHEET       = ""
AIR_LANE_COLS = {
    "oport": ["Origin airport","Origin","POL","Airport","Origin port"],
    "ocode": ["IATA ORIGIN","IATA ORG","IATA_POL","Origin code","IATA ORG CODE","IATA","ORG","OPORT CODE","Origin port code"],
    "dport": ["Destination airport","Destination","POD","Airport","Destination port"],
    "dcode": ["IATA DEST","IATA DST","IATA_POD","Destination code","IATA DST CODE","DSTA","DST","DPORT CODE","Destination port code"],
    "r100":  ["Rate_100","100kg","+100",">=100"],
    "r300":  ["Rate_300","300kg","+300",">=300"],
    "r500":  ["Rate_500","500kg","+500",">=500"],
    "r1000": ["Rate_1000","1000kg","+1000",">=1000"],
    "r1200": ["Rate_1200","+1200",">=1200"],
    "r1500": ["Rate_1500","+1500",">=1500"],
    "r2000": ["Rate_2000","+2000",">=2000","2000+"],
}

def _detect_cols_any(df: pd.DataFrame, candidates: Dict[str, List[str]]) -> Dict[str, str]:
    names = { _norm(c): c for c in df.columns }
    out = {}
    for key, opts in candidates.items():
        for o in opts:
            if _norm(o) in names:
                out[key] = names[_norm(o)]; break
        if key not in out:
            for cn, cr in names.items():
                if any(_norm(o) in cn or cn in _norm(o) for o in opts):
                    out[key] = cr; break
    return out

def _is_iata(s: str) -> bool:
    return isinstance(s, str) and s.strip().isalpha() and len(s.strip())==3

def air_auto_sheet(excel_path: str) -> str:
    xls = pd.ExcelFile(excel_path)
    best = None
    for name in xls.sheet_names:
        for hdr in range(0,6):
            try:
                df_try = pd.read_excel(excel_path, sheet_name=name, header=hdr)
                cols = _detect_cols_any(df_try, AIR_LANE_COLS); score = len(cols)
            except Exception:
                continue
            pref = 1 if any(k in name.upper() for k in ["AIR","AIRFREIGHT","AIR FREIGHT","AIR_LANES"]) else 0
            key = (score, pref, -hdr)
            if (best is None) or (key > best[0]): best = (key, name, hdr)
    if not best: raise ValueError("Could not auto-detect the AIR lanes sheet.")
    return best[1]

def lees_air_lanes(pad: str, sheet_name: str):
    last_err = None; df=None; cols=None
    for hdr in range(0,6):
        try:
            df_try = pd.read_excel(pad, sheet_name=sheet_name, header=hdr)
            cols_try = _detect_cols_any(df_try, AIR_LANE_COLS)
            df, cols = df_try, cols_try; break
        except Exception as e:
            last_err = e; continue
    if df is None: raise last_err if last_err else ValueError("Could not read AIR lanes sheet.")
    # normalize numbers
    def num(x):
        if isinstance(x,str): x = x.replace(".","").replace(",",".")
        return pd.to_numeric(x, errors="coerce")
    for key in ["r100","r300","r500","r1000","r1200","r1500","r2000"]:
        c = cols.get(key)
        if c and c in df.columns: df[c] = num(df[c])
    return df, cols

def resolve_air_input(inp: str, df: pd.DataFrame, cols: Dict[str,str]) -> str:
    s = (inp or "").strip()
    if _is_iata(s): return s.upper()
    oc, ocode = cols.get("oport"), cols.get("ocode")
    dc, dcode = cols.get("dport"), cols.get("dcode")
    name2code = {}
    if oc and ocode:
        for _, r in df[[oc,ocode]].dropna().iterrows():
            name2code[str(r[oc]).strip()] = str(r[ocode]).strip().upper()
    if dc and dcode:
        for _, r in df[[dc,dcode]].dropna().iterrows():
            name2code[str(r[dc]).strip()] = str(r[dcode]).strip().upper()
    if s in name2code: return name2code[s]
    for k,v in name2code.items():
        if k.lower()==s.lower(): return v
    raise ValueError(f"Could not resolve IATA for '{inp}'.")

def air_rates_for_lane(df: pd.DataFrame, cols: Dict[str,str], o_code: str, d_code: str) -> Dict[str,float]:
    oc, dc = cols.get("ocode"), cols.get("dcode")
    if not oc or not dc: raise ValueError("AIR lanes sheet lacks IATA code columns.")
    o = o_code.strip().upper(); d = d_code.strip().upper()
    sub = df[(df[oc].astype(str).str.upper()==o) & (df[dc].astype(str).str.upper()==d)]
    if sub.empty:
        sub = df[(df[oc].astype(str).str.upper()==d) & (df[dc].astype(str).str.upper()==o)]
    if sub.empty: raise LookupError(f"No AIR lane for {o} → {d}.")
    r = sub.iloc[0]; rates = {}
    for key in ["r100","r300","r500","r1000","r1200","r1500","r2000"]:
        c = cols.get(key)
        if c and pd.notna(r.get(c)): rates[key] = float(r[c])
    if not rates: raise LookupError("No AIR bracket rates filled for this lane.")
    return rates

def pick_air_rate(rates: Dict[str,float], acw_kg: float):
    BRK = [("r2000",2000),("r1500",1500),("r1200",1200),("r1000",1000),("r500",500),("r300",300),("r100",100)]
    for key, minw in BRK:
        if acw_kg >= minw and key in rates: return key, rates[key]
    for key,_ in reversed(BRK):
        if key in rates: return key, rates[key]
    raise LookupError("No applicable AIR rate found.")

def find_lcl_rate_per_cbm(df_services: pd.DataFrame, pol_code: str, pod_code: str) -> float:
    """
    Find LCL rate per gross cbm in Services sheet:
      - Type = 'Freight', Mode = 'LCL'
      - PORT CODE equals 'POL-POD' (UN/LOCODE), order-insensitive
    Returns float rate. Raises LookupError with clear message if not found.
    """
    tcol = COLS["type"]; mcol = COLS["mode"]; pcol = COLS["port"]; rcol = COLS["rate_per_cbm"]
    for c in (tcol, mcol, pcol, rcol):
        if c not in df_services.columns:
            raise LookupError("Services sheet mist vereiste kolommen voor LCL (Type/Mode/PORT CODE/Rate per cbm).")

    pair1 = f"{pol_code}-{pod_code}".strip().upper()
    pair2 = f"{pod_code}-{pol_code}".strip().upper()

    m_type = df_services[tcol].astype(str).str.strip().str.casefold() == "freight"
    m_mode = df_services[mcol].astype(str).str.strip().str.casefold() == "lcl"
    pser   = df_services[pcol].astype(str).str.strip().str.upper()
    m_port = (pser == pair1) | (pser == pair2)

    sub = df_services.loc[m_type & m_mode & m_port].copy()
    if sub.empty:
        raise LookupError(f"LCL tarief niet gevonden voor lane {pair1} in tarieven sheet (PORT CODE).")

    import pandas as _pd
    rate = _pd.to_numeric(sub.iloc[0][rcol], errors="coerce")
    if _pd.isna(rate):
        raise LookupError(f"LCL tarief gevonden voor {pair1}, maar waarde is niet numeriek.")
    return float(rate)



import unicodedata, difflib

def _strip_accents(s: str) -> str:
    try:
        return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    except Exception:
        return s

_COUNTRY_SYNONYMS = {
    "the netherlands": "netherlands",
    "nederland": "netherlands",
    "holland": "netherlands",
    "deutschland": "germany",
    "españa": "spain",
    "espana": "spain",
    "éire": "ireland",
    "czech republic": "czechia",
    "u.s.a.": "united states",
    "usa": "united states",
    "u.s.": "united states",
    "us": "united states",
    "u.k.": "united kingdom",
    "uk": "united kingdom",
    "r.o.i.": "ireland",
}

def _canon_country(s: str) -> str:
    if not s: return ""
    s0 = _strip_accents(str(s)).casefold().strip()
    s0 = s0.replace(",", " ").replace(".", " ").replace("  ", " ")
    s0 = s0.replace("the ", " ")
    s0 = s0.strip()
    return _COUNTRY_SYNONYMS.get(s0, s0)

def _detect_cols_any(df: pd.DataFrame, candidates: Dict[str, List[str]]) -> Dict[str, str]:
    out = {}
    names = { _norm(c): c for c in df.columns }
    for key, opts in candidates.items():
        for o in opts:
            if _norm(o) in names:
                out[key] = names[_norm(o)]
                break
        if key not in out:
            for cn, cr in names.items():
                if any(_norm(o) in cn or cn in _norm(o) for o in opts):
                    out[key] = cr
                    break
    return out

def lees_dest_only_charges(pad: str, sheet_name: str = DEST_ONLY_SHEET):
    xls = pd.ExcelFile(pad)
    if sheet_name not in xls.sheet_names:
        raise ValueError(f"Sheet '{sheet_name}' niet gevonden. Maak een tab '{sheet_name}' met de voorgestelde kolommen.")
    df = pd.read_excel(pad, sheet_name=sheet_name, header=0)
    cols = _detect_cols_any(df, DEST_ONLY_COLS)
    if "country" not in cols or "mode" not in cols:
        raise ValueError(f"Kolommen niet gevonden in '{sheet_name}'. Minimaal nodig: Country en Mode.")
    for k in ("rate_20","rate_40","rate_40hq","rate_lcl","rate_air"):
        c = cols.get(k)
        if c and c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df[cols["country"]] = df[cols["country"]].astype(str).str.strip()
    df[cols["mode"]] = df[cols["mode"]].astype(str).str.strip().str.upper()
        # precompute normalized country for robust matching
    # precompute normalized country for robust matching
    df["_N_COUNTRY"] = df[cols["country"]].astype(str).apply(_canon_country)
    return df, cols

def find_dest_only_rate(df: pd.DataFrame, cols: Dict[str,str], country: str, mode: str, container: str|None, gross_cbm: float|None, air_kg: float|None):
    c_country = cols["country"]; c_mode = cols["mode"]
    norm = _canon_country(country)
    sub = df[(df[c_mode] == mode.strip().upper()) & (df["_N_COUNTRY"] == norm)]
    if sub.empty:
        # partial contains either direction
        sub = df[(df[c_mode] == mode.strip().upper()) & (
            df[c_country].str.contains(country, case=False, na=False) | (norm != ""))]
    if sub.empty and norm:
        # fuzzy match as last resort
        cand = df[df[c_mode] == mode.strip().upper()]["_N_COUNTRY"].tolist()
        match = difflib.get_close_matches(norm, cand, n=1, cutoff=0.75)
        if match:
            sub = df[(df[c_mode] == mode.strip().upper()) & (df["_N_COUNTRY"] == match[0])]
    if sub.empty:
        raise LookupError(f"Geen DestOnlyCharges tarief gevonden voor {country} / {mode}.")
    r = sub.iloc[0]
    charge_name = r.get(cols.get("charge_name"), None)
    if not charge_name:
        charge_name = {"FCL":"DTHC","LCL":"NVOCC charges","AIR":"ATHC"}.get(mode, "Destination charges")
    if mode == "FCL":
        if not container: container = "20FT"
        key = {"20FT":"rate_20","40FT":"rate_40","40HQ":"rate_40hq"}.get(container)
        c = cols.get(key)
        if not c or pd.isna(r.get(c)):
            raise LookupError(f"Tarief ontbreekt voor {mode} / {country} / container {container}.")
        rate = float(r[c]); qty_val, qty_unit = 1.0, f"{container}"
        amount = rate
        return (charge_name, qty_val, qty_unit, eur(rate), amount)
    if mode == "LCL":
        c = cols.get("rate_lcl")
        if not c or pd.isna(r.get(c)):
            raise LookupError(f"Tarief ontbreekt voor {mode} / {country} (per cbm gross).")
        rate = float(r[c])
        if gross_cbm is None: gross_cbm = 0.0
        amount = round(rate * gross_cbm, 2)
        return (charge_name, gross_cbm, "cbm gross", eur(rate) + " / cbm gross", amount)
    if mode == "AIR":
        c = cols.get("rate_air")
        if not c or pd.isna(r.get(c)):
            raise LookupError(f"Tarief ontbreekt voor {mode} / {country} (per kg).")
        rate = float(r[c])
        if air_kg is None: air_kg = 0.0
        amount = round(rate * air_kg, 2)
        return (charge_name, air_kg, "kg (charg.)", eur(rate) + " / kg", amount)
    raise LookupError("ROAD heeft geen DestOnlyCharges of onbekende mode.")

# =================== PDF ===================
def _find_brand_logo_file(brand_name: str) -> str:
    """Zoek VOERMAN/TRANSPACK logo (png/jpg) in de huidige map of scriptmap."""
    import os
    base_names = [brand_name.upper(), brand_name.capitalize(), brand_name.lower()]
    exts = ["png","PNG","jpg","JPG","jpeg","JPEG"]
    candidates = []
    for b in base_names:
        for e in exts:
            candidates.append(f"{b}.{e}")
    search_dirs = [os.getcwd(), os.path.dirname(__file__)]
    for d in search_dirs:
        for fn in candidates:
            p = os.path.join(d, fn)
            if os.path.isfile(p):
                return p
    return ""

    def _val(v):
        try:
            v = v.get()
        except Exception:
            pass
        return '' if v is None else str(v)

    # --- Compute VAT gate early (to hide mid Total row correctly) ---
    try:
        oc = _val(getattr(self, 'var_origin_country', None)).strip().upper()
        dc = _val(getattr(self, 'var_dest_country', None)).strip().upper()
    except Exception:
        oc, dc = '', ''
    EU_ISO = {'AT','BE','BG','HR','CY','CZ','DK','EE','FI','FR','DE','GR','HU','IE','IT','LV','LT','LU','MT','NL','PL','PT','RO','SK','SI','ES','SE'}
    intra_eu = (oc in EU_ISO) and (dc in EU_ISO)
    client_private = False
    for _name in ('var_client_private','var_client_is_private','var_is_private','var_client_type'):
        _v = getattr(self, _name, None)
        if _v is None:
            continue
        _s = _val(_v).strip().lower()
        if _s in ('1','true','yes','private','privé'):
            client_private = True
            break
        if _s in ('0','false','no','agent','agent/partner','partner','business'):
            client_private = False
            break
    vat_gate = bool(client_private and intra_eu)


    LEFT_M = 18*mm; RIGHT_M = 18*mm
    PAGE_W = A4[0]; content_w = PAGE_W - LEFT_M - RIGHT_M

    doc = SimpleDocTemplate(save_path, pagesize=A4, leftMargin=LEFT_M, rightMargin=RIGHT_M, topMargin=16*mm, bottomMargin=16*mm)
    story = []

    if logo_path and os.path.exists(logo_path):
        img = _logo_flowable(logo_path, max_w_mm=60, max_h_mm=24)
    else:
        img = None
    # apply brand overrides or fall back to globals

    _name = company_name or BEDRIJFSNAAM

    _addr = company_addr or BEDRIJF_ADRES

    _email = company_email or BEDRIJF_EMAIL

    _tel = company_tel or BEDRIJF_TEL

    company_block = Paragraph(f"<b>{escape(_name)}</b><br/>{escape(_addr)}<br/>{escape(_email)} · {escape(_tel)}", style_small)

    left_stack = []
    if img is not None: left_stack.append([img])
    left_stack.append([company_block])
    left_tbl = Table(left_stack, colWidths=[70*mm])
    left_tbl.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))

    title_tbl = Table([[Paragraph("<b>Cost Estimate</b>", style_title)]], colWidths=[content_w - 70*mm - 10*mm])
    title_tbl.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"RIGHT")]))

    header = Table([[left_tbl, title_tbl]], colWidths=[70*mm, content_w - 70*mm])
    header.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(header); story.append(Spacer(1,6))

    details_right = Table([
        ["SO#", _val(so_number)],
        ["Date", datetime.now().strftime("%Y-%m-%d")],
        ["Debtor number", _val(debtor_number)],
        ["Debtor VAT number", _val(debtor_vat)],
        ["Payment term", _val(payment_term)],
        ["VAT memo", _val(vat_memo)],
    ], colWidths=[40*mm, 45*mm])
    details_right.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.25,colors.lightgrey),
        ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("ALIGN",(0,0),(-1,-1),"LEFT"),
    ]))

    left_par = Paragraph("<b>Account / Partner</b><br/>" + escape(_val(client_name)).replace("\n", "<br/>"), style_norm)
    top_tbl = Table([[left_par, details_right]], colWidths=[content_w - 90*mm, 85*mm])
    top_tbl.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(top_tbl); story.append(Spacer(1,10))

    job_tbl = Table([
        ["Job mode", _val(job_mode), "Volume", f"{float(volume_cbm):.2f} m³"],
        ["Origin", _val(origin_addr) if origin_addr else "-", "Destination", _val(dest_addr) if dest_addr else "-"],
    ], colWidths=[80*mm, 30*mm, 34*mm, 30*mm])
    job_tbl.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.25,colors.lightgrey),
        ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),6),
        ("RIGHTPADDING",(0,0),(-1,-1),6),
    ]))
    job_tbl.hAlign = "LEFT"
    story.append(job_tbl); story.append(Spacer(1,8))

        # Table header & widths depending on show_rates
    if show_rates:
        col_widths = [0.55*content_w, 0.15*content_w, 0.15*content_w, 0.15*content_w]
        table_data = [["Charge", "Quantity", "Estimate Rate", "Amount Total"]]
    else:
        col_widths = [0.62*content_w, 0.15*content_w, 0.23*content_w]
        table_data = [["Charge", "Quantity", "Amount Total"]]
    total_sum = 0.0
    for r in charges_rows:
        descr = r["descr"]; qty = r.get("qty"); rate = r.get("rate"); amt = r.get("amount"); skip = r.get("skip_total", False)
        if show_rates:
            table_data.append([
                descr,
                qty if isinstance(qty, str) else (fmt_qty(qty) if (r.get("qty_unit") is None) else fmt_qty(qty, r.get("qty_unit"))),
                "-" if unit_rate is None else unit_rate,
                "-" if amt  is None else eur(float(amt)),
            ])
        else:
            table_data.append([
                descr,
                qty if isinstance(qty, str) else (fmt_qty(qty) if (r.get("qty_unit") is None) else fmt_qty(qty, r.get("qty_unit"))),
                "-" if amt  is None else eur(float(amt)),
            ])
        if (amt is not None) and not skip:
            total_sum += float(amt)
    if show_rates:
        # Skip intermediate total when VAT is printed
        if not (bool(show_vat) and bool(vat_gate)):
            table_data.append(["", "", Paragraph("<b>Total</b>", style_norm), Paragraph(f"<b>{eur(total_sum)}</b>", style_norm)])
    else:
        if not (bool(show_vat) and bool(vat_gate)):
            table_data.append(["", Paragraph("<b>Total</b>", style_norm), Paragraph(f"<b>{eur(total_sum)}</b>", style_norm)])
    charges_tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    charges_tbl.hAlign = 'LEFT'
    ts = [
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-2),0.25,colors.lightgrey),
        ("BOX",(0,0),(-1,-2),0.25,colors.grey),
    ]
    if show_rates:
        ts += [("ALIGN",(1,1),(1,-2),"RIGHT"), ("ALIGN",(2,1),(2,-2),"RIGHT"), ("ALIGN",(3,1),(3,-1),"RIGHT"), ("RIGHTPADDING",(3,1),(3,-1),6)]
    else:
        ts += [("ALIGN",(1,1),(1,-2),"RIGHT"), ("ALIGN",(2,1),(2,-1),"RIGHT"), ("RIGHTPADDING",(2,1),(2,-1),6)]
    ts += [("LEFTPADDING",(0,0),(-1,-1),6), ("RIGHTPADDING",(0,0),(-1,-1),6)]
    charges_tbl.setStyle(TableStyle(ts))
    story.append(charges_tbl); story.append(Spacer(1,10))

    
    # --- Destination-only charges ---
    if dest_only_rows:
        header = ["Charge (Destination-only)", "Quantity", "Rate"] if show_rates else ["Charge (Destination-only)", "Quantity", "Amount Total"]
        t2 = [header]
        total2 = 0.0
        for r in dest_only_rows:
            descr = r.get("descr", "-")
            qty   = r.get("qty", "-")
            rate  = r.get("rate", "-")
            amount = r.get("amount", 0.0)
            third = rate if show_rates else (eur(float(amount)) if isinstance(amount, (int, float,)) else str(amount))
            t2.append([descr, qty, third])
            try:
                total2 += float(amount or 0.0)
            except Exception:
                pass
    
        col_widths2 = [content_w*0.56, content_w*0.18, content_w*0.26] if show_rates else [content_w*0.62, content_w*0.18, content_w*0.20]
        dest_tbl = Table(t2, colWidths=col_widths2, repeatRows=1)
        dest_tbl.hAlign = 'LEFT'
        ts2 = [
            ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("GRID",(0,0),(-1,-1),0.25,colors.lightgrey),
            ("BOX",(0,0),(-1,-1),0.25,colors.grey),
            ("LEFTPADDING",(0,0),(-1,-1),6), ("RIGHTPADDING",(0,0),(-1,-1),6),
        ]
        ts2 += [("ALIGN",(2,1),(2,-1),"RIGHT"), ("RIGHTPADDING",(2,1),(2,-1),6)]
        dest_tbl.setStyle(TableStyle(ts2))
        story.append(dest_tbl); story.append(Spacer(1,10))
    
    # Ensure default notes exist
    valid_txt = f"Rates valid for {payment_term} from issue date."
    payment_txt = vat_memo
    # Default bullet notes (will be replaced if VAT block sets a custom list later)
    ul_items = [
        "Transit times and rates are estimates only and may vary due to carrier/supplier changes.",
        valid_txt,
        payment_txt,
    ]

    

        # Subtotal used for VAT calculation
    try:
        subtotal = float(total_sum)
    except Exception:
        subtotal = 0.0
    try:
        subtotal += float(total2)
    except Exception:
        pass



    # --- VAT breakdown (if applicable) ---
    show_vat = bool(show_vat)
    vat_rate_val = _num(vat_rate, 21.0) if vat_rate is not None else 21.0
    vat_applies = bool(vat_applies) if vat_applies is not None else False


    # Safety gate: only show VAT for PRIVATE + intra‑EU (independent of any earlier flag)
    def _val(v):
        try:
            return v.get()
        except Exception:
            return v
    oc = _val(getattr(self, 'var_origin_country', None))
    dc = _val(getattr(self, 'var_dest_country', None))
    orig_iso = (str(oc).strip().upper() if oc else "")
    dest_iso = (str(dc).strip().upper() if dc else "")
    EU_ISO = {"AT","BE","BG","HR","CY","CZ","DK","EE","FI","FR","DE","GR","HU","IE","IT","LV","LT","LU","MT","NL","PL","PT","RO","SK","SI","ES","SE"}
    intra_eu = (orig_iso in EU_ISO) and (dest_iso in EU_ISO)

    client_private = False
    for _name in ("var_client_private","var_client_is_private","var_is_private","var_client_type"):
        _v = getattr(self, _name, None)
        if _v is None:
            continue
        _s = str(_val(_v)).strip().lower()
        if _s in ("1","true","yes","private","privé"):
            client_private = True
            break
        if _s in ("0","false","no","agent","agent/partner","partner","business"):
            client_private = False
            break

    vat_gate = bool(client_private and intra_eu)
    if show_vat and vat_gate and subtotal > 0:
        vat_amount = round(subtotal * (vat_rate_val / 100.0), 2)
        total_inc = round(subtotal + vat_amount, 2)

        # Compact, polished VAT block
        vat_data = [
            ["Subtotal", Paragraph(eur(subtotal), style_norm)],
            [f"VAT ({vat_rate_val:.0f}%)", Paragraph(eur(vat_amount), style_norm)],
            [Paragraph("<b>Total incl. VAT</b>", style_norm),
             Paragraph(f"<b>{eur(total_inc)}</b>", style_norm)],
        ]

        amt_col_w = (0.15 if show_rates else 0.23) * content_w
        vat_tbl = Table(vat_data, colWidths=[content_w - amt_col_w, amt_col_w])
        vat_tbl.setStyle(TableStyle([
            ("ALIGN",        (1, 0), (1, -1), "RIGHT"),
            ("LEFTPADDING",  (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING",   (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
            ("LINEABOVE",    (0, 0), (-1, 0), 0.25, colors.HexColor("#BFBFBF")),
            ("LINEABOVE",    (0, 2), (-1, 2), 0.75, colors.black),
            ("BACKGROUND",   (0, 2), (-1, 2), colors.HexColor("#F2F2F2")),
            ("FONTNAME",     (0, 2), (-1, 2), "Helvetica-Bold"),
        ]))

        story.append(Spacer(1, 8))
        story.append(vat_tbl)

        # Clear and concise footer wording when VAT printed
        payment_txt = "Prices incl. VAT."

    # Build bullet notes (always), based on the final payment_txt
    ul_items = [
        "Transit times and rates are estimates only and may vary due to carrier/supplier changes.",
        valid_txt,
        payment_txt,
    ]
    if dest_only_rows:
        ul_items.append("DTHC, ATHC, NVOCC charges will be billed at cost after the invoice of the forwarder + a prepayment fee of 25 Euro.")
    for b in ul_items:
        story.append(Paragraph("• " + escape(b), style_small2))
    
    doc.build(story)

# =================== GUI ===================
class App(tk.Tk):
    def fail(self, msg: str, exc: Exception | None = None, title: str = "Error"):
        """Uniform error handler: set status, log, and show messagebox."""
        try:
            if exc:
                import traceback, os
                log(f"ERROR: {msg}: {exc}")
                log(traceback.format_exc())
            else:
                log(f"ERROR: {msg}")
        except Exception:
            pass
        try:
            self.set_status(msg)
        except Exception:
            pass
        try:
            from tkinter import messagebox
            messagebox.showerror(title, msg if not exc else f"{msg}\n\n{exc}")
        except Exception:
            pass

    def __init__(self):
        

        # --- Quote bar (multi-quote) ---
        # Separate row above 'Select services'

        super().__init__()
        self.title("Voerman Quote Studio — MQ25")
        # Ensure output directory exists
        try:
            os.makedirs(OUTPUT_DIR, exist_ok=True)
        except Exception as _e:
            print('Warning: cannot create OUTPUT_DIR', OUTPUT_DIR, _e)
        self.minsize(1120, 980)
        self.geometry("1120x980")
        self.resizable(True, True)
        # Theming: capture classic ttk theme and optionally init ttkbootstrap style
        self._classic_theme = ttk.Style(self).theme_use()
        self._tb_style = None
        if BOOTSTRAP_AVAILABLE:
            try:
                self._tb_style = tb.Style(theme="flatly")
            except Exception:
                pass
# --- VAT / countries state ---
        self.var_origin_country = tk.StringVar(value="NL")
        self.var_dest_country   = tk.StringVar(value="NL")
        # Visible name vars mirror the code vars (for UI only)
        self.var_origin_country_name = tk.StringVar(value=code_to_name(self.var_origin_country.get()))
        self.var_dest_country_name   = tk.StringVar(value=code_to_name(self.var_dest_country.get()))
        # Keep them in sync if backend changes the code vars
        try:
            self.var_origin_country.trace_add("write", lambda *_: self.var_origin_country_name.set(code_to_name(self.var_origin_country.get())))
            self.var_dest_country.trace_add("write",  lambda *_: self.var_dest_country_name.set(code_to_name(self.var_dest_country.get())))
        except Exception:
            self.var_origin_country.trace("w", lambda *_: self.var_origin_country_name.set(code_to_name(self.var_origin_country.get())))
            self.var_dest_country.trace("w",  lambda *_: self.var_dest_country_name.set(code_to_name(self.var_dest_country.get())))
    
        self.var_vat_rate       = tk.DoubleVar(value=21.0)
        self.var_pdf_show_vat   = tk.BooleanVar(value=False)
        self.var_vat_status     = tk.StringVar(value="VAT: Not applicable")
        self.vat_applies        = False
        self.vat_memo           = "VAT not applicable (export/import / one leg outside EU)."

        # Topbar with theme switcher
        topbar = ttk.Frame(self, padding=6)
        topbar.pack(fill="x", side="top")
        ttk.Label(topbar, text="Thema:").pack(side="left", padx=(0,6))
        self._theme_var = tk.StringVar(value=("flatly" if self._tb_style else "classic"))
        theme_values = ["classic","flatly","cosmo","minty","darkly","cyborg","voerman","sv-ttk-light","sv-ttk-dark","voerman_clam"]
        self._cmb_theme = ttk.Combobox(topbar, textvariable=self._theme_var, values=theme_values, width=14, state="readonly")
        self._cmb_theme.pack(side="left")
        ttk.Button(topbar, text="Toepassen", command=self.apply_theme).pack(side="left", padx=6)
        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=(2,0))

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0)
        vbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(1, weight=0)
        outer = ttk.Frame(canvas, padding=12)
        outer_id = canvas.create_window((0,0), window=outer, anchor="nw")
        def _on_frame_configure(event=None):
            try:
                canvas.configure(scrollregion=canvas.bbox("all"))
            except Exception:
                pass
        outer.bind("<Configure>", _on_frame_configure)
        def _on_canvas_configure(event):
            try:
                canvas.itemconfigure(outer_id, width=event.width)
            except Exception:
                pass
        canvas.bind("<Configure>", _on_canvas_configure)
        # Mousewheel scrolling (Windows/Mac/Linux)
        def _on_mousewheel(event):
            delta = 0
            if hasattr(event, 'delta') and event.delta:
                delta = int(-1 * (event.delta/120))
            elif getattr(event, 'num', None) in (4, 5):
                delta = -1 if event.num == 4 else 1
            if delta:
                canvas.yview_scroll(delta, "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", _on_mousewheel)
        canvas.bind_all("<Button-5>", _on_mousewheel)
        for c in range(2): outer.grid_columnconfigure(c, weight=1, uniform="col")

        # --- Quote bar (multi-quote) ---
        # Separate row above 'Select services'
        self.quote_bar = ttk.Frame(outer)
        self.quote_bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,6))
        # Pre-render a visible Quote 1 / + Add quote so the bar is never empty
        try:
            _btn_q1 = ttk.Button(self.quote_bar, text="Quote 1")
            try: _btn_q1.state(["disabled"])
            except Exception: _btn_q1.configure(state="disabled")
            _btn_q1.pack(side="left", padx=(0,6))
            ttk.Button(self.quote_bar, text="+ Add quote", command=lambda: getattr(self, "_add_quote", lambda: None)()).pack(side="left", padx=(6,0))
        except Exception:
            pass
        sec_services = ttk.LabelFrame(outer, text="Select services")
        sec_services.grid(row=1, column=0, columnspan=2, sticky="ew", padx=0, pady=(0,8))
        for i in range(3): sec_services.grid_columnconfigure(i, weight=1)
        self.var_origin = tk.BooleanVar(value=True)
        self.var_freight = tk.BooleanVar(value=True)
        self.var_dest = tk.BooleanVar(value=True)
        ttk.Checkbutton(sec_services, text="Origin", variable=self.var_origin, command=self._toggle_inputs).grid(row=0, column=0, padx=8, pady=4, sticky="w")
        ttk.Checkbutton(sec_services, text="Freight", variable=self.var_freight, command=self._toggle_inputs).grid(row=0, column=1, padx=8, pady=4, sticky="w")
        ttk.Checkbutton(sec_services, text="Destination", variable=self.var_dest, command=self._toggle_inputs).grid(row=0, column=2, padx=8, pady=4, sticky="w")

        
        sec_ship = ttk.LabelFrame(outer, text="Shipment")
        sec_ship.grid(row=2, column=0, sticky="nsew", padx=(0,8), pady=8)
        for c in range(4):
            sec_ship.grid_columnconfigure(c, weight=1 if c in (1,3) else 0)

        # Row 0: Origin location + Origin country
        ttk.Label(sec_ship, text="Origin location:").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        self.ent_origin = ttk.Entry(sec_ship, width=36)
        self.ent_origin.grid(row=0, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(sec_ship, text="Origin country:").grid(row=0, column=2, sticky="e", padx=6, pady=4)
        self.cmb_ship_o_country = ttk.Combobox(sec_ship, width=22, state='readonly', values=COUNTRY_NAMES, textvariable=self.var_origin_country_name)
        self.cmb_ship_o_country.grid(row=0, column=3, sticky="w", padx=6, pady=4)

        # Row 1: Destination location + Destination country
        ttk.Label(sec_ship, text="Destination location:").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        self.ent_dest = ttk.Entry(sec_ship, width=36)
        self.ent_dest.grid(row=1, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(sec_ship, text="Destination country:").grid(row=1, column=2, sticky="e", padx=6, pady=4)
        self.cmb_ship_d_country = ttk.Combobox(sec_ship, width=22, state='readonly', values=COUNTRY_NAMES, textvariable=self.var_dest_country_name)
        self.cmb_ship_d_country.grid(row=1, column=3, sticky="w", padx=6, pady=4)

        
        # When the user selects a country name, update the hidden ISO2 vars
        def _on_origin_country_changed(event=None):
            try:
                self.var_origin_country.set(name_to_code(self.var_origin_country_name.get()))
            except Exception:
                pass
        def _on_dest_country_changed(event=None):
            try:
                self.var_dest_country.set(name_to_code(self.var_dest_country_name.get()))
            except Exception:
                pass
        self.cmb_ship_o_country.bind("<<ComboboxSelected>>", _on_origin_country_changed)
        self.cmb_ship_d_country.bind("<<ComboboxSelected>>", _on_dest_country_changed)
    # Row 2: Mode + Volume
        ttk.Label(sec_ship, text="Mode:").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        self.cmb_mode = ttk.Combobox(sec_ship, values=VALID_MODES, state="readonly", width=14)
        self.cmb_mode.set("FCL")
        self.cmb_mode.grid(row=2, column=1, sticky="w", padx=6, pady=4)
        self.cmb_mode.bind("<<ComboboxSelected>>", lambda e: self._toggle_inputs())

        ttk.Label(sec_ship, text="Volume (m³):").grid(row=2, column=2, sticky="e", padx=6, pady=4)
        self.ent_volume = ttk.Entry(sec_ship, width=14)
        self.ent_volume.grid(row=2, column=3, sticky="w", padx=6, pady=4)

        # --- FCL panel ---
        self.sec_fcl = ttk.LabelFrame(outer, text="Freight – lanes")
        self.sec_fcl.grid(row=2, column=1, sticky="nsew", padx=(8,0), pady=8)
        for c in range(4): self.sec_fcl.grid_columnconfigure(c, weight=1 if c in (1,3) else 0)
        self.lbl_pol = ttk.Label(self.sec_fcl, text="FCL Origin port (code or name):"); self.lbl_pol.grid(row=0, column=0, sticky="e", padx=6, pady=4)
        self.ent_pol = ttk.Entry(self.sec_fcl, width=30); self.ent_pol.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        self.lbl_pod = ttk.Label(self.sec_fcl, text="FCL Destination port (code or name):"); self.lbl_pod.grid(row=0, column=2, sticky="e", padx=6, pady=4)
        self.ent_pod = ttk.Entry(self.sec_fcl, width=30); self.ent_pod.grid(row=0, column=3, sticky="ew", padx=6, pady=4)
        self.lbl_container = ttk.Label(self.sec_fcl, text="Container (FCL):"); self.lbl_container.grid(row=1, column=0, sticky="e", padx=6, pady=4)
        self.cmb_fcl_choice = ttk.Combobox(self.sec_fcl, values=["Auto (best price)","20FT","40FT","40HQ"], state="readonly", width=18)
        self.cmb_fcl_choice.set("Auto (best price)")
        self.cmb_fcl_choice.grid(row=1, column=1, sticky="w", padx=6, pady=4)

        # --- ROAD panel ---
        self.sec_road = ttk.LabelFrame(outer, text="Freight – ROAD options")
        self.sec_road.grid(row=2, column=1, sticky="nsew", padx=(8,0), pady=8)
        for c in range(4): self.sec_road.grid_columnconfigure(c, weight=1 if c in (1,3) else 0)
        ttk.Label(self.sec_road, text="Road service type:").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        self.cmb_road_type = ttk.Combobox(self.sec_road, values=["Combined","Direct"], state="readonly", width=18)
        self.cmb_road_type.set("Combined"); self.cmb_road_type.grid(row=0, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(self.sec_road, text="Rate (€/km):").grid(row=0, column=2, sticky="e", padx=6, pady=4)
        self.ent_road_rate = ttk.Entry(self.sec_road, width=12); self.ent_road_rate.insert(0, f"{DEFAULT_ROAD_RATE_COMBINED:.2f}")
        self.ent_road_rate.grid(row=0, column=3, sticky="w", padx=6, pady=4)

        def on_road_type_change(event=None):
            t = self.cmb_road_type.get().strip()
            default = DEFAULT_ROAD_RATE_COMBINED if t == "Combined" else DEFAULT_ROAD_RATE_DIRECT
            try:
                cur = float(self.ent_road_rate.get().replace(",", "."))
            except Exception:
                cur = None
            if cur in (DEFAULT_ROAD_RATE_COMBINED, DEFAULT_ROAD_RATE_DIRECT) or cur is None:
                self.ent_road_rate.delete(0, tk.END)
                self.ent_road_rate.insert(0, f"{default:.2f}")
        self.cmb_road_type.bind("<<ComboboxSelected>>", on_road_type_change)

        # --- Email + AI ---
        sec_email = ttk.LabelFrame(outer, text="Email intake – paste RFQ email below, then click Parse")
        sec_email.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=0, pady=8)
        sec_email.grid_columnconfigure(0, weight=1)
        self.txt_email = scrolledtext.ScrolledText(sec_email, height=10, wrap="word")
        self.txt_email.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        email_btns = ttk.Frame(sec_email); email_btns.grid(row=1, column=0, sticky="e", padx=6, pady=(0,6))
        ttk.Button(email_btns, text="Parse (rules)", command=self.do_parse_email).grid(row=0, column=0, padx=6)
        ttk.Button(email_btns, text="AI Parse (OpenAI)", command=self.do_ai_parse_email).grid(row=0, column=1, padx=6)
        ttk.Button(email_btns, text="AI Test", command=self.do_ai_test).grid(row=0, column=2, padx=6)
        ttk.Button(email_btns, text="Clear", command=lambda: self.txt_email.delete("1.0", tk.END)).grid(row=0, column=3, padx=6)

        
        # --- Client (collapsible) ---
        client_toggle_bar = ttk.Frame(outer)
        client_toggle_bar.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(0, 0))
        self.var_client_show = tk.BooleanVar(value=False)
        # Replace checkbox with ttkbootstrap Switch when available
        if BOOTSTRAP_AVAILABLE:
            try:
                from ttkbootstrap.widgets import Switch  # type: ignore
                self.client_switch = Switch(
                    client_toggle_bar,
                    text="Client details",
                    variable=self.var_client_show,
                    bootstyle="success",
                    command=self._toggle_client_section
                )
                self.client_switch.pack(side="left", padx=4)
            except Exception:
                ToggleSwitch(
                    client_toggle_bar,
                    text="Client details",
                    variable=self.var_client_show,
                    command=self._toggle_client_section,
                    width=40, height=20
                ).pack(side="left", padx=4)
        else:
            ToggleSwitch(
                    client_toggle_bar,
                    text="Client details",
                    variable=self.var_client_show,
                    command=self._toggle_client_section,
                    width=40, height=20
                ).pack(side="left", padx=4)

        self.sec_client = ttk.LabelFrame(outer, text="Client")
        self.sec_client.grid(row=8, column=0, columnspan=2, sticky="nsew", padx=0, pady=(0, 8))
        for c in range(6):
            self.sec_client.grid_columnconfigure(c, weight=1 if c in (1, 3, 5) else 0)

        ttk.Label(self.sec_client, text="Type:").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        self.var_client_type = tk.StringVar(value="PRIVATE")
        ttk.Radiobutton(self.sec_client, text="Private", value="PRIVATE", variable=self.var_client_type, command=self._toggle_client_ui).grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(self.sec_client, text="Agent/Partner", value="AGENT", variable=self.var_client_type, command=self._toggle_client_ui).grid(row=0, column=2, sticky="w")

        ttk.Label(self.sec_client, text="Saved:").grid(row=0, column=3, sticky="e", padx=6, pady=4)
        self.cmb_saved_client = ttk.Combobox(self.sec_client, state="readonly", width=34, values=[])
        self.cmb_saved_client.grid(row=0, column=4, sticky="ew", padx=6, pady=4)
        ttk.Button(self.sec_client, text="Load", command=self.client_load_selected).grid(row=0, column=5, sticky="w", padx=4)

        ttk.Label(self.sec_client, text="Name (client/agent):").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        self.ent_client_name = ttk.Entry(self.sec_client, width=36)
        self.ent_client_name.grid(row=1, column=1, columnspan=2, sticky="ew", padx=6, pady=4)

        ttk.Label(self.sec_client, text="Assignee (optional):").grid(row=1, column=3, sticky="e", padx=6, pady=4)
        self.ent_assignee = ttk.Entry(self.sec_client, width=30)
        self.ent_assignee.grid(row=1, column=4, columnspan=2, sticky="ew", padx=6, pady=4)

        ttk.Label(self.sec_client, text="Address line 1:").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        self.ent_addr1 = ttk.Entry(self.sec_client, width=36)
        self.ent_addr1.grid(row=2, column=1, columnspan=2, sticky="ew", padx=6, pady=4)

        ttk.Label(self.sec_client, text="Address line 2:").grid(row=2, column=3, sticky="e", padx=6, pady=4)
        self.ent_addr2 = ttk.Entry(self.sec_client, width=30)
        self.ent_addr2.grid(row=2, column=4, columnspan=2, sticky="ew", padx=6, pady=4)

        ttk.Label(self.sec_client, text="City / Postal:").grid(row=3, column=0, sticky="e", padx=6, pady=4)
        self.ent_city_postal = ttk.Entry(self.sec_client, width=36)
        self.ent_city_postal.grid(row=3, column=1, columnspan=2, sticky="ew", padx=6, pady=4)

        ttk.Label(self.sec_client, text="Country:").grid(row=3, column=3, sticky="e", padx=6, pady=4)
        self.ent_country = ttk.Entry(self.sec_client, width=30)
        self.ent_country.grid(row=3, column=4, columnspan=2, sticky="ew", padx=6, pady=4)

        ttk.Label(self.sec_client, text="Email:").grid(row=4, column=0, sticky="e", padx=6, pady=4)
        self.ent_email = ttk.Entry(self.sec_client, width=36)
        self.ent_email.grid(row=4, column=1, columnspan=2, sticky="ew", padx=6, pady=4)

        ttk.Button(self.sec_client, text="Save/Update", command=self.client_save).grid(row=4, column=3, sticky="e", padx=6, pady=4)
        ttk.Button(self.sec_client, text="Delete", command=self.client_delete).grid(row=4, column=4, sticky="w", padx=6, pady=4)

        # initially hidden
        self.sec_client.grid_remove()
        self._refresh_saved_clients()
        self._toggle_client_ui()
        # --- Document & branding ---
        sec_doc = ttk.LabelFrame(outer, text="Document & branding")
        sec_doc.grid(row=9, column=0, columnspan=2, sticky="nsew", padx=0, pady=8)

        # nette, consistente grid: [label,input] x 3 kolommen
        for c in range(6):
            sec_doc.grid_columnconfigure(c, weight=1 if c in (1,3,5) else 0)

        # Rij 0 — Client / Reference / Brand
        ttk.Label(sec_doc, text="Client name (optional):").grid(row=0, column=0, sticky="e", padx=8, pady=6)
        self.ent_klant = ttk.Entry(sec_doc, width=30)
        self.ent_klant.grid(row=0, column=1, sticky="ew", padx=6, pady=6)

        ttk.Label(sec_doc, text="Reference (optional):").grid(row=0, column=2, sticky="e", padx=8, pady=6)
        self.ent_ref = ttk.Entry(sec_doc, width=24)
        self.ent_ref.grid(row=0, column=3, sticky="ew", padx=6, pady=6)

        ttk.Label(sec_doc, text="Brand:").grid(row=0, column=4, sticky="e", padx=8, pady=6)
        try:
            _brand_values = list(BRANDS.keys())
        except Exception:
            _brand_values = ["Voerman","Transpack"]
        self.cmb_brand = ttk.Combobox(sec_doc, state="readonly", width=16, values=_brand_values)
        self.cmb_brand.grid(row=0, column=5, sticky="ew", padx=6, pady=6)

        # Rij 1 — Excel file
        ttk.Label(sec_doc, text="Excel file:").grid(row=1, column=0, sticky="e", padx=8, pady=6)
        self.ent_excel = ttk.Entry(sec_doc)
        try:
            self.ent_excel.insert(0, EXCEL_PAD)
        except Exception:
            pass
        self.ent_excel.grid(row=1, column=1, columnspan=4, sticky="ew", padx=6, pady=6)
        ttk.Button(sec_doc, text="Browse…", command=getattr(self, "kies_excel", lambda: None)).grid(row=1, column=5, sticky="w", padx=6, pady=6)

        # Rij 2 — Logo file
        ttk.Label(sec_doc, text="Logo (optional):").grid(row=2, column=0, sticky="e", padx=8, pady=6)
        self.ent_logo = ttk.Entry(sec_doc)
        self.ent_logo.grid(row=2, column=1, columnspan=4, sticky="ew", padx=6, pady=6)
        ttk.Button(sec_doc, text="Browse…", command=getattr(self, "kies_logo", lambda: None)).grid(row=2, column=5, sticky="w", padx=6, pady=6)

        # Rij 3 — VAT controls
        ttk.Label(sec_doc, text="VAT rate (%):").grid(row=3, column=0, sticky="e", padx=8, pady=6)
        ttk.Entry(sec_doc, textvariable=self.var_vat_rate, width=6).grid(row=3, column=1, sticky="w", padx=6, pady=6)
        ttk.Checkbutton(sec_doc, text="Show VAT line on PDF", variable=self.var_pdf_show_vat).grid(row=3, column=2, columnspan=3, sticky="w", padx=6, pady=6)

        # --- Actions & Status (inside Document & branding) ---
        sec_doc_actions = ttk.Frame(sec_doc)
        sec_doc_actions.grid(row=4, column=0, columnspan=6, sticky="ew", padx=8, pady=(8, 4))
        sec_doc_actions.grid_columnconfigure(0, weight=1)
        sec_doc_actions.grid_columnconfigure(1, weight=0)
        sec_doc_actions.grid_columnconfigure(2, weight=0)
        sec_doc_actions.grid_columnconfigure(3, weight=0)

        # Status (links)
        self.status = tk.StringVar(value="Ready.")
        ttk.Label(sec_doc_actions, textvariable=self.status).grid(row=0, column=0, sticky="w")

        # Select en Generate (rechts)
        self.btn_select = ttk.Button(sec_doc_actions, text="Select quotes…", padding="6 4", command=getattr(self, "select_quotes", lambda: messagebox.showinfo("Select quotes", "Not available in this build.")))
        self.btn_select.grid(row=0, column=1, sticky="e", padx=(0,6))

        self.btn_generate = ttk.Button(sec_doc_actions, text="Generate PDF", padding="6 4", command=self.run)
        self.btn_generate.grid(row=0, column=2, sticky="e", padx=(0,6))

        self.btn_quit = ttk.Button(sec_doc_actions, text="Close", command=self.destroy, padding="6 4")
        self.btn_quit.grid(row=0, column=3, sticky="e")


        # --- Destination-only charges panel (shown when ONLY Destination is selected) ---
        self.sec_destonly = ttk.LabelFrame(outer, text="Destination-only charges")
        self.sec_destonly.grid(row=5, column=0, columnspan=2, sticky="ew", padx=0, pady=(8,0))
        self.sec_destonly.grid_remove()
        for c in range(6): self.sec_destonly.grid_columnconfigure(c, weight=1 if c in (1,3,5) else 0)

        ttk.Label(self.sec_destonly, text="Charge label:").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        self.var_destonly_label = tk.StringVar(value="(auto)")
        ttk.Label(self.sec_destonly, textvariable=self.var_destonly_label, width=18).grid(row=0, column=1, sticky="w", padx=6, pady=4)
        self.lbl_destonly_rate = ttk.Label(self.sec_destonly, text="Amount (€):")
        self.lbl_destonly_rate.grid(row=0, column=2, sticky="e", padx=6, pady=4)
        self.ent_destonly_rate = ttk.Entry(self.sec_destonly, width=12)
        self.ent_destonly_rate.grid(row=0, column=3, sticky="w", padx=6, pady=4)

    # ================= Multi-quote support =================
    def _quotes_init(self):
        # Build menu and first quote from current UI
        try:
            self._quotes = [self._collect_ui_state()]
        except Exception:
            self._quotes = [self._blank_quote_state()]
        self._active_quote = 0
        self._render_quote_bar()
        try:
            self._rebuild_run_menu()
        except Exception:
            pass

    def _render_quote_bar(self):
        # Build the quote buttons + Add
        for w in self.quote_bar.winfo_children():
            try: w.destroy()
            except Exception: pass
        for i in range(len(getattr(self, "_quotes", []))):
            btn = ttk.Button(self.quote_bar, text=f"Quote {i+1}", command=lambda i=i: self._switch_to_quote(i))
            btn.pack(side="left", padx=(0,6))
            if i == getattr(self, "_active_quote", 0):
                try: btn.state(["disabled"])
                except Exception: btn.configure(state="disabled")
        ttk.Button(self.quote_bar, text="+ Add quote", command=self._add_quote).pack(side="left", padx=(6,0))

        # Delete-knop voor de actieve quote
        _del_btn = ttk.Button(self.quote_bar, text="Delete quote", command=self._delete_active_quote)
        if len(getattr(self, "_quotes", [])) <= 1:
            try:
                _del_btn.state(["disabled"])
            except Exception:
                _del_btn.configure(state="disabled")
        _del_btn.pack(side="left", padx=(12, 0))

        # --- keep dropdown + run menu in sync ---
        try:
            self._init_quote_selection_vars()
            self._rebuild_quote_selector_menu()
        except Exception:
            pass
        try:
            self._rebuild_run_menu()
        except Exception:
            pass


    def _blank_quote_state(self) -> dict:
        try:
            road_default = f"{DEFAULT_ROAD_RATE_COMBINED:.2f}"
        except Exception:
            road_default = "1.10"
        return {
            "origin": bool(self.var_origin.get()) if hasattr(self, "var_origin") else True,
            "freight": bool(self.var_freight.get()) if hasattr(self, "var_freight") else True,
            "dest": bool(self.var_dest.get()) if hasattr(self, "var_dest") else True,
            "origin_addr": self.ent_origin.get() if hasattr(self, "ent_origin") else "",
            "dest_addr": self.ent_dest.get() if hasattr(self, "ent_dest") else "",
            "mode": self.cmb_mode.get() if hasattr(self, "cmb_mode") else "FCL",
            "volume": self.ent_volume.get() if hasattr(self, "ent_volume") else "",
            "pol": self.ent_pol.get() if hasattr(self, "ent_pol") else "",
            "pod": self.ent_pod.get() if hasattr(self, "ent_pod") else "",
            "fcl_choice": self.cmb_fcl_choice.get() if hasattr(self, "cmb_fcl_choice") else "Auto (best price)",
            "road_type": self.cmb_road_type.get() if hasattr(self, "cmb_road_type") else "Combined",
            "road_rate": self.ent_road_rate.get() if hasattr(self, "ent_road_rate") else road_default,
            "destonly_label": self.var_destonly_label.get() if hasattr(self, "var_destonly_label") else "(auto)",
            "destonly_fcl": self.cmb_destonly_fcl.get() if hasattr(self, "cmb_destonly_fcl") else "20FT",
            "destonly_gross": self.ent_destonly_gross.get() if hasattr(self, "ent_destonly_gross") else "0.00",
            "destonly_airkg": self.ent_destonly_airkg.get() if hasattr(self, "ent_destonly_airkg") else "0",
            "destonly_rate": self.ent_destonly_rate.get() if hasattr(self, "ent_destonly_rate") else "",
        }

    def _collect_ui_state(self) -> dict:
        # For now we reuse blank builder reading directly from UI
        return self._blank_quote_state()

    def _apply_quote_state_to_ui(self, st: dict):
        def set_entry(widget, value):
            try: widget.delete(0, tk.END); widget.insert(0, value)
            except Exception: pass
        try:
            if hasattr(self, "var_origin"): self.var_origin.set(bool(st.get("origin", True)))
            if hasattr(self, "var_freight"): self.var_freight.set(bool(st.get("freight", True)))
            if hasattr(self, "var_dest"): self.var_dest.set(bool(st.get("dest", True)))
        except Exception: pass
        set_entry(getattr(self, "ent_origin", None), st.get("origin_addr",""))
        set_entry(getattr(self, "ent_dest", None),   st.get("dest_addr",""))
        try:
            if hasattr(self, "cmb_mode"): self.cmb_mode.set(st.get("mode","FCL"))
        except Exception: pass
        set_entry(getattr(self, "ent_volume", None), st.get("volume",""))
        set_entry(getattr(self, "ent_pol", None), st.get("pol",""))
        set_entry(getattr(self, "ent_pod", None), st.get("pod",""))
        try:
            if hasattr(self, "cmb_fcl_choice"): self.cmb_fcl_choice.set(st.get("fcl_choice","Auto (best price)"))
        except Exception: pass
        try:
            if hasattr(self, "cmb_road_type"): self.cmb_road_type.set(st.get("road_type","Combined"))
        except Exception: pass
        set_entry(getattr(self, "ent_road_rate", None), st.get("road_rate","1.10"))
        try:
            if hasattr(self, "var_destonly_label"): self.var_destonly_label.set(st.get("destonly_label","(auto)"))
        except Exception: pass
        try:
            if hasattr(self, "cmb_destonly_fcl"): self.cmb_destonly_fcl.set(st.get("destonly_fcl","20FT"))
        except Exception: pass
        set_entry(getattr(self, "ent_destonly_gross", None), st.get("destonly_gross","0.00"))
        set_entry(getattr(self, "ent_destonly_airkg", None), st.get("destonly_airkg","0"))
        set_entry(getattr(self, "ent_destonly_rate", None), st.get("destonly_rate",""))
        try: self._toggle_inputs()
        except Exception: pass

    def _sync_active_quote_state(self):
        if 0 <= getattr(self, "_active_quote", 0) < len(getattr(self, "_quotes", [])):
            self._quotes[self._active_quote] = self._collect_ui_state()

    def _switch_to_quote(self, i: int):
        if i == getattr(self, "_active_quote", 0): return
        self._sync_active_quote_state()
        self._active_quote = i
        self._apply_quote_state_to_ui(self._quotes[i])
        self._render_quote_bar()

    def _add_quote(self):
        if not hasattr(self, "_quotes") or not self._quotes:
            self._quotes = [self._collect_ui_state()]; self._active_quote = 0
        else:
            self._sync_active_quote_state()
            self._quotes.append(self._blank_quote_state())
            self._active_quote = len(self._quotes) - 1
        self._apply_quote_state_to_ui(self._quotes[self._active_quote])
        self._render_quote_bar()
        try: self._rebuild_run_menu()
        except Exception: pass

        # --- keep dropdown + run menu in sync ---
        try:
            self._init_quote_selection_vars()
            self._rebuild_quote_selector_menu()
        except Exception:
            pass
        try:
            self._rebuild_run_menu()
        except Exception:
            pass


    def _delete_active_quote(self):
        """Remove the active quote while keeping at least one quote in the list."""
        if not hasattr(self, "_quotes") or not self._quotes:
            self._quotes = [self._blank_quote_state()]
            self._active_quote = 0
        elif len(self._quotes) == 1:
            # Clear the single quote instead of removing it
            self._quotes[0] = self._blank_quote_state()
            self._active_quote = 0
        else:
            try:
                del self._quotes[self._active_quote]
            except Exception:
                pass
            if self._active_quote >= len(self._quotes):
                self._active_quote = len(self._quotes) - 1
        try:
            self._apply_quote_state_to_ui(self._quotes[self._active_quote])
            self._render_quote_bar()
            self._rebuild_run_menu()
        except Exception:
            pass

        # --- keep dropdown + run menu in sync ---
        try:
            self._init_quote_selection_vars()
            self._rebuild_quote_selector_menu()
        except Exception:
            pass
        try:
            self._rebuild_run_menu()
        except Exception:
            pass


    def _rebuild_run_menu(self):
        try: self._run_menu.delete(0, "end")
        except Exception: return
        for i in range(len(getattr(self, "_quotes", []))):
            self._run_menu.add_command(label=f"Quote {i+1}", command=lambda i=i: self._run_for_quote(i))
        if len(getattr(self, "_quotes", [])) >= 2:
            self._run_menu.add_separator()
            self._run_menu.add_command(label="All quotes (multiple PDFs)", command=self._run_for_all_quotes)

    def _run_for_quote(self, i: int):
        self._sync_active_quote_state()
        cur = getattr(self, "_active_quote", 0)
        try:
            if 0 <= i < len(getattr(self, "_quotes", [])):
                self._apply_quote_state_to_ui(self._quotes[i]); self._active_quote = i
            self.run()
        finally:
            try:
                if 0 <= cur < len(getattr(self, "_quotes", [])):
                    self._apply_quote_state_to_ui(self._quotes[cur]); self._active_quote = cur
                self._render_quote_bar()
            except Exception: pass

    def select_quotes(self):
        """Open a simple selection dialog to pick which quotes to generate."""
        # Ensure selection vars exist and match number of quotes
        try:
            self._init_quote_selection_vars()
        except Exception:
            if not hasattr(self, "_sel_vars") or not self._sel_vars:
                self._sel_vars = [tk.BooleanVar(value=True)]
        win = tk.Toplevel(self)
        win.title("Select quotes")
        win.transient(self)
        try: win.grab_set()
        except Exception: pass
        # min size & center on parent
        try:
            win.update_idletasks()
            win.minsize(400, 260)
            px = self.winfo_rootx(); py = self.winfo_rooty()
            pw = self.winfo_width(); ph = self.winfo_height()
            ww = win.winfo_reqwidth(); wh = win.winfo_reqheight()
            if pw <= 1 or ph <= 1:
                sw = win.winfo_screenwidth(); sh = win.winfo_screenheight()
                x = int((sw - ww) / 2); y = int((sh - wh) / 2)
            else:
                x = int(px + (pw - ww) / 2); y = int(py + (ph - wh) / 2)
            win.geometry(f"{max(ww, 400)}x{max(wh, 260)}+{x}+{y}")
        except Exception:
            pass
        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)
        # grid columns for layout: [list/checks] [spacer] [merge]
        frm.grid_columnconfigure(0, weight=1)
        frm.grid_columnconfigure(1, weight=0)
        frm.grid_columnconfigure(2, weight=0)
        ttk.Label(frm, text="Choose the quotes to generate:").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0,8))
        n = len(getattr(self, "_sel_vars", []))
        for i in range(n):
            ttk.Checkbutton(frm, text=f"Quote {i+1}", variable=self._sel_vars[i]).grid(row=i+1, column=0, sticky="w")
        # helpers
        btns = ttk.Frame(frm)
        btns.grid(row=n+1, column=0, columnspan=2, sticky="w", pady=(8,0))
        self._var_merge = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="Merge into one PDF", variable=self._var_merge).grid(row=n+1, column=2, sticky="e")
        ttk.Button(btns, text="Select all", command=self._on_select_all_quotes).pack(side="left")
        ttk.Button(btns, text="Clear", command=self._on_clear_all_quotes).pack(side="left", padx=(6,0))
        # actions
        act = ttk.Frame(frm)
        act.grid(row=n+2, column=0, columnspan=3, sticky="ew", pady=(12,0))
        act.grid_columnconfigure(0, weight=1)
        ttk.Label(act, text="Ready.").grid(row=0, column=0, sticky="w")
        def _do():
            try:
                self._run_for_selected_quotes(merge=bool(self._var_merge.get()))
                try: win.destroy()
                except Exception: pass
            except Exception as e:
                messagebox.showerror("Select quotes", f"Failed: {e}")
        ttk.Button(act, text="Generate selected", command=_do).grid(row=0, column=1, sticky="e")
        ttk.Button(act, text="Close", command=win.destroy).grid(row=0, column=2, sticky="e", padx=(6,0))

    def _run_for_selected_quotes(self, merge=False):
        """Generate PDFs for the quotes currently checked in the selection vars."""
        # snapshot current quote
        cur = getattr(self, "_active_quote", 0)
        # make sure UI state is saved for active quote
        try:
            self._sync_active_quote_state()
        except Exception:
            pass
        # compute selected indices
        selected = []
        try:
            for i, var in enumerate(getattr(self, "_sel_vars", [])):
                try:
                    if var.get(): selected.append(i)
                except Exception:
                    pass
        except Exception:
            selected = [cur]
        if not selected:
            selected = [cur]
        # run each

        if merge and len(selected) > 1:
            import tempfile, os
            from tkinter import filedialog
            # choose final merged file once
            from tkinter import filedialog
            final_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="Quotes_merged.pdf",
                                                      filetypes=[("PDF","*.pdf")], title="Save merged PDF")
            if not final_path:
                return
            import tempfile, os
            tmpdir = tempfile.mkdtemp(prefix="quotes_merge_")
            made_files = []
            try:
                for i in selected:
                    try:
                        if 0 <= i < len(getattr(self, "_quotes", [])):
                            self._apply_quote_state_to_ui(self._quotes[i]); self._active_quote = i
                        part_path = os.path.join(tmpdir, f"part_{len(made_files)+1:02d}.pdf")
                        self.run(output_path=part_path)
                        made_files.append(part_path)
                    except Exception:
                        pass
                merged = self._merge_pdfs(made_files, final_path)
                if merged:
                    try: self.set_status(f"Merged {len(made_files)} PDFs → {final_path}")
                    except Exception: pass
            finally:
                pass
            try:
                if 0 <= cur < len(getattr(self, "_quotes", [])):
                    self._apply_quote_state_to_ui(self._quotes[cur]); self._active_quote = cur
                self._render_quote_bar()
            except Exception:
                pass
            return
            tmpdir = tempfile.mkdtemp(prefix="quotes_merge_")
            made_files = []
            try:
                # Monkeypatch asksaveasfilename to auto-save each quote to a temp file
                real_asksave = filedialog.asksaveasfilename
                def _fake_asksaveas(**kwargs):
                    idx = len(made_files) + 1
                    p = os.path.join(tmpdir, f"part_{idx:02d}.pdf")
                    return p
                filedialog.asksaveasfilename = _fake_asksaveas
                try:
                    for i in selected:
                        try:
                            if 0 <= i < len(getattr(self, "_quotes", [])):
                                self._apply_quote_state_to_ui(self._quotes[i]); self._active_quote = i
                            self.run()
                            made_files.append(os.path.join(tmpdir, f"part_{len(made_files):02d}.pdf"))
                        except Exception:
                            pass
                finally:
                    filedialog.asksaveasfilename = real_asksave
                # Merge
                try:
                    merged = self._merge_pdfs(made_files, final_path)
                    if merged:
                        try: self.set_status(f"Merged {len(made_files)} PDFs → {final_path}")
                        except Exception: pass
                except Exception as e:
                    try: messagebox.showerror("Merge PDFs", f"Failed to merge: {e}")
                    except Exception: pass
            finally:
                pass  # leave tmp files in case of troubleshooting
            # restore previous
            try:
                if 0 <= cur < len(getattr(self, "_quotes", [])):
                    self._apply_quote_state_to_ui(self._quotes[cur]); self._active_quote = cur
                self._render_quote_bar()
            except Exception:
                pass
            return
        try:
            for i in selected:
                try:
                    if 0 <= i < len(getattr(self, "_quotes", [])):
                        self._apply_quote_state_to_ui(self._quotes[i])
                        self._active_quote = i
                    self.run()
                except Exception:
                    pass
        finally:
            # restore previous
            try:
                if 0 <= cur < len(getattr(self, "_quotes", [])):
                    self._apply_quote_state_to_ui(self._quotes[cur])
                    self._active_quote = cur
                self._render_quote_bar()
            except Exception:
                pass

    def _run_for_all_quotes(self):
        self._sync_active_quote_state()
        cur = getattr(self, "_active_quote", 0)
        try:
            for i in range(len(getattr(self, "_quotes", []))):
                try:
                    self._apply_quote_state_to_ui(self._quotes[i]); self._active_quote = i
                except Exception: pass
                self.run()
        finally:
            try:
                if 0 <= cur < len(getattr(self, "_quotes", [])):
                    self._apply_quote_state_to_ui(self._quotes[cur]); self._active_quote = cur
                self._render_quote_bar()
            except Exception: pass

    def _init_quote_selection_vars(self):
        # Create selection vars aligned with quotes; default select active quote
        n = len(getattr(self, "_quotes", [])) or 1
        self._sel_vars = [tk.BooleanVar(value=(i == getattr(self, "_active_quote", 0))) for i in range(n)]

    def _rebuild_quote_selector_menu(self):
        # Ensure we have vars for each quote
        n = len(getattr(self, "_quotes", [])) or 1
        if not hasattr(self, "_sel_vars") or len(self._sel_vars) != n:
            self._sel_vars = [tk.BooleanVar(value=(i == getattr(self, "_active_quote", 0))) for i in range(n)]
        try:
            self._sel_menu.delete(0, "end")
        except Exception:
            return
        # Add one checkbutton per quote
        for i in range(n):
            self._sel_menu.add_checkbutton(label=f"Quote {i+1}", variable=self._sel_vars[i])
        self._sel_menu.add_separator()
        self._sel_menu.add_command(label="Select all", command=self._on_select_all_quotes)
        self._sel_menu.add_command(label="Clear all", command=self._on_clear_all_quotes)

    def _on_select_all_quotes(self):
        for v in getattr(self, "_sel_vars", []):
            try: v.set(True)
            except Exception: pass

    def _on_clear_all_quotes(self):
        for v in getattr(self, "_sel_vars", []):
            try: v.set(False)
            except Exception: pass
        # Keep at least the active one selected for convenience
        if getattr(self, "_sel_vars", []) and 0 <= getattr(self, "_active_quote", 0) < len(self._sel_vars):
            try: self._sel_vars[self._active_quote].set(True)
            except Exception: pass

    def _get_selected_quote_indices(self):
        out = []
        for i, v in enumerate(getattr(self, "_sel_vars", [])):
            try:
                if v.get(): out.append(i)
            except Exception:
                pass
        return out



    def _on_generate_pdf(self):
        """Generate PDFs for selected quotes without losing field values."""
        sel = self._get_selected_quote_indices() if hasattr(self, "_get_selected_quote_indices") else []
        if not sel:
            sel = [getattr(self, "_active_quote", 0)]

        # Persist current UI to its quote before switching
        try:
            self._sync_active_quote_state()
        except Exception:
            pass

        created, skipped = [], []
        cur = getattr(self, "_active_quote", 0)

        try:
            for i in sel:
                # Snapshot original quote state
                try:
                    saved_state = copy.deepcopy(self._quotes[i])
                except Exception:
                    try:
                        saved_state = dict(self._quotes[i])
                    except Exception:
                        saved_state = None

                # Apply quote to UI
                try:
                    self._apply_quote_state_to_ui(self._quotes[i]); self._active_quote = i
                except Exception:
                    pass

                # Soft validation
                ok, reasons = self._validate_current_state()
                if not ok:
                    skipped.append((i, ", ".join(reasons)))
                    if saved_state is not None:
                        try: self._quotes[i] = saved_state
                        except Exception: pass
                    continue

                # Run generator
                self.run()
                created.append(i)

                # If run() cleared inputs, restore snapshot to both UI and memory
                try:
                    origin_empty = False
                    try: origin_empty = not bool(self.ent_origin.get().strip())
                    except Exception: pass
                    pol_empty = False
                    try: pol_empty = not bool(self.ent_pol.get().strip())
                    except Exception: pass
                    if saved_state is not None and (origin_empty and pol_empty):
                        self._quotes[i] = saved_state
                        self._apply_quote_state_to_ui(self._quotes[i])
                except Exception:
                    pass

                # Persist after run() (in case run() changed derived fields)
                try:
                    self._sync_active_quote_state()
                except Exception:
                    pass
        finally:
            try:
                if 0 <= cur < len(getattr(self, "_quotes", [])):
                    self._apply_quote_state_to_ui(self._quotes[cur]); self._active_quote = cur
                self._render_quote_bar()
            except Exception:
                pass

        parts = []
        if created:
            parts.append("Generated: " + ", ".join(f"Quote {j+1}" for j in created))
        if skipped:
            parts.append("Skipped: " + "; ".join(f"Quote {j+1} ({why})" for j, why in skipped))
        try:
            self.status.set(" | ".join(parts) if parts else "No quotes generated.")
        except Exception:
            pass

    def _validate_current_state(self):
        """Return (ok, reasons) for the current UI state without showing popups."""
        reasons = []
        try:
            origin_on = bool(self.var_origin.get())
        except Exception:
            origin_on = False
        try:
            freight_on = bool(self.var_freight.get())
        except Exception:
            freight_on = False

        # Origin location required when Origin service is on
        try:
            if origin_on and not (self.ent_origin.get().strip()):
                reasons.append("Origin location")
        except Exception:
            pass

        # Freight minimal checks
        try:
            mode = self.cmb_mode.get().strip()
        except Exception:
            mode = ""
        if freight_on and mode.upper() == "FCL":
            try:
                pol = self.ent_pol.get().strip()
            except Exception:
                pol = ""
            try:
                pod = self.ent_pod.get().strip()
            except Exception:
                pod = ""
            if not pol or not pod:
                reasons.append("FCL ports (POL/POD)")
        return (len(reasons) == 0, reasons)
        # Use selected quotes; if none selected, fall back to active quote
        sel = self._get_selected_quote_indices()
        if not sel:
            sel = [getattr(self, "_active_quote", 0)]
        # Remember current active, then iterate and run for each
        cur = getattr(self, "_active_quote", 0)
        try:
            for i in sel:
                try:
                    self._apply_quote_state_to_ui(self._quotes[i]); self._active_quote = i
                except Exception:
                    pass
                self.run()  # existing routine that generates PDF from current UI state
        finally:
            try:
                # Restore previous active and UI
                self._apply_quote_state_to_ui(self._quotes[cur]); self._active_quote = cur
                self._render_quote_bar()
            except Exception:
                pass
        try:
            self.status.set(f"Generated PDF for: {', '.join(f'Quote {i+1}' for i in sel)}")
        except Exception:
            pass
        self._sync_active_quote_state()
        cur = getattr(self, "_active_quote", 0)
        try:
            for i in range(len(getattr(self, "_quotes", []))):
                try:
                    self._apply_quote_state_to_ui(self._quotes[i]); self._active_quote = i
                except Exception: pass
                self.run()
        finally:
            try:
                if 0 <= cur < len(getattr(self, "_quotes", [])):
                    self._apply_quote_state_to_ui(self._quotes[cur]); self._active_quote = cur
                self._render_quote_bar()
            except Exception: pass

        ttk.Entry(self.sec_destonly, textvariable=self.var_destonly_label, width=24).grid(row=0, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(self.sec_destonly, text="FCL container:").grid(row=0, column=2, sticky="e", padx=6, pady=4)
        self.cmb_destonly_fcl = ttk.Combobox(self.sec_destonly, values=["20FT","40FT","40HQ"], state="readonly", width=10)
        self.cmb_destonly_fcl.set("20FT")
        self.cmb_destonly_fcl.grid(row=0, column=3, sticky="w", padx=6, pady=4)

        ttk.Label(self.sec_destonly, text="LCL gross cbm:").grid(row=0, column=4, sticky="e", padx=6, pady=4)
        self.ent_destonly_gross = ttk.Entry(self.sec_destonly, width=10); self.ent_destonly_gross.insert(0,"0.00")
        self.ent_destonly_gross.grid(row=0, column=5, sticky="w", padx=6, pady=4)

        ttk.Label(self.sec_destonly, text="AIR chargeable kg:").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        self.ent_destonly_airkg = ttk.Entry(self.sec_destonly, width=10); self.ent_destonly_airkg.insert(0,"0")
        self.ent_destonly_airkg.grid(row=1, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(self.sec_destonly, text="Amount (€):").grid(row=0, column=2, sticky="e", padx=6, pady=4)
        self.ent_destonly_rate = ttk.Entry(self.sec_destonly, width=12)
        self.ent_destonly_rate.grid(row=0, column=3, sticky="w", padx=6, pady=4)
        # ttk.Label(self.sec_destonly, text="Tip: Paneel verschijnt als alléén Destination is aangevinkt (ROAD uitgesloten).").grid(row=1, column=4, columnspan=2, sticky="w", padx=6, pady=4)

        sec_doc.grid(row=9, column=0, columnspan=2, sticky="nsew", padx=0, pady=8)
        for c in range(5): sec_doc.grid_columnconfigure(c, weight=1 if c in (1,3) else 0)

        ttk.Label(sec_doc, text="Client name (optional):").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        self.ent_klant = ttk.Entry(sec_doc, width=32); self.ent_klant.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(sec_doc, text="Reference (optional):").grid(row=0, column=2, sticky="e", padx=6, pady=4)
        self.ent_ref = ttk.Entry(sec_doc, width=18); self.ent_ref.grid(row=0, column=3, sticky="ew", padx=6, pady=4)
        ttk.Label(sec_doc, text="Brand:").grid(row=0, column=4, sticky="e", padx=6, pady=4)
        self.cmb_brand = ttk.Combobox(sec_doc, state="readonly", width=14, values=["Voerman","Transpack"])
        self.cmb_brand.set("Voerman")
        self.cmb_brand.grid(row=0, column=5, sticky="w", padx=6, pady=4)

        ttk.Label(sec_doc, text="Excel file:").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        self.ent_excel = ttk.Entry(sec_doc, width=60); self.ent_excel.insert(0, EXCEL_PAD)
        self.ent_excel.grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(sec_doc, text="Browse…", command=self.kies_excel).grid(row=1, column=2, sticky="w", padx=6, pady=4)
        # (FCL lanes sheet field hidden)

        
        # PDF options
        self.var_pdf_show_rate = tk.BooleanVar(value=False)
        ttk.Checkbutton(sec_doc, text="Show 'Estimate Rate' column on PDF", variable=self.var_pdf_show_rate).grid(row=2, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(sec_doc, text="Logo (optional):").grid(row=2, column=2, sticky="e", padx=6, pady=4)
        self.ent_logo = ttk.Entry(sec_doc, width=28)
        self.ent_logo.grid(row=2, column=3, sticky="ew", padx=6, pady=4)
        ttk.Button(sec_doc, text="Browse…", command=self.kies_logo).grid(row=2, column=4, sticky="w", padx=6, pady=4)

        try:
            self._toggle_destonly_panel()
        except Exception:
            pass

        self._toggle_inputs()

        # Initialize multi-quote model
        self._quotes_init()
    # ---------- Email parsing UI ----------
    def do_parse_email(self):
        raw = self.txt_email.get("1.0", tk.END)
        data = parse_rfq_text(raw)
        self._apply_parsed_data(data, label="Parser (rules)")

    def do_ai_test(self):
        key = os.getenv("OPENAI_API_KEY")
        if not key:
            messagebox.showwarning("AI Test", "No OPENAI_API_KEY set.\nAdd it to .env or environment first.")
            return
        model = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
        versions = _sdk_versions()
        msg = [f"Key present ✓", f"Model: {model}", f"SDK: {versions}"]
        ok = False
        try:
            client = _load_new_client()
            if client is not None:
                r = client.chat.completions.create(model=model, messages=[{"role":"user","content":"Reply with: {\"ok\": true}"}], temperature=0, response_format={"type":"json_object"})
                content = r.choices[0].message.content
                import json as _json
                data = _json.loads(content)
                if data.get("ok") is True: ok = True
        except Exception as e:
            log(f"AI test (new SDK) failed: {e}")
            msg.append(f"New SDK error: {e}")
        if not ok and _has_legacy():
            try:
                import openai
                openai.api_key = key
                r = openai.ChatCompletion.create(model=model, messages=[{"role":"user","content":"Reply with: {\"ok\": true}"}], temperature=0)
                content = r["choices"][0]["message"]["content"]
                if "{“ok”" in content or '"ok": true' in content:
                    ok = True
            except Exception as e:
                log(f"AI test (legacy) failed: {e}")
                msg.append(f"Legacy SDK error: {e}")
        messagebox.showinfo("AI Test", "\n".join(msg + ([ "Test ✓ OK" ] if ok else [ "Test ✗ FAILED – see ai_debug.log" ])))

    def do_ai_parse_email(self):
        key = os.getenv("OPENAI_API_KEY")
        if not key:
            messagebox.showwarning("AI Parse", "No OPENAI_API_KEY found.\nUse a .env or environment variable.")
            return
        raw = self.txt_email.get("1.0", tk.END)
        if not raw.strip():
            messagebox.showinfo("AI Parse", "Paste an RFQ email first.")
            return
        self.set_status("AI parsing…"); self.update_idletasks()
        data = ai_parse_rfq_text(raw, model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"))
        if not data:
            messagebox.showerror("AI Parse", "No data returned or API error.\nSee ai_debug.log for details.\nFalling back to rules parser.")
            data = parse_rfq_text(raw)
        self._apply_parsed_data(data, label="AI Parse")
        self.set_status("Ready.")

    def _apply_parsed_data(self, data: dict, label="Parser"):
        if not data:
            messagebox.showinfo(label, "No fields recognised.")
            return
        applied = []
        if "mode" in data and data["mode"] in VALID_MODES:
            self.cmb_mode.set(data["mode"]); applied.append(f"Mode={data['mode']}")
        if "services" in data and isinstance(data["services"], list):
            s = set(x.lower() for x in data["services"])
            self.var_origin.set("origin" in s); self.var_freight.set("freight" in s); self.var_dest.set("destination" in s)
            applied.append("Services=" + ",".join(sorted(s))); self._toggle_inputs()
        if "origin_location" in data:
            self.ent_origin.delete(0, tk.END); self.ent_origin.insert(0, data["origin_location"]); applied.append(f"Origin={data['origin_location']}")
        if "destination_location" in data:
            self.ent_dest.delete(0, tk.END); self.ent_dest.insert(0, data["destination_location"]); applied.append(f"Destination={data['destination_location']}")
        if "pol" in data:
            self.ent_pol.delete(0, tk.END); self.ent_pol.insert(0, data["pol"]); applied.append(f"POL={data['pol']}")
        if "pod" in data:
            self.ent_pod.delete(0, tk.END); self.ent_pod.insert(0, data["pod"]); applied.append(f"POD={data['pod']}")
        if "volume_cbm" in data:
            try:
                self.ent_volume.delete(0, tk.END); self.ent_volume.insert(0, f"{float(data['volume_cbm']):.2f}")
                applied.append(f"Volume={data['volume_cbm']} m³")
            except Exception:
                pass
        messagebox.showinfo(label, "Prefilled:\n- " + "\n- ".join(applied) if applied else f"{label}: nothing new applied.")
        self._toggle_inputs()

    def _toggle_inputs(self):
        origin_on = self.var_origin.get(); dest_on = self.var_dest.get(); freight_on = self.var_freight.get()
        mode = self.cmb_mode.get().strip().upper()

        # Enable/disable origin/dest entries
        self.ent_origin.configure(state=("normal" if origin_on else "disabled"))
        self.ent_dest.configure(state=("normal" if dest_on else "disabled"))

        if freight_on:
            if mode in ("FCL", "LCL", "AIR"):
                # show ocean/air lanes panel, hide road
                self.sec_fcl.grid()
                self.sec_road.grid_remove()
                for w in (self.ent_pol, self.ent_pod, self.cmb_fcl_choice):
                    try: w.configure(state="normal")
                    except Exception: pass

                if mode == "AIR":
                    try: self.sec_fcl.configure(text="Freight – AIR lanes")
                    except Exception: pass
                    try:
                        self.lbl_pol.configure(text="Origin Airport (or IATA):")
                        self.lbl_pod.configure(text="Destination Airport (or IATA):")
                    except Exception: pass
                    # hide container controls
                    try:
                        self.lbl_container.grid_remove()
                        self.cmb_fcl_choice.grid_remove()
                    except Exception: pass

                elif mode == "LCL":
                    try: self.sec_fcl.configure(text="Freight – LCL lanes")
                    except Exception: pass
                    try:
                        self.lbl_pol.configure(text="FCL Origin port (code or name):")
                        self.lbl_pod.configure(text="FCL Destination port (code or name):")
                    except Exception: pass
                    # hide container controls
                    try:
                        self.lbl_container.grid_remove()
                        self.cmb_fcl_choice.grid_remove()
                    except Exception: pass

                else:  # FCL
                    try: self.sec_fcl.configure(text="Freight – FCL lanes")
                    except Exception: pass
                    try:
                        self.lbl_pol.configure(text="FCL Origin port (code or name):")
                        self.lbl_pod.configure(text="FCL Destination port (code or name):")
                    except Exception: pass
                    try:
                        self.lbl_container.grid()
                        self.cmb_fcl_choice.grid()
                    except Exception: pass
            elif mode == "ROAD":
                self.sec_road.grid()
                self.sec_fcl.grid_remove()
            else:
                self.sec_fcl.grid_remove()
                self.sec_road.grid_remove()
        else:
            self.sec_fcl.grid_remove()
            self.sec_road.grid_remove()
        self._toggle_destonly_panel()


    # ---------- Client DB + UI ----------
    def _client_db_path(self):
        try:
            base = os.path.dirname(__file__)
        except Exception:
            base = os.getcwd()
        return os.path.join(base, "clients_agents.json")

    def _client_db_load(self):
        p = self._client_db_path()
        if os.path.isfile(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def _client_db_save(self, data: dict):
        p = self._client_db_path()
        try:
            with open(p, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showwarning("Client DB", f"Could not save clients file: {e}")

    def _refresh_saved_clients(self):
        try:
            data = self._client_db_load()
            names = sorted(list(data.keys()))
            self.cmb_saved_client["values"] = names
        except Exception:
            pass
    # --- LGR numbering ----------------------------------------------------
    
    # --- LGR numbering (reset each ISO week) -------------------------------
    def _lgr_counter_path(self):
        try:
            base = os.path.dirname(__file__)
        except Exception:
            base = os.getcwd()
        return os.path.join(base, "lgr_counter.json")

    def _current_yearweek(self):
        iso = datetime.now().isocalendar()
        return f"{iso[0]}-{iso[1]:02d}"  # e.g. 2025-36

    def _next_lgr_number(self) -> int:
        path = self._lgr_counter_path()
        cur_week = self._current_yearweek()
        next_n = 1
        try:
            if os.path.isfile(path):
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f) or {}
                if data.get("week") == cur_week:
                    next_n = int(data.get("next", 1))
                else:
                    next_n = 1
        except Exception:
            next_n = 1
        # store incremented counter with current week
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump({"week": cur_week, "next": next_n + 1}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return next_n



    

    def _update_vat_state(self):
        # Determine VAT applicability based on client type and EU countries
        try:
            client_is_private = str(self.var_client_type.get()).strip().lower().startswith('private')
        except Exception:
            client_is_private = True
        o = (self.var_origin_country.get() or '').strip().upper()
        d = (self.var_dest_country.get() or '').strip().upper()
        vat_applies = client_is_private and (o in EU_COUNTRIES) and (d in EU_COUNTRIES)
        # central VAT gate (keeps checkbox outside):
        try:
            chk = bool(self.var_pdf_show_vat.get())
        except Exception:
            chk = False
        self.vat_applies = should_show_vat(is_private=client_is_private, is_agent=False, checkbox=chk, origin_iso=o, dest_iso=d)
        try:
            rate = float(self.var_vat_rate.get() or 0)
        except Exception:
            rate = 0.0
        
        # Update VAT status text (no auto-toggling of user checkboxes)
        self.var_vat_status.set(f"VAT: {'Applies' if vat_applies else 'Not applicable'}" + (f" ({rate:.0f}%)" if vat_applies else ""))
        if vat_applies:
            self.vat_memo = f"Prices excl. VAT ({rate:.0f}% VAT applies for private intra-EU moves)."
        else:
            self.vat_memo = "VAT not applicable (export/import / one leg outside EU)."

    def _toggle_client_section(self):
        if bool(self.var_client_show.get()):
            self.sec_client.grid()
        else:
            self.sec_client.grid_remove()

    def client_load_selected(self):
        key = (self.cmb_saved_client.get() or "").strip()
        if not key:
            return
        data = self._client_db_load()
        rec = data.get(key, {})
        if not rec:
            messagebox.showinfo("Client", "No saved data for selection."); return
        self.var_client_type.set(rec.get("type", "PRIVATE"))
        fields = [
            (self.ent_client_name, rec.get("name","")),
            (self.ent_assignee, rec.get("assignee","")),
            (self.ent_addr1, rec.get("addr1","")),
            (self.ent_addr2, rec.get("addr2","")),
            (self.ent_city_postal, rec.get("city_postal","")),
            (self.ent_country, rec.get("country","")),
            (self.ent_email, rec.get("email","")),
        ]
        for w, val in fields:
            try:
                w.delete(0, tk.END); w.insert(0, val)
            except Exception:
                pass
        self._toggle_client_ui()

    def client_save(self):
        t = (self.var_client_type.get() or "PRIVATE").upper()
        name = (self.ent_client_name.get() or "").strip()
        if not name:
            messagebox.showinfo("Client", "Fill in Name (client/agent) before saving."); return
        key = f"{t}:{name}"
        rec = {
            "type": t,
            "name": name,
            "assignee": (self.ent_assignee.get() or "").strip(),
            "addr1": (self.ent_addr1.get() or "").strip(),
            "addr2": (self.ent_addr2.get() or "").strip(),
            "city_postal": (self.ent_city_postal.get() or "").strip(),
            "country": (self.ent_country.get() or "").strip(),
            "email": (self.ent_email.get() or "").strip(),
        }
        data = self._client_db_load()
        data[key] = rec
        self._client_db_save(data)
        self._refresh_saved_clients()
        messagebox.showinfo("Client", f"Saved: {key}")

    def client_delete(self):
        key = (self.cmb_saved_client.get() or "").strip()
        if not key:
            messagebox.showinfo("Client", "Select a saved client first."); return
        data = self._client_db_load()
        if key in data:
            del data[key]
            self._client_db_save(data)
            self._refresh_saved_clients()
            messagebox.showinfo("Client", f"Deleted: {key}")

    def _toggle_client_ui(self):
        t = (self.var_client_type.get() or "PRIVATE").upper()
        try:
            if t == "AGENT":
                self.ent_assignee.configure(state="normal")
            else:
                self.ent_assignee.configure(state="disabled")
        except Exception:
            pass

    def kies_excel(self):
        p = filedialog.askopenfilename(title="Choose tarieven.xlsx", filetypes=[("Excel","*.xlsx *.xls")])
        if p: self.ent_excel.delete(0, tk.END); self.ent_excel.insert(0, p)

    def kies_logo(self):
        p = filedialog.askopenfilename(title="Choose logo (PNG/JPG)", filetypes=[("Image","*.png *.jpg *.jpeg")])
        if p: self.ent_logo.delete(0, tk.END); self.ent_logo.insert(0, p)

    def set_status(self, s: str):
        self.status.set(s); self.update_idletasks()


    def _toggle_destonly_panel(self):
        only_dest = (not self.var_origin.get()) and (not self.var_freight.get()) and self.var_dest.get()
        mode = self.cmb_mode.get().strip().upper()
        if only_dest and mode in ("FCL","LCL","AIR"):
            # Update auto label + defaults
            try:
                vol = float(self.ent_volume.get().replace(",", "."))
                gross = round(vol * 1.2, 2) if mode=="LCL" else 0.0
            except Exception:
                gross = 0.0
            if mode == "FCL":
                self.var_destonly_label.set("DTHC")
            elif mode == "LCL":
                self.var_destonly_label.set("NVOCC charges")
            elif mode == "AIR":
                self.var_destonly_label.set("ATHC")
            # Show the frame
            self.sec_destonly.grid()
            # Toggle field visibility per mode
            try:
                if mode == "FCL":
                    # show FCL container, hide LCL/AIR fields
                    self.cmb_destonly_fcl.grid()
                    self.ent_destonly_gross.grid_remove()
                    self.ent_destonly_airkg.grid_remove()
                elif mode == "LCL":
                    self.cmb_destonly_fcl.grid_remove()
                    self.ent_destonly_gross.grid()
                    self.ent_destonly_airkg.grid_remove()
                    # Set gross cbm default
                    try:
                        self.ent_destonly_gross.delete(0, tk.END)
                        self.ent_destonly_gross.insert(0, f"{gross:.2f}")
                    except Exception:
                        pass
                elif mode == "AIR":
                    self.cmb_destonly_fcl.grid_remove()
                    self.ent_destonly_gross.grid_remove()
                    self.ent_destonly_airkg.grid()
            except Exception:
                pass
            # Always show override rate controls and the tip row
            try:
                self.ent_destonly_rate.grid()
            except Exception:
                pass
        else:
            try:
                self.sec_destonly.grid_remove()
            except Exception:
                pass

    def run(self, output_path: str | None = None):
        """Generate a PDF for the active quote.
        If output_path is provided, saves directly to that path (no Save-as dialog).
        Otherwise, opens a Save-as dialog with a sensible default filename.
        """
        try: self._run_safe(output_path)
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error:\n{e}\n\n{traceback.format_exc()}")
            self.set_status("Error. See message.")

    def _run_safe(self, output_path: str | None = None):
        """Internal runner with validation, Excel/geo lookups and PDF render.
        When output_path is provided, bypasses the Save-as dialog.
        """
        use_origin  = self.var_origin.get(); use_freight = self.var_freight.get(); use_dest = self.var_dest.get()
        if not (use_origin or use_freight or use_dest):
            messagebox.showwarning("Input", "Select at least one service (Origin / Freight / Destination)."); return

        origin_addr = self.ent_origin.get().strip(); dest_addr = self.ent_dest.get().strip()
        if (use_origin or use_freight) and not origin_addr:
            messagebox.showwarning("Input", "Enter an Origin location or untick related services."); return
        if use_dest and not dest_addr:
            messagebox.showwarning("Input", "Enter a Destination location or untick related services."); return

        mode = self.cmb_mode.get().strip().upper()
        # Determine countries for origin/destination (skip origin when only Destination is selected)
        try:
            cc_o = None
            if (use_origin or use_freight) and origin_addr:
                (_o_ll, cc_o, _name_o) = geocode_country(origin_addr)
            cc_d = None
            if dest_addr:
                (_d_ll, cc_d, _name_d) = geocode_country(dest_addr)
        except Exception as e:
            messagebox.showerror("Geocoding", f"Could not geocode country: {e}"); return
        try:
            volume_cbm = float(self.ent_volume.get().replace(",", ".").strip())
            if volume_cbm <= 0: raise ValueError
        except ValueError:
            messagebox.showwarning("Input", "Volume (m³) invalid or ≤ 0."); return

        klant = self.ent_klant.get().strip() or None; ref = self.ent_ref.get().strip() or None
        excel = self.ent_excel.get().strip() or EXCEL_PAD; logo = self.ent_logo.get().strip() or None
        fcl_sheet_name = (getattr(self, 'ent_fclsheet', None).get().strip() if getattr(self, 'ent_fclsheet', None) else "")
        fcl_choice = self.cmb_fcl_choice.get().strip() if self.cmb_fcl_choice.winfo_exists() else "Auto (best price)"

        if not os.path.isfile(excel):
            messagebox.showwarning("File not found", f"Excel not found: {excel}\nPick the correct file now.")
            picked = filedialog.askopenfilename(title="Choose tarieven.xlsx", filetypes=[("Excel","*.xlsx *.xls")])
            if not picked: self.set_status("Cancelled: no Excel chosen."); return
            excel = picked; self.ent_excel.delete(0, tk.END); self.ent_excel.insert(0, excel)

        try: df = lees_services_sheet(excel, SERVICES_SHEET)
        except Exception as e:
            messagebox.showerror("Excel read error (services)", f"{e}"); return

        # Distances to warehouse
        okm = onote = None; dkm = dnote = None
        if use_origin: okm, onote = afstand_km_via_warehouse(origin_addr, WAREHOUSE_LOCATION)
        if use_dest:   dkm, dnote = afstand_km_via_warehouse(dest_addr, WAREHOUSE_LOCATION)

        charges_rows = []
        dest_only_rows = []
        only_dest = (not use_origin) and (not use_freight) and use_dest
        if only_dest and mode in ("FCL","LCL","AIR"):
            try:
                df_dest, dest_cols = lees_dest_only_charges(excel, DEST_ONLY_SHEET)
            except Exception as e:
                messagebox.showerror("Excel (DestOnlyCharges)", f"{e}"); return
            try:
                (_, cc_d, name_d) = geocode_country(dest_addr)
                dest_country = ("Netherlands" if (cc_d=="nl" and name_d) else (name_d or "Netherlands"))
            except Exception:
                dest_country = "Netherlands"
            container = None; gross_cbm_val = None; air_kg_val = None
            if mode == "FCL":
                container = self.cmb_destonly_fcl.get().strip() if hasattr(self,"cmb_destonly_fcl") and self.cmb_destonly_fcl.winfo_exists() else "20FT"
            elif mode == "LCL":
                try:
                    gross_cbm_val = float((self.ent_destonly_gross.get() or "0").replace(",", "."))
                except Exception:
                    try:
                        gross_cbm_val = float(self.ent_volume.get().replace(",", ".")) * 1.2
                    except Exception:
                        gross_cbm_val = 0.0
            elif mode == "AIR":
                try:
                    air_kg_val = float((self.ent_destonly_airkg.get() or "0").replace(",", "."))
                except Exception:
                    air_kg_val = 0.0
                # Fallback: if not provided, derive chargeable kg from volume (cbm * 167)
                try:
                    if not air_kg_val:
                        vol_tmp = float((self.ent_volume.get() or "0").replace(",", "."))
                        air_kg_val = round(vol_tmp * 167.0, 2)
                except Exception:
                    pass
            try:
                (cname, qty_val, qty_unit, rate_str, amt) = find_dest_only_rate(df_dest, dest_cols, dest_country, mode, container, gross_cbm_val, air_kg_val)
            except Exception as e:
                messagebox.showerror("Dest-only rates", f"{e}"); return
            # override rate
            override_rate = None
            try:
                t = (self.ent_destonly_rate.get() or "").strip()
                if t: override_rate = float(t.replace(",", "."))
            except Exception:
                override_rate = None
            if override_rate is not None:
                if mode == "FCL":
                    amt = override_rate
                    rate_str = eur(override_rate)
                elif mode == "LCL":
                    amt = round(override_rate * (qty_val or 0.0), 2)
                    rate_str = eur(override_rate) + " / cbm gross"
                elif mode == "AIR":
                    amt = round(override_rate * (qty_val or 0.0), 2)
                    rate_str = eur(override_rate) + " / kg"
            qty_disp = f"{qty_val:g} {qty_unit}" if qty_unit else (f"{qty_val:g}" if qty_val is not None else "-")
            label_override = (self.var_destonly_label.get() or "").strip()
            descr = label_override if label_override and label_override != "(auto)" else cname
            dest_only_rows.append({"descr": descr, "qty": qty_disp, "rate": rate_str, "amount": amt})


        # ========== ROAD logic ==========
        if mode == "ROAD":
            # Determine domestic vs international (only for the addresses we actually need)
            try:
                cc_o = cc_d = None
                if (use_origin or use_freight) and origin_addr:
                    (_, cc_o, _name_o) = geocode_country(origin_addr)
                if (use_dest or use_freight) and dest_addr:
                    (_, cc_d, _name_d) = geocode_country(dest_addr)
                if use_freight and (not origin_addr or not dest_addr):
                    raise ValueError("Both origin and destination required for road freight.")
            except Exception as e:
                messagebox.showerror("Geocoding", f"Could not geocode country: {e}"); return
            is_domestic = (cc_o == "nl" and cc_d == "nl") if (cc_o and cc_d) else False

            if is_domestic and use_origin and use_dest:
                # One consolidated line: Domestic door-to-door services (ROAD)
                try:
                    p_origin = LineParams("Origin services", "Origin", origin_addr, "ROAD", volume_cbm, okm)
                    row_o = match_service_rij(df, p_origin); amt_o = calc_service_prijs(row_o, volume_cbm)
                    p_dest = LineParams("Destination services", "Destination", dest_addr, "ROAD", volume_cbm, dkm)
                    row_d = match_service_rij(df, p_dest); amt_d = calc_service_prijs(row_d, volume_cbm)
                    total_amt = round(amt_o + amt_d, 2)
                except Exception as e:
                    messagebox.showerror("Rates (ROAD)", f"Could not match ROAD origin/destination rows: {e}"); return

                descr = f"Domestic door-to-door services (ROAD) {origin_addr} – {dest_addr}"
                charges_rows.append({"descr": descr, "qty": "-", "rate": "-", "amount": total_amt})

            else:
                # International ROAD: individual origin/dest + a Road Freight line (if ticked)
                if use_origin:
                    try:
                        p_origin = LineParams("Origin services", "Origin", origin_addr, "ROAD", volume_cbm, okm)
                        row_o = match_service_rij_strict_op(df, p_origin) if cc_o != "nl" else match_service_rij(df, p_origin)
                        amt_o = calc_service_prijs(row_o, volume_cbm)
                        descr_o = f"Origin services (ROAD) {origin_addr}"
                        if str(row_o[COLS["rate_type"]]).strip().upper() == "FLAT":
                            charges_rows.append({"descr": descr_o, "qty": "1 flat", "rate": eur(float(row_o[COLS['flat']])), "amount": amt_o})
                        else:
                            charges_rows.append({"descr": descr_o, "qty": f"{volume_cbm:g} m³", "rate": f"{eur(float(row_o[COLS['rate_per_cbm']]))} / m³", "amount": amt_o})
                    except Exception as e:
                        messagebox.showerror("Rates (ROAD origin)", f"{e}"); return

                if use_freight:
                    try:
                        km_between, note = road_distance_between_addrs(origin_addr, dest_addr)
                        try:
                            per_km = float(self.ent_road_rate.get().replace(",", ".").strip())
                        except Exception:
                            per_km = DEFAULT_ROAD_RATE_COMBINED if self.cmb_road_type.get() == "Combined" else DEFAULT_ROAD_RATE_DIRECT
                        amt_f = round(km_between * per_km, 2)
                        descr_f = f"Road freight ({self.cmb_road_type.get()}) {origin_addr} – {dest_addr}"
                        charges_rows.append({"descr": descr_f, "qty": f"{km_between:.0f} km", "rate": f"{eur(per_km)} / km", "amount": amt_f})
                    except Exception as e:
                        messagebox.showerror("Road freight", f"Could not compute road distance:\n{e}"); return

                if use_dest:
                    try:
                        p_dest = LineParams("Destination services", "Destination", dest_addr, "ROAD", volume_cbm, dkm)
                        row_d = match_service_rij_strict_op(df, p_dest) if cc_d != "nl" else match_service_rij(df, p_dest)
                        amt_d = calc_service_prijs(row_d, volume_cbm)
                        descr_d = f"Destination services (ROAD) {dest_addr}"
                        if str(row_d[COLS["rate_type"]]).strip().upper() == "FLAT":
                            charges_rows.append({"descr": descr_d, "qty": "1 flat", "rate": eur(float(row_d[COLS['flat']])), "amount": amt_d})
                        else:
                            charges_rows.append({"descr": descr_d, "qty": f"{volume_cbm:g} m³", "rate": f"{eur(float(row_d[COLS['rate_per_cbm']]))} / m³", "amount": amt_d})
                    except Exception as e:
                        messagebox.showerror("Rates (ROAD destination)", f"{e}"); return

        # ========== Non-ROAD logic ==========
        else:
            if use_origin:
                try:
                    p_origin = LineParams("Origin services", "Origin", origin_addr, mode, volume_cbm, okm)
                    row_o = match_service_rij_strict_op(df, p_origin) if cc_o != "nl" else match_service_rij(df, p_origin)
                    amt_o = calc_service_prijs(row_o, volume_cbm)
                    descr_o = f"Origin services ({row_o[COLS['mode']]}) {origin_addr}"
                    if str(row_o[COLS["rate_type"]]).strip().upper() == "FLAT":
                        charges_rows.append({"descr": descr_o, "qty": "1 flat", "rate": eur(float(row_o[COLS['flat']])), "amount": amt_o})
                    else:
                        charges_rows.append({"descr": descr_o, "qty": f"{volume_cbm:g} m³", "rate": f"{eur(float(row_o[COLS['rate_per_cbm']]))} / m³", "amount": amt_o})
                except Exception as e:
                    messagebox.showerror("Rates (Origin)", f"{e}"); return

            if use_freight:
                if mode == "LCL":
                    pol_in = self.ent_pol.get().strip(); pod_in = self.ent_pod.get().strip()
                    if not pol_in or not pod_in:
                        messagebox.showwarning("Input", "Enter both POL and POD (code or name) or untick Freight."); return
                    try:
                        if not fcl_sheet_name:
                            auto_name = auto_find_fcl_lanes_sheet(excel)
                            df_lanes = lees_fcl_lanes(excel, auto_name)
                        else:
                            df_lanes = lees_fcl_lanes(excel, fcl_sheet_name)
                    except Exception as e:
                        messagebox.showerror("Excel (lanes)", f"{e}"); return
                    try:
                        pol = resolve_port_input(pol_in, df_lanes)
                        pod = resolve_port_input(pod_in, df_lanes)
                    except Exception as e:
                        messagebox.showerror("POL/POD", str(e)); return
                    code2name = build_code_to_name(df_lanes); pol_name = code2name.get(pol, pol); pod_name = code2name.get(pod, pod)
                    try:
                        # try to get from lanes sheet first
                        rate_per_cbm = None
                        if "RATE_LCL_CBM" in df_lanes.columns:
                            pair = df_lanes[((df_lanes["OPORT_CODE"]==pol) & (df_lanes["DPORT_CODE"]==pod)) | ((df_lanes["OPORT_CODE"]==pod) & (df_lanes["DPORT_CODE"]==pol))]
                            if not pair.empty:
                                import pandas as _pd
                                val = _pd.to_numeric(pair.iloc[0].get("RATE_LCL_CBM"), errors="coerce")
                                if not _pd.isna(val):
                                    rate_per_cbm = float(val)
                        if rate_per_cbm is None:
                            rate_per_cbm = find_lcl_rate_per_cbm(df, pol, pod)
                    except Exception as e:
                        messagebox.showerror("Rates (LCL)", str(e)); return
                    gross_cbm = round(volume_cbm * 1.2, 4)
                    amount = round(gross_cbm * rate_per_cbm, 2)
                    descr = f"Freight (LCL) {pol_name} - {pod_name}"
                    charges_rows.append({
                        "descr": descr,
                        "qty": f"{gross_cbm:.2f} cbm gross",
                        "rate": f"{eur(rate_per_cbm)} / cbm gross",
                        "amount": amount
                    })
                elif mode == "AIR":
                    pol_in = self.ent_pol.get().strip(); pod_in = self.ent_pod.get().strip()
                    if not pol_in or not pod_in:
                        messagebox.showwarning("Input", "Enter both Origin and Destination airport (IATA code or name) or untick Freight."); return
                    # Read AIR sheet (auto if empty)
                    try:
                        air_sheet = AIR_SHEET.strip() if "AIR_SHEET" in globals() else ""
                        if not air_sheet:
                            auto_name = air_auto_sheet(excel)
                            df_air, air_cols = lees_air_lanes(excel, auto_name)
                            messagebox.showinfo("AIR lanes", f"Auto-selected sheet: ‘{auto_name}’.")
                        else:
                            try:
                                df_air, air_cols = lees_air_lanes(excel, air_sheet)
                            except Exception as e:
                                if "Worksheet named" in str(e) or "not found" in str(e).lower():
                                    auto_name = air_auto_sheet(excel)
                                    df_air, air_cols = lees_air_lanes(excel, auto_name)
                                    messagebox.showinfo("AIR lanes", f"Sheet ‘{air_sheet}’ not found. Automatically used ‘{auto_name}’.")
                                else:
                                    raise
                    except Exception as e:
                        messagebox.showerror("Excel (AIR lanes)", f"{e}"); return
                    # Resolve IATA
                    try:
                        pol = resolve_air_input(pol_in, df_air, air_cols)
                        pod = resolve_air_input(pod_in, df_air, air_cols)
                    except Exception as e:
                        messagebox.showerror("AIR IATA", str(e)); return
                    if (pol != pol_in) or (pod != pod_in):
                        messagebox.showinfo("AIR lanes", f"Input converted to codes: ORG={pol_in}→{pol}, DST={pod_in}→{pod}")
                    # ACW calc
                    try:
                        vol = float(self.ent_volume.get().replace(",", "."))
                    except Exception:
                        vol = 0.0
                    acw = max(100.0, math.ceil(max(0.0, vol) * 1.2 * 167.0))
                    try:
                        rates = air_rates_for_lane(df_air, air_cols, pol, pod)
                    except Exception as e:
                        messagebox.showerror("AIR lane not found", str(e)); return
                    brk_key, rate_per_kg = pick_air_rate(rates, acw)
                    amount = round(acw * rate_per_kg, 2)
                    descr = f"Freight (AIR) {pol} - {pod}"
                    charges_rows.append({
                        "descr": descr,
                        "qty": f"{acw:.0f} kg (charg.)",
                        "rate": f"{eur(rate_per_kg)} / kg",
                        "amount": amount
                    })

                elif mode == "FCL":
                    pol_in = self.ent_pol.get().strip(); pod_in = self.ent_pod.get().strip()
                    if not pol_in or not pod_in:
                        messagebox.showwarning("Input", "Enter both POL and POD (code or name) or untick Freight."); return
                    try:
                        if not fcl_sheet_name:
                            auto_name = auto_find_fcl_lanes_sheet(excel)
                            df_lanes = lees_fcl_lanes(excel, auto_name)
                            messagebox.showinfo("FCL lanes", f"Auto-selected sheet: ‘{auto_name}’.")
                        else:
                            try:
                                df_lanes = lees_fcl_lanes(excel, fcl_sheet_name)
                            except Exception as e:
                                if "Worksheet named" in str(e) or "not found" in str(e).lower():
                                    auto_name = auto_find_fcl_lanes_sheet(excel)
                                    df_lanes = lees_fcl_lanes(excel, auto_name)
                                    messagebox.showinfo("FCL lanes", f"Sheet ‘{fcl_sheet_name}’ not found. Automatically used ‘{auto_name}’.")
                                else: raise
                    except Exception as e:
                        messagebox.showerror("Excel read error (FCL lanes)", f"{e}"); return

                    try:
                        pol = resolve_port_input(pol_in, df_lanes)
                        pod = resolve_port_input(pod_in, df_lanes)
                    except Exception as e:
                        messagebox.showerror("POL/POD not recognised", str(e)); return
                    if (pol != pol_in) or (pod != pod_in):
                        messagebox.showinfo("FCL lanes", f"Input converted to codes: POL={pol_in}→{pol}, POD={pod_in}→{pod}")

                    try: rates = fcl_rates_for_lane(df_lanes, pol, pod)
                    except Exception as e:
                        messagebox.showerror("Lane not found", str(e)); return

                    code2name = build_code_to_name(df_lanes); pol_name = code2name.get(pol, pol); pod_name = code2name.get(pod, pod)
                    choice = fcl_choice.upper()
                    if choice.startswith("AUTO"):
                        combo = choose_fcl_combo(volume_cbm, rates)
                        for ctype in ["40HQ","40FT","20FT"]:
                            n = combo.get(ctype, 0)
                            if n > 0:
                                r = rates[ctype]; descr = f"Freight (FCL) {pol_name} - {pod_name}"
                                qty_text = f"{n} × {PRETTY_TYPE.get(ctype, ctype)}"
                                charges_rows.append({"descr": descr, "qty": qty_text, "rate": f"{eur(r)} / container", "amount": n*r})
                    else:
                        if choice not in rates:
                            messagebox.showerror("Container not available", f"{choice} has no rate for this lane."); return
                        cap = FCL_CAPACITY.get(choice, 9999)
                        n = math.ceil(volume_cbm / cap) if volume_cbm > 0 else 1
                        r = rates[choice]; descr = f"Freight (FCL) {pol_name} - {pod_name}"
                        qty_text = f"{n} × {PRETTY_TYPE.get(choice, choice)}"
                        charges_rows.append({"descr": descr, "qty": qty_text, "rate": f"{eur(r)} / container", "amount": n*r})

            if use_dest:
                try:
                    p_dest = LineParams("Destination services", "Destination", dest_addr, mode, volume_cbm, dkm)
                    row_d = match_service_rij_strict_op(df, p_dest) if cc_d != "nl" else match_service_rij(df, p_dest)
                    amt_d = calc_service_prijs(row_d, volume_cbm)
                    descr_d = f"Destination services ({row_d[COLS['mode']]}) {dest_addr}"
                    if str(row_d[COLS["rate_type"]]).strip().upper() == "FLAT":
                        charges_rows.append({"descr": descr_d, "qty": "1 flat", "rate": eur(float(row_d[COLS['flat']])), "amount": amt_d})
                    else:
                        charges_rows.append({"descr": descr_d, "qty": f"{volume_cbm:g} m³", "rate": f"{eur(float(row_d[COLS['rate_per_cbm']]))} / m³", "amount": amt_d})
                except Exception as e:
                    messagebox.showerror("Rates (Destination)", f"{e}"); return

        if not charges_rows:
            messagebox.showwarning("Input", "Nothing to show. Tick at least one service."); return

        
        
        # Build suggested file name: "Cost estimate Voerman <MODE> <Client> LGR 00336"
        try:
            mode_txt = (self.cmb_mode.get() or "").strip().upper()
        except Exception:
            mode_txt = ""
        try:
            client_txt = (self.ent_klant.get() or "").strip()
        except Exception:
            client_txt = ""
        safe_client = re.sub(r"[^A-Za-z0-9 _-]+", "", client_txt)
        iso = datetime.now().isocalendar()
        week_no = iso[1]
        try:
            lgr_no = self._next_lgr_number()
        except Exception:
            lgr_no = 1
        parts = ["Cost estimate Voerman"]
        if mode_txt: parts.append(mode_txt)
        if safe_client: parts.append(safe_client)
        parts.append(f"LGR {lgr_no:03d}{week_no:02d}")
        initial_name = _sanitize_filename(" ".join(parts)) + ".pdf"

        save_path = output_path
        # ensure default output dir exists
        try:
            os.makedirs(OUTPUT_DIR, exist_ok=True)
        except Exception:
            pass
        if not save_path:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                initialfile=initial_name,
                initialdir=OUTPUT_DIR,
                filetypes=[("PDF","*.pdf")],
                title="Save PDF"
            )
        if not save_path: 
            self.set_status("Cancelled."); 
            return
        if logo and not os.path.isfile(logo):
            messagebox.showwarning("Logo not found", f"Logo not found:\n{logo}\nSkipping logo."); logo = None

        # Apply branding (name/address/email/tel) and default logo (no global mutation)
        try:
            _brand = (self.cmb_brand.get() or "Voerman").strip()
            _cfg = BRANDS.get(_brand, BRANDS.get("Voerman", {})).copy()
            brand_name  = _cfg.get("name", BEDRIJFSNAAM)
            brand_addr  = _cfg.get("addr", BEDRIJF_ADRES)
            brand_email = _cfg.get("email", BEDRIJF_EMAIL)
            brand_tel   = _cfg.get("tel", BEDRIJF_TEL)
            if not logo:
                # default from config or autodetect by brand
                if _cfg.get("logo") and os.path.isfile(_cfg["logo"]):
                    logo = _cfg["logo"]
                else:
                    auto = _find_brand_logo_file(_brand)
                    if auto:
                        logo = auto
        except Exception as _e:
            self.fail("Kon branding niet toepassen", _e, title="Branding")

        
        # --- Build client info block for PDF ---
        try:
            ctype = (self.var_client_type.get() or "PRIVATE").upper() if hasattr(self, "var_client_type") else "PRIVATE"
        except Exception:
            ctype = "PRIVATE"
        nm = (getattr(self, "ent_client_name", None).get() if hasattr(self, "ent_client_name") else "") or ""
        asg = (getattr(self, "ent_assignee", None).get() if hasattr(self, "ent_assignee") else "") or ""
        a1  = (getattr(self, "ent_addr1", None).get() if hasattr(self, "ent_addr1") else "") or ""
        a2  = (getattr(self, "ent_addr2", None).get() if hasattr(self, "ent_addr2") else "") or ""
        cp  = (getattr(self, "ent_city_postal", None).get() if hasattr(self, "ent_city_postal") else "") or ""
        co  = (getattr(self, "ent_country", None).get() if hasattr(self, "ent_country") else "") or ""
        em  = (getattr(self, "ent_email", None).get() if hasattr(self, "ent_email") else "") or ""
        lines = []
        if ctype == "AGENT":
            if nm: lines.append(f"Agent: {nm}")
            if asg: lines.append(f"Assignee: {asg}")
        else:
            if nm: lines.append(nm)
        for v in (a1,a2,cp,co):
            if v: lines.append(v)
        if em: lines.append(em)
        klant_block = "\n".join(lines).strip()
        if not klant_block:
            try:
                klant_block = self.ent_klant.get().strip()
            except Exception:
                klant_block = "-"        # Ensure VAT state is up-to-date before building PDF
        try:
            self._update_vat_state()
        except Exception:
            pass

        try:
            maak_pdf_voerman_style(
            save_path=save_path, charges_rows=charges_rows, client_name=klant_block, so_number=ref,
            job_mode=mode, origin_addr=origin_addr if (use_origin or use_freight) else "", dest_addr=dest_addr if (use_dest or use_freight) else "",
            origin_km=okm, origin_km_note=onote, dest_km=dkm, dest_km_note=dnote, volume_cbm=volume_cbm, logo_path=logo, dest_only_rows=dest_only_rows,
            show_rates=bool(getattr(self, 'var_pdf_show_rate', tk.BooleanVar(value=False)).get())
        , show_vat=bool(getattr(self, 'var_pdf_show_vat', tk.BooleanVar(value=False)).get()), vat_rate=float(getattr(self, 'var_vat_rate', tk.DoubleVar(value=21.0)).get() or 0), vat_applies=bool(getattr(self, 'vat_applies', False)), vat_memo=(self.vat_memo if hasattr(self, 'vat_memo') else 'Prices excl. VAT'))
                
        except Exception as e:
            self.fail("Fout bij PDF genereren", e, title="PDF"); return
        self.set_status(f"Done: {save_path}"); messagebox.showinfo("Success", f"PDF created:\n{save_path}")
    def apply_theme(self):
        choice = (self._theme_var.get() or "classic").strip().lower()

        # Classic pure ttk (restore saved theme)
        if choice == "classic":
            try:
                ttk.Style(self).theme_use(self._classic_theme)
            except Exception:
                pass
            try:
                self.set_status("Klassiek thema actief")
            except Exception:
                pass
            return

        # sv-ttk light/dark (no bootstrap)
        if choice in ("sv-ttk-light", "sv-ttk-dark"):
            if SVTTK_AVAILABLE:
                try:
                    # sv-ttk manages theme internally; pick light/dark
                    sv_ttk.set_theme("light" if choice.endswith("light") else "dark")
                    self.set_status(f"Thema toegepast: {choice}")
                    return
                except Exception as e:
                    try:
                        messagebox.showwarning("Thema", f"sv-ttk fout: {e}")
                    except Exception:
                        pass
            else:
                try:
                    messagebox.showwarning("Thema", "sv-ttk niet geïnstalleerd. Installeer met: pip install sv-ttk")
                except Exception:
                    pass
            return

        # Pure ttk: Voerman clam
        if choice == "voerman_clam":
            use_voerman_clam(self)
            try:
                self.set_status("Thema toegepast: voerman_clam")
            except Exception:
                pass
            return

        # From here: ttkbootstrap themes (require bootstrap)
        if not BOOTSTRAP_AVAILABLE or self._tb_style is None:
            try:
                messagebox.showwarning("Thema", "ttkbootstrap niet beschikbaar. Installeer met: pip install ttkbootstrap")
            except Exception:
                pass
            return

        # Apply bootstrap themes
        try:
            self._tb_style.theme_use(choice)
            try:
                self.set_status(f"Thema toegepast: {choice}")
            except Exception:
                pass
        except Exception as e:
            try:
                messagebox.showwarning("Thema", f"Kon thema niet toepassen: {e}")
            except Exception:
                pass

# ===== Robust geocoding overrides (timeout + retries + cache + Nootdorp fallback) =====
_GEOCACHE = {}
_GEOCACHE_FILE = os.path.join(os.path.abspath(os.path.dirname(sys.argv[0] or __file__)), "geocache.json")

def _cache_load():
    global _GEOCACHE
    try:
        if os.path.isfile(_GEOCACHE_FILE):
            with open(_GEOCACHE_FILE, "r", encoding="utf-8") as f:
                _GEOCACHE = json.load(f)
    except Exception:
        _GEOCACHE = {}

def _cache_save():
    try:
        with open(_GEOCACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(_GEOCACHE, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

_cache_load()

class _LocObj:
    def __init__(self, lat, lon, raw=None, address=""):
        self.latitude = float(lat)
        self.longitude = float(lon)
        self.raw = raw or {}
        self.address = address or ""

def geocode_raw(addr: str):
    from geopy.geocoders import Nominatim
    key = (addr or "").strip()
    if not key:
        return None
    hit = _GEOCACHE.get(key)
    if hit and isinstance(hit, dict) and "lat" in hit and "lon" in hit:
        return _LocObj(hit["lat"], hit["lon"], raw=hit.get("raw", {}), address=hit.get("address", key))
    timeout = float(os.getenv("GEOCODE_TIMEOUT", "12"))
    ua = os.getenv("GEOCODE_UA", "voerman_quote_app/4.4 (contact: info@voerman.com)")
    geolocator = Nominatim(user_agent=ua, timeout=timeout)
    last_err = None
    for attempt in range(3):
        try:
            loc = geolocator.geocode(key, addressdetails=True, timeout=timeout)
            if loc:
                raw = getattr(loc, "raw", {})
                _GEOCACHE[key] = {"lat": float(loc.latitude), "lon": float(loc.longitude), "raw": raw, "address": getattr(loc, "address", key)}
                _cache_save()
                return loc
            last_err = RuntimeError("No result")
        except Exception as e:
            last_err = e
        time.sleep(0.5 * (attempt + 1))
    if "nootdorp" in key.lower():
        lat, lon = 52.051, 4.396
        raw = {"address": {"country_code": "nl", "country": "Netherlands"}}
        _GEOCACHE[key] = {"lat": lat, "lon": lon, "raw": raw, "address": "Nootdorp, Netherlands"}
        _cache_save()
        return _LocObj(lat, lon, raw=raw, address="Nootdorp, Netherlands")
    return None

def geocode(addr: str):
    loc = geocode_raw(addr)
    if not loc:
        raise ValueError(f"Kon adres niet geocoderen (timeout of geen resultaat): {addr}")
    return (float(loc.latitude), float(loc.longitude))

def geocode_country(addr: str):
    loc = geocode_raw(addr)
    if not loc:
        raise ValueError(f"Kon land niet bepalen uit adres: {addr}")
    cc = ""
    name = ""
    try:
        addr_dict = loc.raw.get("address", {}) if hasattr(loc, "raw") else {}
        cc = (addr_dict.get("country_code") or "").lower()
        name = addr_dict.get("country") or ""
    except Exception:
        pass
    if not cc and getattr(loc, "address", ""):
        if "Netherlands" in loc.address or "Nederland" in loc.address:
            cc = "nl"; name = "Netherlands"
    return (float(loc.latitude), float(loc.longitude)), cc, name
# ===== End robust geocoding overrides =====


# ===== Clean PDF renderer (no `self` references; uses passed arguments only) =====

    def _num(v, default=0.0):
        try:
            v = v.get()
        except Exception:
            pass
        try:
            if isinstance(v, str):
                s = v.strip().replace('%','').replace(',', '.')
                if s == '':
                    return default
                return float(s)
            return float(v)
        except Exception:
            return default

    def _bool(v):
        try:
            v = v.get()
        except Exception:
            pass
        if isinstance(v, str):
            s = v.strip().lower()
            if s in {'true','yes','on'}: return True
            if s in {'false','no','off'}: return False
            if s.isdigit(): return bool(int(s))
        return bool(v)


    def _val(v):
        try:
            v = v.get()
        except Exception:
            pass
        return '' if v is None else str(v)

    rate = float(vat_rate) if vat_rate is not None else 21.0
    vat_gate = bool(vat_applies)

    LEFT_M = 18*mm; RIGHT_M = 18*mm
    PAGE_W = A4[0]; content_w = PAGE_W - LEFT_M - RIGHT_M

    doc = SimpleDocTemplate(save_path, pagesize=A4, leftMargin=LEFT_M, rightMargin=RIGHT_M, topMargin=16*mm, bottomMargin=16*mm)
    story = []

    img = _logo_flowable(logo_path, max_w_mm=60, max_h_mm=24) if (logo_path and os.path.exists(logo_path)) else None
    company_block = Paragraph(f"<b>{escape(_name)}</b><br/>{escape(_addr)}<br/>{escape(_email)} · {escape(_tel)}", style_small)

    left_stack = []
    if img is not None: left_stack.append([img])
    left_stack.append([company_block])
    left_tbl = Table(left_stack, colWidths=[70*mm])
    left_tbl.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))

    title_tbl = Table([[Paragraph("<b>Cost Estimate</b>", style_title)]], colWidths=[content_w - 70*mm - 10*mm])
    title_tbl.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"RIGHT")]))

    header = Table([[left_tbl, title_tbl]], colWidths=[70*mm, content_w - 70*mm])
    header.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(header); story.append(Spacer(1,6))

    left_par = Paragraph("<b>Account / Partner</b><br/>" + escape(_val(client_name)).replace("\n", "<br/>"), style_norm)
    top_tbl = Table([[left_par, details_right]], colWidths=[content_w - 90*mm, 85*mm])
    top_tbl.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(top_tbl); story.append(Spacer(1,10))

    job_tbl = Table([
        ["Job mode", _val(job_mode), "Volume", f"{float(volume_cbm):.2f} m³"],
        ["Origin", _val(origin_addr) if origin_addr else "-", "Destination", _val(dest_addr) if dest_addr else "-"],
    ], colWidths=[80*mm, 30*mm, 34*mm, 30*mm])
    job_tbl.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.25,colors.lightgrey),
        ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),6),
        ("RIGHTPADDING",(0,0),(-1,-1),6),
    ]))
    job_tbl.hAlign = "LEFT"
    story.append(job_tbl); story.append(Spacer(1,8))

    if show_rates:
        col_widths = [0.55*content_w, 0.15*content_w, 0.15*content_w, 0.15*content_w]
        table_data = [["Charge", "Quantity", "Estimate Rate", "Amount Total"]]
    else:
        col_widths = [0.62*content_w, 0.15*content_w, 0.23*content_w]
        table_data = [["Charge", "Quantity", "Amount Total"]]
    total_sum = 0.0
    for r in charges_rows:
        descr = r.get("descr", "")
        qty = r.get("qty")
        unit_rate = r.get("rate")
        amt = r.get("amount")
        skip = bool(r.get("skip_total", False))
        if show_rates:
            table_data.append([descr, qty if isinstance(qty, str) else (fmt_qty(qty, r.get("qty_unit")) if r.get("qty_unit") else fmt_qty(qty)),
                               "-" if unit_rate is None else unit_rate,
                               "-" if amt  is None else eur(float(amt))])
        else:
            table_data.append([descr, qty if isinstance(qty, str) else (fmt_qty(qty, r.get("qty_unit")) if r.get("qty_unit") else fmt_qty(qty)),
                               "-" if amt  is None else eur(float(amt))])
        if (amt is not None) and not skip:
            total_sum += float(amt)

    if show_rates:
        if not (bool(show_vat) and bool(vat_gate)):
            table_data.append(["", "", Paragraph("<b>Total</b>", style_norm), Paragraph(f"<b>{eur(total_sum)}</b>", style_norm)])
    else:
        if not (bool(show_vat) and bool(vat_gate)):
            table_data.append(["", Paragraph("<b>Total</b>", style_norm), Paragraph(f"<b>{eur(total_sum)}</b>", style_norm)])

    charges_tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    charges_tbl.hAlign = 'LEFT'
    ts = [
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-2),0.25,colors.lightgrey),
        ("BOX",(0,0),(-1,-2),0.25,colors.grey),
    ]
    if show_rates:
        ts += [("ALIGN",(1,1),(1,-2),"RIGHT"), ("ALIGN",(2,1),(2,-2),"RIGHT"), ("ALIGN",(3,1),(3,-1),"RIGHT"), ("RIGHTPADDING",(3,1),(3,-1),6)]
    else:
        ts += [("ALIGN",(1,1),(1,-2),"RIGHT"), ("ALIGN",(2,1),(2,-1),"RIGHT"), ("RIGHTPADDING",(2,1),(2,-1),6)]
    ts += [("LEFTPADDING",(0,0),(-1,-1),6), ("RIGHTPADDING",(0,0),(-1,-1),6)]
    charges_tbl.setStyle(TableStyle(ts))
    story.append(charges_tbl); story.append(Spacer(1,10))

    total2 = 0.0
    if dest_only_rows:
        header = ["Charge (Destination-only)", "Quantity", "Rate"] if show_rates else ["Charge (Destination-only)", "Quantity", "Amount Total"]
        t2 = [header]
        for r in dest_only_rows:
            descr = r.get("descr", "-")
            qty   = r.get("qty", "-")
            rate  = r.get("rate", "-")
            amount = r.get("amount", 0.0)
            third = rate if show_rates else (eur(float(amount)) if isinstance(amount, (int, float,)) else str(amount))
            t2.append([descr, qty, third])
            try:
                total2 += float(amount or 0.0)
            except Exception:
                pass
        col_widths2 = [content_w*0.56, content_w*0.18, content_w*0.26] if show_rates else [content_w*0.62, content_w*0.18, content_w*0.20]
        dest_tbl = Table(t2, colWidths=col_widths2, repeatRows=1)
        dest_tbl.hAlign = 'LEFT'
        ts2 = [
            ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("GRID",(0,0),(-1,-1),0.25,colors.lightgrey),
            ("BOX",(0,0),(-1,-1),0.25,colors.grey),
            ("LEFTPADDING",(0,0),(-1,-1),6), ("RIGHTPADDING",(0,0),(-1,-1),6),
        ]
        ts2 += [("ALIGN",(2,1),(2,-1),"RIGHT"), ("RIGHTPADDING",(2,1),(2,-1),6)]
        dest_tbl.setStyle(TableStyle(ts2))
        story.append(dest_tbl); story.append(Spacer(1,10))

    subtotal = 0.0
    try:
        subtotal = float(total_sum)
    except Exception:
        pass
    try:
        subtotal += float(total2)
    except Exception:
        pass

    payment_txt = vat_memo
    valid_txt = f"Rates valid for {payment_term} from issue date."
    if show_vat and vat_gate and subtotal > 0:
        vat_amount = round(subtotal * (vat_rate_val / 100.0), 2)
        total_inc = round(subtotal + vat_amount, 2)
        vat_data = [
            ["Subtotal", Paragraph(eur(subtotal), style_norm)],
            [f"VAT ({vat_rate_val:.0f}%)", Paragraph(eur(vat_amount), style_norm)],
            [Paragraph("<b>Total incl. VAT</b>", style_norm),
             Paragraph(f"<b>{eur(total_inc)}</b>", style_norm)],
        ]
        amt_col_w = (0.15 if show_rates else 0.23) * content_w
        vat_tbl = Table(vat_data, colWidths=[content_w - amt_col_w, amt_col_w])
        vat_tbl.setStyle(TableStyle([
            ("ALIGN",        (1, 0), (1, -1), "RIGHT"),
            ("LEFTPADDING",  (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING",   (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
            ("LINEABOVE",    (0, 0), (-1, 0), 0.25, colors.HexColor("#BFBFBF")),
            ("LINEABOVE",    (0, 2), (-1, 2), 0.75, colors.black),
            ("BACKGROUND",   (0, 2), (-1, 2), colors.HexColor("#F2F2F2")),
            ("FONTNAME",     (0, 2), (-1, 2), "Helvetica-Bold"),
        ]))
        story.append(Spacer(1, 8))
        story.append(vat_tbl)
        payment_txt = "Prices incl. VAT."

    ul_items = [
        "Transit times and rates are estimates only and may vary due to carrier/supplier changes.",
        valid_txt,
        payment_txt,
    ]
    if dest_only_rows:
        ul_items.append("DTHC, ATHC, NVOCC charges will be billed at cost after the invoice of the forwarder + a prepayment fee of 25 Euro.")
    for b in ul_items:
        story.append(Paragraph("• " + escape(b), style_small2))

    doc.build(story)
# ===== End PDF renderer override =====


    # --- Fallback: simple "Select quotes" dialog (stub) ---
    


    def _merge_pdfs(self, paths, out_path):
        """Merge a list of PDF paths into a single file at out_path. Returns True on success."""
        # Try PyPDF2 first, then pypdf
        Merger = None
        try:
            from PyPDF2 import PdfMerger as _M
            Merger = _M
        except Exception:
            try:
                from pypdf import PdfMerger as _M
                Merger = _M
            except Exception:
                self.fail("Please install PyPDF2 or pypdf to merge PDFs.", title="Merge PDFs")
                return False
        try:
            merger = Merger()
            for p in paths:
                try:
                    merger.append(p)
                except Exception:
                    pass
            with open(out_path, "wb") as f:
                merger.write(f)
            merger.close()
            return True
        except Exception as e:
            try: self.fail(f"Error while merging: {e}", title="Merge PDFs")
            except Exception: pass
            return False

def main():app = App(); app.mainloop()


def maak_pdf_voerman_style(
    save_path, charges_rows: List[dict], client_name=None, so_number=None,
    debtor_number=None, debtor_vat=None, payment_term="14 days", vat_memo="Prices excl. VAT",
    job_mode="-", origin_addr="-", dest_addr="-", origin_km=None, dest_km=None,
    origin_km_note="", dest_km_note="", volume_cbm=0.0, logo_path=None,
    dest_only_rows: List[dict] | None = None,
    show_rates: bool = False, show_vat: bool = False, vat_rate: float | None = None, vat_applies: bool | None = None,
    company_name: str | None = None, company_addr: str | None = None, company_email: str | None = None, company_tel: str | None = None
):
    """
    Generate the PDF. VAT shows ONLY if both are true:
    - show_vat (checkbox)
    - vat_applies (client is Private, not Agent)
    """
    if not REPORTLAB_OK:
        raise RuntimeError("ReportLab is niet geïnstalleerd. Installeer met: pip install reportlab")

    styles = getSampleStyleSheet()
    style_title = styles["Title"]
    style_small = styles["Normal"]; style_small.fontSize = 9; style_small.wordWrap = "CJK"
    style_norm  = styles["Normal"]; style_norm.wordWrap  = "CJK"
    style_small2 = style_small

    def _val(v):
        try:
            v = v.get()
        except Exception:
            pass
        return '' if v is None else str(v)

    def _num(v, default=0.0):
        try:
            v = v.get()
        except Exception:
            pass
        try:
            if isinstance(v, str):
                s = v.strip().replace('%','').replace(',', '.')
                if s == '':
                    return default
                return float(s)
            return float(v)
        except Exception:
            return default

    def _bool(v):
        try:
            v = v.get()
        except Exception:
            pass
        if isinstance(v, str):
            s = v.strip().lower()
            if s in {'true','yes','on'}: return True
            if s in {'false','no','off'}: return False
            if s.isdigit(): return bool(int(s))
        return bool(v)

    # VAT control
    vat_rate_val = _num(vat_rate, 21.0)
    vat_gate = _bool(vat_applies) and _bool(show_vat)

    LEFT_M = 18*mm; RIGHT_M = 18*mm
    PAGE_W = A4[0]; content_w = PAGE_W - LEFT_M - RIGHT_M

    doc = SimpleDocTemplate(save_path, pagesize=A4, leftMargin=LEFT_M, rightMargin=RIGHT_M, topMargin=16*mm, bottomMargin=16*mm)
    story = []

    # Header
    img = _logo_flowable(logo_path, max_w_mm=60, max_h_mm=24) if (logo_path and os.path.exists(logo_path)) else None
    company_block = Paragraph(f"<b>{escape(_name)}</b><br/>{escape(_addr)}<br/>{escape(_email)} · {escape(_tel)}", style_small)

    left_stack = []
    if img is not None: left_stack.append([img])
    left_stack.append([company_block])
    left_tbl = Table(left_stack, colWidths=[70*mm])
    left_tbl.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))

    title_tbl = Table([[Paragraph("<b>Cost Estimate</b>", style_title)]], colWidths=[content_w - 70*mm - 10*mm])
    title_tbl.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"RIGHT")]))

    header = Table([[left_tbl, title_tbl]], colWidths=[70*mm, content_w - 70*mm])
    header.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(header); story.append(Spacer(1,6))

    # Right detail box
    details_right = Table([
        ["SO#", _val(so_number)],
        ["Date", datetime.now().strftime("%Y-%m-%d")],
        ["Debtor number", _val(debtor_number)],
        ["Debtor VAT number", _val(debtor_vat)],
        ["Payment term", _val(payment_term)],
        ["VAT memo", "Prices incl. VAT" if vat_gate else _val(vat_memo)],
    ], colWidths=[40*mm, 45*mm])
    details_right.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.25,colors.lightgrey),
        ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("ALIGN",(0,0),(-1,-1),"LEFT"),
    ]))

    left_par = Paragraph("<b>Account / Partner</b><br/>" + escape(_val(client_name)).replace("\\n", "<br/>"), style_norm)
    top_tbl = Table([[left_par, details_right]], colWidths=[content_w - 90*mm, 85*mm])
    top_tbl.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(top_tbl); story.append(Spacer(1,10))

    # Job summary
    job_tbl = Table([
        ["Job mode", _val(job_mode), "Volume", f"{float(volume_cbm):.2f} m³"],
        ["Origin", _val(origin_addr) if origin_addr else "-", "Destination", _val(dest_addr) if dest_addr else "-"],
    ], colWidths=[80*mm, 30*mm, 34*mm, 30*mm])
    job_tbl.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.25,colors.lightgrey),
        ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),6),
        ("RIGHTPADDING",(0,0),(-1,-1),6),
    ]))
    job_tbl.hAlign = "LEFT"
    story.append(job_tbl); story.append(Spacer(1,8))

    # Charges
    if show_rates:
        col_widths = [0.55*content_w, 0.15*content_w, 0.15*content_w, 0.15*content_w]
        table_data = [["Charge", "Quantity", "Estimate Rate", "Amount Total"]]
    else:
        col_widths = [0.62*content_w, 0.15*content_w, 0.23*content_w]
        table_data = [["Charge", "Quantity", "Amount Total"]]
    total_sum = 0.0
    for r in charges_rows:
        descr = r.get("descr", "")
        qty = r.get("qty")
        unit_rate = r.get("rate")
        amt = r.get("amount")
        skip = bool(r.get("skip_total", False))
        if show_rates:
            table_data.append([descr, qty if isinstance(qty, str) else (fmt_qty(qty, r.get("qty_unit")) if r.get("qty_unit") else fmt_qty(qty)),
                               "-" if unit_rate is None else unit_rate,
                               "-" if amt  is None else eur(float(amt))])
        else:
            table_data.append([descr, qty if isinstance(qty, str) else (fmt_qty(qty, r.get("qty_unit")) if r.get("qty_unit") else fmt_qty(qty)),
                               "-" if amt  is None else eur(float(amt))])
        if (amt is not None) and not skip:
            total_sum += float(amt)

    if show_rates:
        if not vat_gate:
            table_data.append(["", "", Paragraph("<b>Total</b>", style_norm), Paragraph(f"<b>{eur(total_sum)}</b>", style_norm)])
    else:
        if not vat_gate:
            table_data.append(["", Paragraph("<b>Total</b>", style_norm), Paragraph(f"<b>{eur(total_sum)}</b>", style_norm)])

    charges_tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    charges_tbl.hAlign = 'LEFT'
    ts = [
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-2),0.25,colors.lightgrey),
        ("BOX",(0,0),(-1,-2),0.25,colors.grey),
    ]
    if show_rates:
        ts += [("ALIGN",(1,1),(1,-2),"RIGHT"), ("ALIGN",(2,1),(2,-2),"RIGHT"), ("ALIGN",(3,1),(3,-1),"RIGHT"), ("RIGHTPADDING",(3,1),(3,-1),6)]
    else:
        ts += [("ALIGN",(1,1),(1,-2),"RIGHT"), ("ALIGN",(2,1),(2,-1),"RIGHT"), ("RIGHTPADDING",(2,1),(2,-1),6)]
    ts += [("LEFTPADDING",(0,0),(-1,-1),6), ("RIGHTPADDING",(0,0),(-1,-1),6)]
    charges_tbl.setStyle(TableStyle(ts))
    story.append(charges_tbl); story.append(Spacer(1,10))

    # Destination-only charges
    total2 = 0.0
    if dest_only_rows:
        header = ["Charge (Destination-only)", "Quantity", "Rate"] if show_rates else ["Charge (Destination-only)", "Quantity", "Amount Total"]
        t2 = [header]
        for r in dest_only_rows:
            descr = r.get("descr", "-")
            qty   = r.get("qty", "-")
            rate2 = r.get("rate", "-")
            amount = r.get("amount", 0.0)
            third = rate2 if show_rates else (eur(float(amount)) if isinstance(amount, (int, float,)) else str(amount))
            t2.append([descr, qty, third])
            try:
                total2 += float(amount or 0.0)
            except Exception:
                pass
        col_widths2 = [content_w*0.56, content_w*0.18, content_w*0.26] if show_rates else [content_w*0.62, content_w*0.18, content_w*0.20]
        dest_tbl = Table(t2, colWidths=col_widths2, repeatRows=1)
        dest_tbl.hAlign = 'LEFT'
        ts2 = [
            ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("GRID",(0,0),(-1,-1),0.25,colors.lightgrey),
            ("BOX",(0,0),(-1,-1),0.25,colors.grey),
            ("LEFTPADDING",(0,0),(-1,-1),6), ("RIGHTPADDING",(0,0),(-1,-1),6),
        ]
        ts2 += [("ALIGN",(2,1),(2,-1),"RIGHT"), ("RIGHTPADDING",(2,1),(2,-1),6)]
        dest_tbl.setStyle(TableStyle(ts2))
        story.append(dest_tbl); story.append(Spacer(1,10))

    # VAT section
    subtotal = 0.0
    try:
        subtotal = float(total_sum)
    except Exception:
        pass
    try:
        subtotal += float(total2)
    except Exception:
        pass

    if vat_gate and subtotal > 0:
        vat_amount = round(subtotal * (vat_rate_val / 100.0), 2)
        total_inc = round(subtotal + vat_amount, 2)
        vat_data = [
            ["Subtotal", Paragraph(eur(subtotal), style_norm)],
            [f"VAT ({vat_rate_val:.0f}%)", Paragraph(eur(vat_amount), style_norm)],
            [Paragraph("<b>Total incl. VAT</b>", style_norm), Paragraph(f"<b>{eur(total_inc)}</b>", style_norm)],
        ]
        amt_col_w = (0.15 if show_rates else 0.23) * content_w
        vat_tbl = Table(vat_data, colWidths=[content_w - amt_col_w, amt_col_w])
        vat_tbl.setStyle(TableStyle([
            ("ALIGN",        (1, 0), (1, -1), "RIGHT"),
            ("LEFTPADDING",  (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("LINEABOVE",    (0, 0), (-1, 0), 0.25, colors.HexColor("#BFBFBF")),
            ("LINEABOVE",    (0, 2), (-1, 2), 0.75, colors.black),
            ("BACKGROUND",   (0, 2), (-1, 2), colors.HexColor("#F2F2F2")),
            ("FONTNAME",     (0, 2), (-1, 2), "Helvetica-Bold"),
        ]))
        story.append(Spacer(1, 10))
        story.append(vat_tbl)

    # Notes
    ul_items = [
        "Transit times and rates are estimates only and may vary due to carrier/supplier changes.",
        f"Rates valid for {payment_term} from issue date.",
        "Prices incl. VAT." if vat_gate else _val(vat_memo),
    ]
    if dest_only_rows:
        ul_items.append("DTHC, ATHC, NVOCC charges will be billed at cost after the invoice of the forwarder + a prepayment fee of 25 Euro.")
    for b in ul_items:
        story.append(Paragraph("• " + escape(b), style_small2))

    doc.build(story)


if __name__ == '__main__':
    main()
# --- public API for headless integration ---
# --- public API for headless integration (DASHBOARD) ---
def api_generate_pdf(
    *,
    brand: str,
    services,
    mode: str,
    total_cbm: float,
    origin: str,
    destination: str,
    out_path: str,
    charges_rows=None,            # ← optioneel: lijst met regels [{descr, qty, rate, amount}]
    client_name: str = None,      # ← optioneel: voor PDF kop
    show_rates: bool = False,     # ← optioneel: tarievenkolom tonen
    show_vat: bool = False,       # ← optioneel: BTW-blok tonen wanneer van toepassing
    vat_rate: float = 21.0        # ← optioneel
) -> str:
    """
    Headless ingang voor het Voerman dashboard.
    Rendert de échte Voerman-offerte-PDF via maak_pdf_voerman_style(...).
    """
    try:
        # Basisregels wanneer niets is meegegeven (zodat PDF nooit leeg is)
        rows = list(charges_rows or [])
        if not rows:
            # simpele, veilige defaultregels op basis van aangevinkte services
            sset = {str(s).strip().lower() for s in (services or [])}
            if "origin" in sset:
                rows.append({"descr": "Origin services", "qty": "1", "rate": "—", "amount": None})
            if "freight" in sset:
                if str(mode).upper() == "LCL":
                    rows.append({"descr": "Ocean freight (LCL)", "qty": f"{total_cbm:.2f} m³", "rate": "—", "amount": None})
                elif str(mode).upper() == "FCL":
                    rows.append({"descr": "Ocean freight (FCL)", "qty": "1", "rate": "—", "amount": None})
                elif str(mode).upper() == "AIR":
                    rows.append({"descr": "Air freight", "qty": "—", "rate": "—", "amount": None})
                elif str(mode).upper() == "ROAD":
                    rows.append({"descr": "Road freight", "qty": "—", "rate": "—", "amount": None})
            if "destination" in sset:
                rows.append({"descr": "Destination services", "qty": "1", "rate": "—", "amount": None})

        # Branding + logo (zoals je GUI doet)
        _cfg = BRANDS.get(brand or "Voerman", BRANDS.get("Voerman", {})).copy()
        logo_path = _cfg.get("logo") or _find_brand_logo_file(brand or "Voerman")

        # Render met jouw Voerman-layout
        maak_pdf_voerman_style(
            save_path=out_path,
            charges_rows=rows,
            client_name=client_name or "",
            so_number=None,
            payment_term="14 days",
            vat_memo="Prices excl. VAT",
            job_mode=str(mode or "-").upper(),
            origin_addr=origin or "",
            dest_addr=destination or "",
            volume_cbm=float(total_cbm or 0.0),
            logo_path=logo_path,
            dest_only_rows=None,
            show_rates=bool(show_rates),
            show_vat=bool(show_vat),
            vat_rate=float(vat_rate or 21.0),
            vat_applies=True,  # dashboard beslist al of dit van toepassing is; hier aanzetten i.c.m. show_vat
            company_name=_cfg.get("name"),
            company_addr=_cfg.get("addr"),
            company_email=_cfg.get("email"),
            company_tel=_cfg.get("tel"),
        )
    except Exception:
        # laatste redmiddel: lege file zodat dashboard altijd een bijlage heeft
        try:
            open(out_path, "wb").close()
        except Exception:
            pass
    return out_path
