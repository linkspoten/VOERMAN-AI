# extractor.py
import re
from typing import Any, Dict, List
from models_contracts import QuoteRequest, Place, Measure

VAN_NAAR = re.compile(r'\bvan\s+([A-Za-zÀ-ÿ .\'-]{2,50})\s+naar\s+([A-Za-zÀ-ÿ .\'-]{2,50})', re.I)

def _clean_city(s: str) -> str:
    s = s.strip().strip(".:,;")
    parts = [p for p in re.split(r'\s+', s) if p]
    return " ".join(parts[:3])

def _detect_units(text: str):
    CF_TO_M3=0.0283168; LB_TO_KG=0.453592
    vols=[]; wts=[]
    for m in re.finditer(r'(\d+(?:[.,]\d+)?)\s*(m3|cbm|cf|kg|lb)\b', text or '', re.I):
        val=float(m.group(1).replace(',', '.')); unit=m.group(2).lower()
        if unit in ('m3','cbm','cf'):
            if unit=='cf': val*=CF_TO_M3
            vols.append({'unit':'m3','value':round(val,2)})
        elif unit in ('kg','lb'):
            if unit=='lb': val*=LB_TO_KG
            wts.append({'unit':'kg','value':round(val,1)})
    return vols, wts

def _detect_route(text: str):
    m = VAN_NAAR.search(text or '')
    if not m:
        return Place(city='Amsterdam', country='NL'), Place(city='Montreal', country='CA')
    return Place(city=_clean_city(m.group(1))), Place(city=_clean_city(m.group(2)))

def _detect_mode(text: str) -> List[str]:
    t = (text or '').lower()
    if any(k in t for k in ['air','luchtvracht']): return ['AIR']
    if any(k in t for k in ['fcl']): return ['FCL']
    if any(k in t for k in ['groupage','lcl']): return ['LCL']
    if any(k in t for k in ['road','truck','weg','groupage road']): return ['ROAD']
    return ['LCL']

def _detect_services(text: str) -> List[str]:
    t = (text or '').lower()
    s = set()
    if any(k in t for k in ['origin','ophaal','afhalen','stairs','inpack']): s.add('origin')
    if any(k in t for k in ['freight','vracht','zee','lucht']): s.add('freight')
    if any(k in t for k in ['destination','levering','uitpak']): s.add('destination')
    # sensible default if nothing stated
    if not s: s.update(['origin','freight','destination'])
    return list(s)

def extract_from_unified(msg: Dict[str, Any]):
    body = msg.get('body','') or ''
    vols, wts = _detect_units(body)
    origin, dest = _detect_route(body)
    modes = _detect_mode(body)
    services = _detect_services(body)
    q = []
    if not vols: q.append({'key':'volumes[0]','question':'Wat is het volume in m³?'})
    if not wts and modes and modes[0]=='AIR': q.append({'key':'weights[0]','question':'Wat is het gewicht in kg?'})
    req = QuoteRequest(source_id=msg['id'], modes=modes, origin=origin, destination=dest,
                       volumes=[Measure(**v) for v in vols], weights=[Measure(**w) for w in wts],
                       language=msg.get('language','nl'), services=services)
    class Res: pass
    res = Res(); res.request = req; res.clarifying_questions = q
    return res
