from pydantic import BaseModel
from typing import List, Optional

class Place(BaseModel):
    address: Optional[str] = None
    city: Optional[str] = None
    country: Optional[str] = None
    postal_code: Optional[str] = None
    POL: Optional[str] = None
    POD: Optional[str] = None
    IATA: Optional[str] = None

class Measure(BaseModel):
    unit: str
    value: float

class QuoteRequest(BaseModel):
    source_id: str
    modes: List[str] = ["LCL"]
    services: List[str] = []
    origin: Place
    destination: Place
    volumes: List[Measure] = []
    weights: List[Measure] = []
    currency: str = "EUR"
    language: str = "nl"
    options_requested: int = 1
    destination_only: bool = False

class QuoteOption(BaseModel):
    label: str
    buy_total: float
    sell_total: float
    validity: str
    pdf_path: str
    assumptions: List[str] = []
    mode: Optional[str] = None
    service: Optional[str] = None
