from fastapi import APIRouter
from typing import List
from models_contracts import QuoteRequest, QuoteOption
import storage, pricing_core
router = APIRouter()
@router.post("/quote", response_model=List[QuoteOption])
def quote(req: QuoteRequest):
    qid = storage.new_quote(req.source_id, currency=req.currency)
    options = []
    for mode in (req.modes or ["LCL"]):
        req_single = req.copy(update={"modes":[mode]})
        for opt in pricing_core.generate_quote(req_single):
            opt["mode"] = mode
            storage.add_option(qid, opt); options.append(opt)
    storage.set_quote_status(qid, 'priced'); return options
