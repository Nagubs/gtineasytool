from fastapi import FastAPI
from fastapi.responses import JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional
import json
import os
import openpyxl

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ✅ Serve frontend at /static
app.mount("/static", StaticFiles(directory="static", html=True), name="static")

STATE_FILE = "state.json"
DATA_FILE = "generated_codes.json"
EXCEL_FILE = "generated_codes.xlsx"

class GenerateRequest(BaseModel):
    code_type: str
    prefix: str
    sku: str
    indicator: Optional[str] = None

def calculate_check_digit(number: str) -> str:
    digits = list(map(int, number))
    odd_sum = sum(digits[-1::-2])
    even_sum = sum(digits[-2::-2])
    total = (odd_sum * 3) + even_sum
    return str((10 - (total % 10)) % 10)

def get_next_reference() -> str:
    if not os.path.exists(STATE_FILE):
        with open(STATE_FILE, "w") as f:
            json.dump({"item_reference": 10000}, f)
    with open(STATE_FILE, "r") as f:
        state = json.load(f)
    ref = state.get("item_reference", 10000)
    state["item_reference"] = ref + 1
    with open(STATE_FILE, "w") as f:
        json.dump(state, f)
    return str(ref)

def save_code(record):
    data = []
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as f:
            data = json.load(f)
    data.append(record)
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=2)

def export_history_to_excel(filename=EXCEL_FILE, records=None):
    if records is None:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r") as f:
                records = json.load(f)
        else:
            records = []

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Code Type", "Prefix", "SKU", "Indicator", "Item Reference", "Generated Code"])

    for record in records:
        sheet.append([
            record.get("code_type"),
            record.get("prefix"),
            record.get("sku"),
            record.get("indicator") or "",
            record.get("item_reference"),
            record.get("generated_code"),
        ])

    workbook.save(filename)

@app.post("/generate")
async def generate_code(req: GenerateRequest):
    try:
        # ✅ Prevent duplicate SKUs
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r") as f:
                existing_data = json.load(f)
            for record in existing_data:
                if record.get("sku") == req.sku:
                    return JSONResponse(
                        status_code=400,
                        content={"error": f"SKU '{req.sku}' already has GTIN: {record.get('generated_code')}"}
                    )

        item_reference = get_next_reference()
        indicator = req.indicator if req.code_type == "GTIN-14" and req.indicator else "0"

        if req.code_type == "GTIN-13":
            base = req.prefix + item_reference
            base = base[:12].ljust(12, '0')
            code = base + calculate_check_digit(base)

        elif req.code_type == "GTIN-14":
            base = indicator + req.prefix + item_reference
            base = base[:13].rjust(13, '0')
            code = base + calculate_check_digit(base)

        elif req.code_type == "GMN":
            code = (req.prefix + item_reference).ljust(18, '0')

        elif req.code_type == "UDI-DI":
            code = (req.prefix + item_reference).ljust(14, '0')

        else:
            return JSONResponse(status_code=400, content={"error": "Invalid code_type."})

        record = {
            "code_type": req.code_type,
            "prefix": req.prefix,
            "sku": req.sku,
            "indicator": indicator if req.code_type == "GTIN-14" else None,
            "item_reference": item_reference,
            "generated_code": code
        }

        save_code(record)
        return {"generated_code": code}

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@app.get("/history")
async def get_history():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as f:
            data = json.load(f)
    else:
        data = []
    return {"history": data}

@app.get("/export-excel")
async def export_excel():
    export_history_to_excel()
    return FileResponse(EXCEL_FILE, filename=EXCEL_FILE)

@app.get("/export-excel/{sku}")
async def export_excel_for_sku(sku: str):
    if not os.path.exists(DATA_FILE):
        return JSONResponse(status_code=404, content={"error": "No data available."})
    with open(DATA_FILE, "r") as f:
        data = json.load(f)
    filtered = [rec for rec in data if rec.get("sku") == sku]
    if not filtered:
        return JSONResponse(status_code=404, content={"error": "SKU not found."})
    filename = f"report_{sku}.xlsx"
    export_history_to_excel(filename, filtered)
    return FileResponse(filename, filename=filename)
