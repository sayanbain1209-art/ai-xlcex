from __future__ import annotations

import csv
import io
import re
from datetime import datetime
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field

app = FastAPI(
    title="SmartSheet AI Backend",
    description="Create formatted Excel workbooks from structured data or CSV/XLSX uploads.",
    version="1.0.0",
)

CATEGORY_KEYWORDS = {
    "Student Information Entry": ["student", "roll", "admission", "class", "school", "guardian"],
    "Employee Data Entry": ["employee", "emp id", "designation", "department", "joining"],
    "Hospital/Patient Records": ["patient", "doctor", "diagnosis", "hospital", "病"],
    "Sales & Purchase": ["sales", "purchase", "quantity", "unit price", "supplier"],
    "Inventory": ["stock", "inventory", "sku", "warehouse", "reorder"],
    "Bank Statement": ["transaction", "account", "withdrawal", "deposit", "balance", "ifsc"],
    "Invoice/Billing": ["invoice", "bill no", "tax", "gst", "subtotal", "total amount"],
    "Survey": ["survey", "response", "feedback", "question", "rating"],
    "Product Listing": ["product", "product name", "brand", "price", "sku"],
    "Customer Database": ["customer", "client", "email", "phone", "address"],
    "Attendance & Payroll": ["salary", "basic salary", "hra", "overtime", "deduction", "attendance", "payroll"],
    "Accounting": ["ledger", "debit", "credit", "journal", "accounting"],
}

class GenerateRequest(BaseModel):
    rows: list[dict[str, Any]] = Field(..., min_length=1)
    sheet_name: str = "SmartSheet Data"
    category: str | None = None

def detect_category(rows: list[dict[str, Any]]) -> str:
    combined = " ".join(str(k) + " " + " ".join(str(v) for v in row.values())
                        for row in rows for k in row.keys()).lower()
    scores = {
        category: sum(1 for keyword in keywords if keyword in combined)
        for category, keywords in CATEGORY_KEYWORDS.items()
    }
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "General Data"

def safe_sheet_name(name: str) -> str:
    cleaned = re.sub(r'[\[\]:*?/\\]', " ", name).strip() or "Sheet1"
    return cleaned[:31]

def as_number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r"[₹,$,\s]", "", str(value))
    cleaned = cleaned.replace(",", "")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

def add_payroll_calculations(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result = [dict(row) for row in rows]
    payroll_words = ("salary", "hra", "allowance", "overtime", "deduction", "payroll")
    keys = {str(k).strip().lower() for row in result for k in row}
    if not any(any(word in key for word in payroll_words) for key in keys):
        return result

    for row in result:
        normalized = {str(k).strip().lower(): k for k in row}
        def get(*names: str) -> float:
            for name in names:
                if name in normalized:
                    return as_number(row[normalized[name]])
            return 0.0

        basic = get("basic salary", "basic")
        hra = get("hra", "house rent allowance")
        allowance = get("other allowance", "allowance", "other allowances")
        overtime = get("overtime", "overtime amount")
        deduction = get("total deduction", "total deductions", "deduction", "deductions")

        # These calculations use zero for missing components. Check source data before relying on totals.
        gross = basic + hra + allowance + overtime
        net = gross - deduction
        row["Gross Salary"] = round(gross, 2)
        row["Net Salary"] = round(net, 2)
    return result

def make_workbook(rows: list[dict[str, Any]], sheet_name: str) -> io.BytesIO:
    if not rows:
        raise ValueError("At least one data row is required.")

    headers: list[str] = []
    for row in rows:
        for key in row.keys():
            header = str(key)
            if header not in headers:
                headers.append(header)

    wb = Workbook()
    ws = wb.active
    ws.title = safe_sheet_name(sheet_name)
    ws.append(headers)

    for row in rows:
        ws.append([row.get(header, "") for header in headers])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 28

    for col_idx, header in enumerate(headers, start=1):
        values = [str(header)]
        values.extend(str(ws.cell(row=r, column=col_idx).value or "")
                      for r in range(2, ws.max_row + 1))
        width = min(max(max(len(v) for v in values) + 3, 12), 42)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 22

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def parse_upload(filename: str, content: bytes) -> list[dict[str, Any]]:
    lower = filename.lower()
    if lower.endswith(".csv"):
        try:
            text = content.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = content.decode("latin-1")
        reader = csv.DictReader(io.StringIO(text))
        return [dict(row) for row in reader if row]
    if lower.endswith(".xlsx"):
        wb = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
        ws = wb.active
        values = list(ws.iter_rows(values_only=True))
        if len(values) < 2:
            raise ValueError("Excel file must contain a header row and at least one data row.")
        headers = [str(v).strip() if v is not None else "" for v in values[0]]
        if not any(headers):
            raise ValueError("Could not find column headings in the first row.")
        return [
            {headers[i]: row[i] if i < len(row) else ""
             for i in range(len(headers)) if headers[i]}
            for row in values[1:] if any(v is not None for v in row)
        ]
    raise ValueError("Supported upload formats are .csv and .xlsx only.")

@app.get("/")
def root():
    return {
        "name": "SmartSheet AI Backend",
        "status": "running",
        "docs": "/docs",
        "supported_uploads": [".csv", ".xlsx"],
    }

@app.get("/health")
def health():
    return {"status": "ok", "time": datetime.utcnow().isoformat() + "Z"}

@app.post("/analyze")
def analyze(request: GenerateRequest):
    category = request.category or detect_category(request.rows)
    return {
        "category": category,
        "row_count": len(request.rows),
        "columns": list(dict.fromkeys(str(k) for row in request.rows for k in row.keys())),
        "payroll_calculations_added": category == "Attendance & Payroll"
            or any("salary" in str(k).lower() for row in request.rows for k in row),
        "note": "Keyword-based category detection; this is not an AI model yet.",
    }

@app.post("/generate-excel")
def generate_excel(request: GenerateRequest):
    category = request.category or detect_category(request.rows)
    rows = add_payroll_calculations(request.rows) if category == "Attendance & Payroll" else request.rows
    try:
        file_obj = make_workbook(rows, request.sheet_name)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    filename = "smartsheet_ai_output.xlsx"
    return StreamingResponse(
        file_obj,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )

@app.post("/upload-and-generate")
async def upload_and_generate(file: UploadFile = File(...)):
    filename = file.filename or "upload"
    content = await file.read()
    try:
        rows = parse_upload(filename, content)
        if not rows:
            raise ValueError("No data rows found in the uploaded file.")
        category = detect_category(rows)
        if category == "Attendance & Payroll":
            rows = add_payroll_calculations(rows)
        file_obj = make_workbook(rows, category)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))

    return StreamingResponse(
        file_obj,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="smartsheet_ai_output.xlsx"'},
    )
