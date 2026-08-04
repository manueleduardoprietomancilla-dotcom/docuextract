import os
import base64
import json
import sqlite3
import hashlib
import secrets
from datetime import datetime, timedelta
from fastapi import FastAPI, File, UploadFile, HTTPException, Header, Depends
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional
import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import io

app = FastAPI(title="DocuExtract AI")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory="static"), name="static")

ALLOWED_TYPES = {
    "application/pdf": "pdf",
    "image/jpeg": "jpeg",
    "image/jpg": "jpeg",
    "image/png": "png",
    "image/webp": "webp",
    "image/gif": "gif",
}

PLANS = {
    "starter": {"name": "Starter", "limit": 100, "price_cop": 37000, "link": "https://mpago.la/1aAxegd"},
    "pro": {"name": "Pro", "limit": 300, "price_cop": 78000, "link": "https://mpago.la/1UvoSPg"},
}

EXTRACTION_PROMPT = """Analyze this document and extract all relevant information.

Return ONLY a valid JSON object with this structure (no markdown, no explanation):
{
  "document_type": "invoice|contract|receipt|quote|other",
  "language": "detected language",
  "summary": "brief one-line description of the document",
  "key_fields": {
    "date": "document date if found",
    "due_date": "due date if found",
    "document_number": "invoice/document number if found",
    "total_amount": "total amount with currency",
    "subtotal": "subtotal if found",
    "tax": "tax amount if found",
    "currency": "currency code (USD, EUR, MXN, COP, etc)"
  },
  "parties": {
    "issuer": {
      "name": "company or person name",
      "address": "address if found",
      "tax_id": "tax ID / RFC / NIF / NIT if found",
      "email": "email if found",
      "phone": "phone if found"
    },
    "recipient": {
      "name": "company or person name",
      "address": "address if found",
      "tax_id": "tax ID / RFC / NIF / NIT if found"
    }
  },
  "line_items": [
    {
      "description": "item description",
      "quantity": "quantity",
      "unit_price": "unit price",
      "total": "line total"
    }
  ],
  "additional_info": {
    "payment_method": "payment method if found",
    "payment_terms": "payment terms if found",
    "notes": "any important notes or conditions"
  }
}

If a field is not found, use null. Extract everything visible in the document."""


def get_db():
    db_path = os.environ.get("DB_PATH", "docuextract.db")
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            plan TEXT DEFAULT 'none',
            docs_used INTEGER DEFAULT 0,
            docs_reset_date TEXT,
            token TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        )
    """)
    conn.commit()
    conn.close()


init_db()


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()


def get_user_by_token(token: str):
    conn = get_db()
    user = conn.execute("SELECT * FROM users WHERE token = ?", (token,)).fetchone()
    conn.close()
    return user


def reset_docs_if_needed(user):
    if not user["docs_reset_date"]:
        return
    reset_date = datetime.fromisoformat(user["docs_reset_date"])
    if datetime.now() > reset_date:
        conn = get_db()
        next_reset = (datetime.now() + timedelta(days=30)).isoformat()
        conn.execute("UPDATE users SET docs_used = 0, docs_reset_date = ? WHERE id = ?",
                     (next_reset, user["id"]))
        conn.commit()
        conn.close()


class RegisterRequest(BaseModel):
    email: str
    password: str


class LoginRequest(BaseModel):
    email: str
    password: str


class ActivateRequest(BaseModel):
    token: str
    plan: str


@app.get("/")
async def root():
    return FileResponse("static/index.html")


@app.post("/auth/register")
async def register(req: RegisterRequest):
    conn = get_db()
    existing = conn.execute("SELECT id FROM users WHERE email = ?", (req.email,)).fetchone()
    if existing:
        conn.close()
        raise HTTPException(status_code=400, detail="Email already registered")
    token = secrets.token_hex(32)
    conn.execute(
        "INSERT INTO users (email, password_hash, token) VALUES (?, ?, ?)",
        (req.email, hash_password(req.password), token)
    )
    conn.commit()
    conn.close()
    return {"success": True, "token": token, "email": req.email, "plan": "none", "docs_used": 0}


@app.post("/auth/login")
async def login(req: LoginRequest):
    conn = get_db()
    user = conn.execute(
        "SELECT * FROM users WHERE email = ? AND password_hash = ?",
        (req.email, hash_password(req.password))
    ).fetchone()
    conn.close()
    if not user:
        raise HTTPException(status_code=401, detail="Invalid email or password")
    return {
        "success": True,
        "token": user["token"],
        "email": user["email"],
        "plan": user["plan"],
        "docs_used": user["docs_used"],
        "docs_limit": PLANS.get(user["plan"], {}).get("limit", 0)
    }


@app.get("/auth/me")
async def me(authorization: Optional[str] = Header(None)):
    token = authorization.replace("Bearer ", "") if authorization else None
    if not token:
        raise HTTPException(status_code=401, detail="Not authenticated")
    user = get_user_by_token(token)
    if not user:
        raise HTTPException(status_code=401, detail="Invalid token")
    reset_docs_if_needed(user)
    user = get_user_by_token(token)
    return {
        "email": user["email"],
        "plan": user["plan"],
        "docs_used": user["docs_used"],
        "docs_limit": PLANS.get(user["plan"], {}).get("limit", 0)
    }


@app.post("/auth/activate")
async def activate_plan(req: ActivateRequest):
    if req.plan not in PLANS:
        raise HTTPException(status_code=400, detail="Invalid plan")
    user = get_user_by_token(req.token)
    if not user:
        raise HTTPException(status_code=401, detail="Invalid token")
    conn = get_db()
    next_reset = (datetime.now() + timedelta(days=30)).isoformat()
    conn.execute(
        "UPDATE users SET plan = ?, docs_used = 0, docs_reset_date = ? WHERE token = ?",
        (req.plan, next_reset, req.token)
    )
    conn.commit()
    conn.close()
    return {"success": True, "plan": req.plan}


@app.get("/plans")
async def get_plans():
    return {"plans": PLANS}


@app.post("/extract")
async def extract_document(
    file: UploadFile = File(...),
    authorization: Optional[str] = Header(None),
    x_api_key: Optional[str] = Header(None)
):
    token = authorization.replace("Bearer ", "") if authorization else None
    user = get_user_by_token(token) if token else None

    if user:
        reset_docs_if_needed(user)
        user = get_user_by_token(token)
        if user["plan"] == "none":
            raise HTTPException(status_code=403, detail="Please subscribe to a plan to process documents.")
        plan_limit = PLANS[user["plan"]]["limit"]
        if user["docs_used"] >= plan_limit:
            raise HTTPException(status_code=403, detail=f"Monthly limit reached ({plan_limit} documents). Please upgrade your plan.")
        api_key = os.environ.get("ANTHROPIC_API_KEY")
    else:
        api_key = x_api_key or os.environ.get("ANTHROPIC_API_KEY")

    if not api_key:
        raise HTTPException(status_code=400, detail="API key required.")

    content_type = file.content_type
    if content_type not in ALLOWED_TYPES:
        raise HTTPException(status_code=400, detail="File type not supported. Use PDF, JPG, PNG, WEBP or GIF.")

    file_bytes = await file.read()
    if len(file_bytes) > 10 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="File too large. Maximum size is 10MB.")

    file_b64 = base64.standard_b64encode(file_bytes).decode("utf-8")
    media_type = content_type if content_type != "image/jpg" else "image/jpeg"

    client = anthropic.Anthropic(api_key=api_key)

    if content_type == "application/pdf":
        message = client.messages.create(
            model="claude-sonnet-5",
            max_tokens=4096,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": file_b64}},
                    {"type": "text", "text": EXTRACTION_PROMPT}
                ]
            }]
        )
    else:
        message = client.messages.create(
            model="claude-sonnet-5",
            max_tokens=4096,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": file_b64}},
                    {"type": "text", "text": EXTRACTION_PROMPT}
                ]
            }]
        )

    raw_text = message.content[0].text.strip()
    if raw_text.startswith("```"):
        lines = raw_text.split("\n")
        raw_text = "\n".join(lines[1:-1])

    extracted = json.loads(raw_text)

    if user:
        conn = get_db()
        conn.execute("UPDATE users SET docs_used = docs_used + 1 WHERE token = ?", (token,))
        conn.commit()
        conn.close()

    return {"success": True, "data": extracted, "filename": file.filename}


@app.post("/export-excel")
async def export_excel(data: dict):
    extracted = data.get("data", {})
    filename = data.get("filename", "document")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    section_fill = PatternFill(start_color="16213e", end_color="16213e", fill_type="solid")
    section_font = Font(color="4fc3f7", bold=True, size=10)

    def write_header(row, text):
        cell = ws.cell(row=row, column=1, value=text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        ws.merge_cells(f"A{row}:D{row}")

    def write_section(row, text):
        cell = ws.cell(row=row, column=1, value=text)
        cell.fill = section_fill
        cell.font = section_font
        ws.merge_cells(f"A{row}:D{row}")

    def write_row(row, label, value):
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=2, value=str(value) if value else "")

    current_row = 1
    write_header(current_row, f"DocuExtract AI - {filename}")
    current_row += 2

    write_section(current_row, "DOCUMENT INFO")
    current_row += 1
    write_row(current_row, "Document Type", extracted.get("document_type", ""))
    current_row += 1
    write_row(current_row, "Language", extracted.get("language", ""))
    current_row += 1
    write_row(current_row, "Summary", extracted.get("summary", ""))
    current_row += 2

    key_fields = extracted.get("key_fields", {})
    if key_fields:
        write_section(current_row, "KEY FIELDS")
        current_row += 1
        for k, v in key_fields.items():
            if v:
                write_row(current_row, k.replace("_", " ").title(), v)
                current_row += 1
        current_row += 1

    parties = extracted.get("parties", {})
    issuer = parties.get("issuer", {})
    recipient = parties.get("recipient", {})

    if any(issuer.values()):
        write_section(current_row, "ISSUER / FROM")
        current_row += 1
        for k, v in issuer.items():
            if v:
                write_row(current_row, k.replace("_", " ").title(), v)
                current_row += 1
        current_row += 1

    if any(v for v in recipient.values() if v):
        write_section(current_row, "RECIPIENT / TO")
        current_row += 1
        for k, v in recipient.items():
            if v:
                write_row(current_row, k.replace("_", " ").title(), v)
                current_row += 1
        current_row += 1

    line_items = extracted.get("line_items", [])
    if line_items:
        write_section(current_row, "LINE ITEMS")
        current_row += 1
        headers = ["Description", "Quantity", "Unit Price", "Total"]
        for col, h in enumerate(headers, 1):
            ws.cell(row=current_row, column=col, value=h).font = Font(bold=True)
        current_row += 1
        for item in line_items:
            ws.cell(row=current_row, column=1, value=item.get("description", ""))
            ws.cell(row=current_row, column=2, value=item.get("quantity", ""))
            ws.cell(row=current_row, column=3, value=item.get("unit_price", ""))
            ws.cell(row=current_row, column=4, value=item.get("total", ""))
            current_row += 1
        current_row += 1

    additional = extracted.get("additional_info", {})
    if any(v for v in additional.values() if v):
        write_section(current_row, "ADDITIONAL INFO")
        current_row += 1
        for k, v in additional.items():
            if v:
                write_row(current_row, k.replace("_", " ").title(), v)
                current_row += 1

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    safe_name = filename.rsplit(".", 1)[0] if "." in filename else filename
    return StreamingResponse(
        excel_buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={safe_name}_extracted.xlsx"},
    )
