import base64
import io
import json
import os
import re
import secrets
import imaplib
import email as email_lib
import zipfile
from datetime import datetime
from functools import wraps
from pathlib import Path
from typing import Optional

from cryptography.fernet import Fernet
from docx import Document
from flask import (
    Flask, flash, jsonify, redirect, render_template, request,
    send_file, session, url_for
)
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pypdf import PdfReader
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

try:
    from google import genai
    from google.genai import types
except Exception:
    genai = None
    types = None

try:
    from apscheduler.schedulers.background import BackgroundScheduler
except Exception:
    BackgroundScheduler = None


BASE_DIR = Path(__file__).resolve().parent
app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "change-this-secret-in-render")
database_url = os.getenv("DATABASE_URL", f"sqlite:///{BASE_DIR / 'svai.db'}")
if database_url.startswith("postgres://"):
    database_url = database_url.replace("postgres://", "postgresql://", 1)
app.config["SQLALCHEMY_DATABASE_URI"] = database_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "50")) * 1024 * 1024
db = SQLAlchemy(app)

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash").strip()
AI_CLIENT = genai.Client(api_key=GEMINI_API_KEY) if (genai and GEMINI_API_KEY) else None

BANK_KEYWORDS = (
    "valuation", "technical", "property", "lap", "home loan", "housing loan",
    "initiate", "inspection", "applicant", "application", "case id", "customer"
)

DOCUMENT_EXTENSIONS = {
    ".pdf", ".jpg", ".jpeg", ".png", ".webp", ".docx", ".xlsx", ".xls"
}
PHOTO_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".heic"}
TEMPLATE_EXTENSIONS = {".xlsx", ".xlsm"}

PHOTO_CATEGORIES = {
    "front": "Front Elevation",
    "elevation": "Front Elevation",
    "approach": "Approach Road",
    "road": "Approach Road",
    "kitchen": "Kitchen",
    "hall": "Hall / Drawing Room",
    "bed": "Bedroom",
    "room": "Internal Room",
    "meter": "Electricity Meter",
    "selfie": "Property Selfie",
    "side": "Side View",
    "back": "Rear View",
    "toilet": "Toilet / Bathroom",
    "bath": "Toilet / Bathroom",
    "terrace": "Terrace",
    "map": "Location / Map",
    "sketch": "Site Sketch",
}


class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(180), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    name = db.Column(db.String(180), nullable=False, default="SVAI User")
    role = db.Column(db.String(30), nullable=False, default="admin")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class EmailAccount(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(180), unique=True, nullable=False)
    encrypted_password = db.Column(db.Text, nullable=False)
    provider = db.Column(db.String(30), nullable=False, default="auto")
    imap_host = db.Column(db.String(180))
    bank_name = db.Column(db.String(180))
    active = db.Column(db.Boolean, default=True)
    last_fetch_at = db.Column(db.DateTime)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class ValuationCase(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    application_number = db.Column(db.String(180), index=True)
    customer_name = db.Column(db.String(220), index=True)
    contact_number = db.Column(db.String(40))
    property_address = db.Column(db.Text)
    bank_name = db.Column(db.String(180), index=True)
    branch_name = db.Column(db.String(180))
    case_type = db.Column(db.String(100), default="LAP")
    source_email = db.Column(db.String(180))
    source_message_id = db.Column(db.String(255), unique=True)
    email_subject = db.Column(db.Text)
    email_received_at = db.Column(db.DateTime)
    visit_by = db.Column(db.String(180))
    status = db.Column(db.String(80), default="New")
    archived = db.Column(db.Boolean, default=False)
    extracted_json = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class FileAsset(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    case_id = db.Column(db.Integer, db.ForeignKey("valuation_case.id"), nullable=True, index=True)
    asset_type = db.Column(db.String(40), nullable=False)  # document/photo/template/report
    category = db.Column(db.String(120))
    filename = db.Column(db.String(255), nullable=False)
    mime_type = db.Column(db.String(120))
    content = db.Column(db.LargeBinary, nullable=False)
    extracted_text = db.Column(db.Text)
    extraction_json = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Valuation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    case_id = db.Column(db.Integer, db.ForeignKey("valuation_case.id"), unique=True, nullable=False)
    land_area = db.Column(db.Float, default=0)
    land_rate = db.Column(db.Float, default=0)
    builtup_area = db.Column(db.Float, default=0)
    construction_rate = db.Column(db.Float, default=0)
    age_years = db.Column(db.Float, default=0)
    depreciation_percent = db.Column(db.Float, default=0)
    govt_land_rate = db.Column(db.Float, default=0)
    govt_construction_rate = db.Column(db.Float, default=0)
    conservative_percent = db.Column(db.Float, default=80)
    distress_percent = db.Column(db.Float, default=70)
    land_value = db.Column(db.Float, default=0)
    gross_building_value = db.Column(db.Float, default=0)
    depreciation_amount = db.Column(db.Float, default=0)
    net_building_value = db.Column(db.Float, default=0)
    market_value = db.Column(db.Float, default=0)
    conservative_value = db.Column(db.Float, default=0)
    distress_value = db.Column(db.Float, default=0)
    govt_value = db.Column(db.Float, default=0)
    remarks = db.Column(db.Text)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


def encryption_key() -> bytes:
    configured = os.getenv("ENCRYPTION_KEY", "").strip()
    if configured:
        return configured.encode()
    raw = app.config["SECRET_KEY"].encode()
    return base64.urlsafe_b64encode(__import__("hashlib").sha256(raw).digest())


FERNET = Fernet(encryption_key())


def encrypt_password(value: str) -> str:
    return FERNET.encrypt(value.encode()).decode()


def decrypt_password(value: str) -> str:
    return FERNET.decrypt(value.encode()).decode()


def login_required(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapped


def api_login_required(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            return jsonify({"success": False, "message": "Login required"}), 401
        return fn(*args, **kwargs)
    return wrapped


def safe_json(text: Optional[str], default=None):
    if not text:
        return default if default is not None else {}
    try:
        return json.loads(text)
    except Exception:
        return default if default is not None else {}


def detect_imap(email_address: str, provider: str = "auto", custom_host: str = ""):
    if custom_host:
        return custom_host
    domain = email_address.lower().split("@")[-1]
    provider = (provider or "auto").lower()
    if "gmail" in provider or domain in {"gmail.com", "googlemail.com"}:
        return "imap.gmail.com"
    if "yahoo" in provider or "yahoo" in domain:
        return "imap.mail.yahoo.com"
    if "outlook" in provider or domain in {"outlook.com", "hotmail.com", "live.com"}:
        return "outlook.office365.com"
    return f"imap.{domain}"


def decode_header_value(value: str) -> str:
    if not value:
        return ""
    parts = []
    for piece, charset in email_lib.header.decode_header(value):
        if isinstance(piece, bytes):
            parts.append(piece.decode(charset or "utf-8", errors="ignore"))
        else:
            parts.append(str(piece))
    return "".join(parts)


def email_body(message) -> str:
    chunks = []
    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition", ""))
            if content_type in ("text/plain", "text/html") and "attachment" not in disposition:
                try:
                    payload = part.get_payload(decode=True) or b""
                    text = payload.decode(part.get_content_charset() or "utf-8", errors="ignore")
                    if content_type == "text/html":
                        text = re.sub(r"<[^>]+>", " ", text)
                    chunks.append(text)
                except Exception:
                    continue
    else:
        payload = message.get_payload(decode=True) or b""
        chunks.append(payload.decode(message.get_content_charset() or "utf-8", errors="ignore"))
    return re.sub(r"\s+", " ", " ".join(chunks)).strip()[:25000]


def regex_extract(subject: str, body: str, sender: str):
    text = f"{subject}\n{body}"
    patterns = {
        "application_number": [
            r"(?i)(?:application|app|case|lead|proposal|loan)\s*(?:no|number|id|#)?\s*[:\-]\s*([A-Z0-9\/\-]{5,30})",
            r"(?i)\b(?:LAN|APPL|APP)[\s:\-]*([A-Z0-9\/\-]{5,30})",
        ],
        "customer_name": [
            r"(?i)(?:customer|applicant|borrower)\s*(?:name)?\s*[:\-]\s*([A-Za-z][A-Za-z .]{2,60})",
            r"(?i)technical\s+report\s+(?:initiate|initiation)\s+(?:of)?\s*[:\-]\s*([A-Za-z][A-Za-z .]{2,60})",
        ],
        "contact_number": [r"(?<!\d)([6-9]\d{9})(?!\d)"],
        "property_address": [
            r"(?i)(?:property|site|collateral)\s*address\s*[:\-]\s*(.{10,300}?)(?:\bmobile\b|\bcontact\b|\bbranch\b|\bbank\b|$)",
            r"(?i)address\s*[:\-]\s*(.{10,300}?)(?:\bmobile\b|\bcontact\b|\bbranch\b|\bbank\b|$)",
        ],
        "branch_name": [r"(?i)branch\s*(?:name)?\s*[:\-]\s*([A-Za-z0-9 .\-]{2,80})"],
        "case_type": [r"(?i)(?:case|loan)\s*type\s*[:\-]\s*([A-Za-z0-9 .\-]{2,50})"],
    }
    result = {}
    for key, options in patterns.items():
        for pattern in options:
            match = re.search(pattern, text)
            if match:
                result[key] = match.group(1).strip(" .,-")
                break

    sender_domain = sender.lower().split("@")[-1].split(">")[0] if "@" in sender else ""
    bank_guess = sender_domain.split(".")[0].replace("-", " ").title() if sender_domain else ""
    result["bank_name"] = bank_guess
    return result


def ai_extract_email(subject: str, body: str, sender: str):
    fallback = regex_extract(subject, body, sender)
    if not AI_CLIENT:
        return fallback
    prompt = f"""
Extract only genuine bank property valuation case data from the email below.
Do not invent any value. Missing fields must be empty strings.
Return valid JSON with keys:
application_number, customer_name, contact_number, property_address,
bank_name, branch_name, case_type.
Sender: {sender}
Subject: {subject}
Body: {body}
"""
    try:
        response = AI_CLIENT.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config={"response_mime_type": "application/json"},
        )
        parsed = json.loads(response.text)
        for key, value in fallback.items():
            if not parsed.get(key):
                parsed[key] = value
        return parsed
    except Exception:
        return fallback


def is_valuation_email(subject: str, body: str) -> bool:
    text = f"{subject} {body}".lower()
    score = sum(1 for word in BANK_KEYWORDS if word in text)
    return score >= 2 or ("valuation" in text and "property" in text)


def fetch_email_account(account: EmailAccount):
    created = 0
    ignored = 0
    host = detect_imap(account.email, account.provider, account.imap_host or "")
    password = decrypt_password(account.encrypted_password)
    mail = imaplib.IMAP4_SSL(host, 993)
    try:
        mail.login(account.email, password)
        mail.select("INBOX")
        status, payload = mail.search(None, "UNSEEN")
        if status != "OK":
            return {"created": 0, "ignored": 0, "message": "Inbox search failed"}
        ids = payload[0].split()[-100:]
        for msg_id in ids:
            status, msg_data = mail.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue
            raw = next((x[1] for x in msg_data if isinstance(x, tuple)), None)
            if not raw:
                continue
            message = email_lib.message_from_bytes(raw)
            subject = decode_header_value(message.get("Subject", ""))
            sender = decode_header_value(message.get("From", ""))
            body = email_body(message)
            unique_id = message.get("Message-ID") or f"{account.email}:{msg_id.decode()}"
            if ValuationCase.query.filter_by(source_message_id=unique_id).first():
                continue
            if not is_valuation_email(subject, body):
                ignored += 1
                continue

            details = ai_extract_email(subject, body, sender)
            received = None
            try:
                received = email_lib.utils.parsedate_to_datetime(message.get("Date")).replace(tzinfo=None)
            except Exception:
                received = datetime.utcnow()

            case = ValuationCase(
                application_number=details.get("application_number", ""),
                customer_name=details.get("customer_name", "") or "To be reviewed",
                contact_number=details.get("contact_number", ""),
                property_address=details.get("property_address", ""),
                bank_name=details.get("bank_name", "") or account.bank_name or "",
                branch_name=details.get("branch_name", ""),
                case_type=details.get("case_type", "") or "LAP",
                source_email=account.email,
                source_message_id=unique_id,
                email_subject=subject,
                email_received_at=received,
                status="New - Email",
                extracted_json=json.dumps(details, ensure_ascii=False),
            )
            db.session.add(case)
            db.session.commit()
            created += 1
        account.last_fetch_at = datetime.utcnow()
        db.session.commit()
        return {"created": created, "ignored": ignored, "message": "Fetch completed"}
    finally:
        try:
            mail.logout()
        except Exception:
            pass


def classify_photo(filename: str) -> str:
    name = Path(filename).stem.lower()
    for key, category in PHOTO_CATEGORIES.items():
        if key in name:
            return category
    return "Other Site Photo"


def extract_basic_text(filename: str, content: bytes) -> str:
    ext = Path(filename).suffix.lower()
    try:
        if ext == ".pdf":
            reader = PdfReader(io.BytesIO(content))
            return "\n".join((page.extract_text() or "") for page in reader.pages)[:50000]
        if ext == ".docx":
            document = Document(io.BytesIO(content))
            return "\n".join(p.text for p in document.paragraphs)[:50000]
        if ext in {".xlsx", ".xlsm"}:
            wb = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
            output = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    line = " | ".join(str(v) for v in row if v not in (None, ""))
                    if line:
                        output.append(line)
            return "\n".join(output)[:50000]
    except Exception:
        return ""
    return ""


def ai_extract_document(filename: str, content: bytes, existing_text: str = ""):
    schema_keys = [
        "document_type", "owner_name", "property_address", "village", "tehsil",
        "district", "state", "survey_khasra_plot_no", "land_area",
        "builtup_area", "boundaries", "registration_number", "registration_date"
    ]
    empty = {key: "" for key in schema_keys}
    if not AI_CLIENT:
        return empty
    prompt = """
Read this Indian property document for a valuation report.
Do not invent information. Return JSON only with:
document_type, owner_name, property_address, village, tehsil, district, state,
survey_khasra_plot_no, land_area, builtup_area, boundaries,
registration_number, registration_date.
Use empty string for missing information.
"""
    try:
        ext = Path(filename).suffix.lower()
        parts = [prompt]
        if existing_text:
            parts.append(existing_text[:30000])
        elif ext in {".jpg", ".jpeg", ".png", ".webp", ".pdf"} and types:
            mime = {
                ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
                ".webp": "image/webp", ".pdf": "application/pdf"
            }[ext]
            parts.append(types.Part.from_bytes(data=content, mime_type=mime))
        response = AI_CLIENT.models.generate_content(
            model=GEMINI_MODEL,
            contents=parts,
            config={"response_mime_type": "application/json"},
        )
        parsed = json.loads(response.text)
        return {key: parsed.get(key, "") for key in schema_keys}
    except Exception:
        return empty


def store_asset(case_id, asset_type, filename, content, mime_type=None, category=None):
    filename = secure_filename(filename) or f"file_{secrets.token_hex(4)}"
    text = ""
    extraction = {}
    if asset_type == "document":
        text = extract_basic_text(filename, content)
        extraction = ai_extract_document(filename, content, text)
    elif asset_type == "photo":
        category = category or classify_photo(filename)
    asset = FileAsset(
        case_id=case_id,
        asset_type=asset_type,
        category=category,
        filename=filename,
        mime_type=mime_type or "application/octet-stream",
        content=content,
        extracted_text=text,
        extraction_json=json.dumps(extraction, ensure_ascii=False) if extraction else None,
    )
    db.session.add(asset)
    db.session.commit()
    return asset


def safe_zip_members(zf):
    for member in zf.infolist():
        path = Path(member.filename)
        if member.is_dir() or ".." in path.parts or path.is_absolute():
            continue
        yield member


def valuation_calculation(form):
    def f(name, default=0):
        try:
            return float(form.get(name, default) or default)
        except Exception:
            return float(default)

    data = {
        "land_area": f("land_area"),
        "land_rate": f("land_rate"),
        "builtup_area": f("builtup_area"),
        "construction_rate": f("construction_rate"),
        "age_years": f("age_years"),
        "depreciation_percent": f("depreciation_percent"),
        "govt_land_rate": f("govt_land_rate"),
        "govt_construction_rate": f("govt_construction_rate"),
        "conservative_percent": f("conservative_percent", 80),
        "distress_percent": f("distress_percent", 70),
        "remarks": form.get("remarks", ""),
    }
    data["land_value"] = data["land_area"] * data["land_rate"]
    data["gross_building_value"] = data["builtup_area"] * data["construction_rate"]
    data["depreciation_amount"] = data["gross_building_value"] * data["depreciation_percent"] / 100
    data["net_building_value"] = data["gross_building_value"] - data["depreciation_amount"]
    data["market_value"] = data["land_value"] + data["net_building_value"]
    data["conservative_value"] = data["market_value"] * data["conservative_percent"] / 100
    data["distress_value"] = data["market_value"] * data["distress_percent"] / 100
    data["govt_value"] = (
        data["land_area"] * data["govt_land_rate"] +
        data["builtup_area"] * data["govt_construction_rate"]
    )
    return data


def report_mapping(case: ValuationCase, valuation: Valuation):
    extracted = safe_json(case.extracted_json)
    values = {
        "CASE_ID": case.id,
        "APPLICATION_NUMBER": case.application_number or "",
        "CUSTOMER_NAME": case.customer_name or "",
        "CONTACT_NUMBER": case.contact_number or "",
        "PROPERTY_ADDRESS": case.property_address or "",
        "BANK_NAME": case.bank_name or "",
        "BRANCH_NAME": case.branch_name or "",
        "CASE_TYPE": case.case_type or "",
        "STATUS": case.status or "",
        "VISIT_BY": case.visit_by or "",
        "REPORT_DATE": datetime.now().strftime("%d-%m-%Y"),
        "LAND_AREA": valuation.land_area,
        "LAND_RATE": valuation.land_rate,
        "LAND_VALUE": valuation.land_value,
        "BUILTUP_AREA": valuation.builtup_area,
        "CONSTRUCTION_RATE": valuation.construction_rate,
        "GROSS_BUILDING_VALUE": valuation.gross_building_value,
        "DEPRECIATION_PERCENT": valuation.depreciation_percent,
        "DEPRECIATION_AMOUNT": valuation.depreciation_amount,
        "NET_BUILDING_VALUE": valuation.net_building_value,
        "MARKET_VALUE": valuation.market_value,
        "CONSERVATIVE_VALUE": valuation.conservative_value,
        "DISTRESS_VALUE": valuation.distress_value,
        "GOVT_VALUE": valuation.govt_value,
        "REMARKS": valuation.remarks or "",
    }
    for key, value in extracted.items():
        values[f"EXTRACTED_{str(key).upper()}"] = value
    return values


def fill_template(content: bytes, mapping: dict) -> bytes:
    keep_vba = False
    wb = load_workbook(io.BytesIO(content), keep_vba=keep_vba)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    value = cell.value
                    for key, replacement in mapping.items():
                        value = value.replace(f"{{{{{key}}}}}", str(replacement if replacement is not None else ""))
                    cell.value = value
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def generic_report(case: ValuationCase, valuation: Valuation, assets):
    wb = Workbook()
    ws = wb.active
    ws.title = "Valuation Report"
    ws.merge_cells("A1:D1")
    ws["A1"] = "SVAI - SAKSHAM ASSOCIATE PROPERTY VALUATION REPORT"
    ws["A1"].font = Font(size=15, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    mapping = report_mapping(case, valuation)
    ordered = [
        ("Application Number", mapping["APPLICATION_NUMBER"], "Bank", mapping["BANK_NAME"]),
        ("Customer Name", mapping["CUSTOMER_NAME"], "Branch", mapping["BRANCH_NAME"]),
        ("Contact Number", mapping["CONTACT_NUMBER"], "Case Type", mapping["CASE_TYPE"]),
        ("Property Address", mapping["PROPERTY_ADDRESS"], "Status", mapping["STATUS"]),
        ("Land Area", mapping["LAND_AREA"], "Land Rate", mapping["LAND_RATE"]),
        ("Land Value", mapping["LAND_VALUE"], "Built-up Area", mapping["BUILTUP_AREA"]),
        ("Construction Rate", mapping["CONSTRUCTION_RATE"], "Gross Building Value", mapping["GROSS_BUILDING_VALUE"]),
        ("Depreciation %", mapping["DEPRECIATION_PERCENT"], "Depreciation Amount", mapping["DEPRECIATION_AMOUNT"]),
        ("Net Building Value", mapping["NET_BUILDING_VALUE"], "Market Value", mapping["MARKET_VALUE"]),
        ("Conservative Value", mapping["CONSERVATIVE_VALUE"], "Distress Value", mapping["DISTRESS_VALUE"]),
        ("Government Value", mapping["GOVT_VALUE"], "Report Date", mapping["REPORT_DATE"]),
        ("Remarks", mapping["REMARKS"], "", ""),
    ]
    for r, row in enumerate(ordered, 3):
        for c, value in enumerate(row, 1):
            cell = ws.cell(r, c, value)
            cell.border = border
            if c in (1, 3):
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="DDEBF7")
    start = 3 + len(ordered) + 2
    ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=4)
    ws.cell(start, 1, "UPLOADED DOCUMENTS / PHOTOS").font = Font(bold=True)
    for c, value in enumerate(["Type", "Category", "Filename", "Extraction"], 1):
        ws.cell(start + 1, c, value).font = Font(bold=True)
        ws.cell(start + 1, c).border = border
    for r, asset in enumerate(assets, start + 2):
        extraction = safe_json(asset.extraction_json)
        values = [asset.asset_type, asset.category or "", asset.filename, json.dumps(extraction, ensure_ascii=False)]
        for c, value in enumerate(values, 1):
            ws.cell(r, c, value).border = border
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
    for col, width in zip("ABCD", [24, 34, 26, 55]):
        ws.column_dimensions[col].width = width
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


@app.before_request
def ensure_setup():
    db.create_all()
    admin_email = os.getenv("ADMIN_EMAIL", "sakshamvaluer@yahoo.com").lower()
    admin_password = os.getenv("ADMIN_PASSWORD", "ChangeMe123!")
    if not User.query.filter_by(email=admin_email).first():
        db.session.add(User(
            email=admin_email,
            password_hash=generate_password_hash(admin_password),
            name="Saksham Associate Admin",
            role="admin",
        ))
        db.session.commit()


@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "database": "connected",
        "ai_enabled": bool(AI_CLIENT),
        "time": datetime.utcnow().isoformat(),
    })


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = User.query.filter_by(email=request.form.get("email", "").strip().lower()).first()
        if user and check_password_hash(user.password_hash, request.form.get("password", "")):
            session.clear()
            session["user_id"] = user.id
            session["role"] = user.role
            session["name"] = user.name
            return redirect(url_for("dashboard"))
        flash("Email or password is incorrect.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def dashboard():
    search = request.args.get("q", "").strip()
    include_archived = request.args.get("archived") == "1"
    query = ValuationCase.query
    if not include_archived:
        query = query.filter_by(archived=False)
    if search:
        pattern = f"%{search}%"
        query = query.filter(db.or_(
            ValuationCase.application_number.ilike(pattern),
            ValuationCase.customer_name.ilike(pattern),
            ValuationCase.property_address.ilike(pattern),
            ValuationCase.bank_name.ilike(pattern),
        ))
    cases = query.order_by(ValuationCase.created_at.desc()).limit(500).all()
    stats = {
        "cases": ValuationCase.query.filter_by(archived=False).count(),
        "documents": FileAsset.query.filter_by(asset_type="document").count(),
        "photos": FileAsset.query.filter_by(asset_type="photo").count(),
        "reports": FileAsset.query.filter_by(asset_type="report").count(),
    }
    return render_template("dashboard.html", cases=cases, stats=stats, search=search, include_archived=include_archived)


@app.route("/cases/new", methods=["GET", "POST"])
@login_required
def new_case():
    if request.method == "POST":
        case = ValuationCase(
            application_number=request.form.get("application_number", "").strip(),
            customer_name=request.form.get("customer_name", "").strip() or "To be reviewed",
            contact_number=request.form.get("contact_number", "").strip(),
            property_address=request.form.get("property_address", "").strip(),
            bank_name=request.form.get("bank_name", "").strip(),
            branch_name=request.form.get("branch_name", "").strip(),
            case_type=request.form.get("case_type", "LAP").strip(),
            visit_by=request.form.get("visit_by", "").strip(),
            status="New - Manual",
        )
        db.session.add(case)
        db.session.commit()
        return redirect(url_for("case_detail", case_id=case.id))
    return render_template("new_case.html")


@app.route("/cases/<int:case_id>")
@login_required
def case_detail(case_id):
    case = ValuationCase.query.get_or_404(case_id)
    valuation = Valuation.query.filter_by(case_id=case_id).first()
    assets = FileAsset.query.filter_by(case_id=case_id).order_by(FileAsset.created_at.desc()).all()
    extraction = {}
    for asset in assets:
        data = safe_json(asset.extraction_json)
        if data:
            extraction[asset.filename] = data
    templates = FileAsset.query.filter_by(asset_type="template").order_by(FileAsset.created_at.desc()).all()
    return render_template(
        "case_detail.html", case=case, valuation=valuation, assets=assets,
        extraction=extraction, templates=templates
    )


@app.route("/cases/<int:case_id>/update", methods=["POST"])
@login_required
def update_case(case_id):
    case = ValuationCase.query.get_or_404(case_id)
    for field in [
        "application_number", "customer_name", "contact_number", "property_address",
        "bank_name", "branch_name", "case_type", "visit_by", "status"
    ]:
        setattr(case, field, request.form.get(field, getattr(case, field)) or "")
    db.session.commit()
    flash("Case details saved.", "success")
    return redirect(url_for("case_detail", case_id=case_id))


@app.route("/cases/<int:case_id>/upload", methods=["POST"])
@login_required
def upload_case_files(case_id):
    case = ValuationCase.query.get_or_404(case_id)
    uploaded = request.files.getlist("files")
    count = 0
    for item in uploaded:
        if not item or not item.filename:
            continue
        filename = secure_filename(item.filename)
        content = item.read()
        ext = Path(filename).suffix.lower()
        if ext == ".zip":
            try:
                with zipfile.ZipFile(io.BytesIO(content)) as zf:
                    for member in safe_zip_members(zf):
                        inner_name = secure_filename(Path(member.filename).name)
                        inner_ext = Path(inner_name).suffix.lower()
                        if inner_ext not in DOCUMENT_EXTENSIONS | PHOTO_EXTENSIONS:
                            continue
                        inner = zf.read(member)
                        asset_type = "photo" if inner_ext in PHOTO_EXTENSIONS else "document"
                        store_asset(case_id, asset_type, inner_name, inner)
                        count += 1
            except zipfile.BadZipFile:
                flash(f"{filename} is not a valid ZIP.", "error")
        elif ext in PHOTO_EXTENSIONS:
            store_asset(case_id, "photo", filename, content, item.mimetype)
            count += 1
        elif ext in DOCUMENT_EXTENSIONS:
            store_asset(case_id, "document", filename, content, item.mimetype)
            count += 1
    if count:
        case.status = "Documents Uploaded"
        db.session.commit()
    flash(f"{count} file(s) processed.", "success")
    return redirect(url_for("case_detail", case_id=case_id))


@app.route("/assets/<int:asset_id>")
@login_required
def download_asset(asset_id):
    asset = FileAsset.query.get_or_404(asset_id)
    return send_file(
        io.BytesIO(asset.content), mimetype=asset.mime_type,
        as_attachment=True, download_name=asset.filename
    )


@app.route("/assets/<int:asset_id>/delete", methods=["POST"])
@login_required
def delete_asset(asset_id):
    asset = FileAsset.query.get_or_404(asset_id)
    case_id = asset.case_id
    db.session.delete(asset)
    db.session.commit()
    return redirect(url_for("case_detail", case_id=case_id))


@app.route("/cases/<int:case_id>/valuation", methods=["POST"])
@login_required
def save_valuation(case_id):
    case = ValuationCase.query.get_or_404(case_id)
    data = valuation_calculation(request.form)
    valuation = Valuation.query.filter_by(case_id=case_id).first() or Valuation(case_id=case_id)
    for key, value in data.items():
        setattr(valuation, key, value)
    db.session.add(valuation)
    case.status = "Valuation Completed"
    db.session.commit()
    flash("Valuation calculated and saved.", "success")
    return redirect(url_for("case_detail", case_id=case_id))


@app.route("/templates", methods=["GET", "POST"])
@login_required
def templates_page():
    if request.method == "POST":
        item = request.files.get("template")
        bank_name = request.form.get("bank_name", "").strip()
        if item and Path(item.filename).suffix.lower() in TEMPLATE_EXTENSIONS:
            content = item.read()
            asset = FileAsset(
                asset_type="template", category=bank_name,
                filename=secure_filename(item.filename), mime_type=item.mimetype,
                content=content,
            )
            db.session.add(asset)
            db.session.commit()
            flash("Bank template uploaded.", "success")
    templates = FileAsset.query.filter_by(asset_type="template").order_by(FileAsset.created_at.desc()).all()
    return render_template("templates.html", templates=templates)


@app.route("/templates/<int:asset_id>/delete", methods=["POST"])
@login_required
def delete_template(asset_id):
    asset = FileAsset.query.get_or_404(asset_id)
    db.session.delete(asset)
    db.session.commit()
    return redirect(url_for("templates_page"))


@app.route("/cases/<int:case_id>/report", methods=["POST"])
@login_required
def generate_report(case_id):
    case = ValuationCase.query.get_or_404(case_id)
    valuation = Valuation.query.filter_by(case_id=case_id).first()
    if not valuation:
        flash("Please save valuation first.", "error")
        return redirect(url_for("case_detail", case_id=case_id))
    assets = FileAsset.query.filter_by(case_id=case_id).all()
    template_id = request.form.get("template_id")
    output = None
    report_name = f"SVAI_{case.application_number or case.id}_{case.customer_name or 'Report'}.xlsx"
    if template_id:
        template = FileAsset.query.filter_by(id=int(template_id), asset_type="template").first()
        if template:
            try:
                output = fill_template(template.content, report_mapping(case, valuation))
            except Exception as exc:
                flash(f"Template mapping failed; generic report used. {exc}", "error")
    if output is None:
        output = generic_report(case, valuation, assets)
    report = FileAsset(
        case_id=case_id, asset_type="report", category=case.bank_name,
        filename=secure_filename(report_name), mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        content=output,
    )
    db.session.add(report)
    case.status = "Report Generated"
    db.session.commit()
    return send_file(
        io.BytesIO(output), as_attachment=True, download_name=report.filename,
        mimetype=report.mime_type
    )


@app.route("/cases/<int:case_id>/archive", methods=["POST"])
@login_required
def archive_case(case_id):
    case = ValuationCase.query.get_or_404(case_id)
    case.archived = not case.archived
    case.status = "Archived" if case.archived else "Reopened"
    db.session.commit()
    return redirect(url_for("dashboard", archived="1" if case.archived else "0"))


@app.route("/email-accounts", methods=["GET", "POST"])
@login_required
def email_accounts():
    if request.method == "POST":
        address = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        provider = request.form.get("provider", "auto")
        custom_host = request.form.get("imap_host", "").strip()
        bank_name = request.form.get("bank_name", "").strip()
        host = detect_imap(address, provider, custom_host)
        try:
            mail = imaplib.IMAP4_SSL(host, 993)
            mail.login(address, password)
            mail.logout()
            account = EmailAccount.query.filter_by(email=address).first() or EmailAccount(email=address)
            account.encrypted_password = encrypt_password(password)
            account.provider = provider
            account.imap_host = custom_host
            account.bank_name = bank_name
            account.active = True
            db.session.add(account)
            db.session.commit()
            flash("Email verified and linked.", "success")
        except Exception as exc:
            flash(f"Email verification failed: {exc}", "error")
    accounts = EmailAccount.query.order_by(EmailAccount.created_at.desc()).all()
    return render_template("email_accounts.html", accounts=accounts)


@app.route("/email-accounts/<int:account_id>/fetch", methods=["POST"])
@login_required
def fetch_one_email(account_id):
    account = EmailAccount.query.get_or_404(account_id)
    try:
        result = fetch_email_account(account)
        flash(f"{result['created']} new valuation case(s), {result['ignored']} unrelated email(s) ignored.", "success")
    except Exception as exc:
        flash(f"Email fetch failed: {exc}", "error")
    return redirect(url_for("email_accounts"))


@app.route("/email-accounts/fetch-all", methods=["POST"])
@login_required
def fetch_all_emails():
    total = 0
    errors = []
    for account in EmailAccount.query.filter_by(active=True).all():
        try:
            total += fetch_email_account(account)["created"]
        except Exception as exc:
            errors.append(f"{account.email}: {exc}")
    flash(f"{total} new case(s) created." + (f" Errors: {'; '.join(errors)}" if errors else ""), "success" if not errors else "error")
    return redirect(url_for("dashboard"))


@app.route("/email-accounts/<int:account_id>/delete", methods=["POST"])
@login_required
def delete_email_account(account_id):
    account = EmailAccount.query.get_or_404(account_id)
    db.session.delete(account)
    db.session.commit()
    return redirect(url_for("email_accounts"))


@app.route("/mis/export")
@login_required
def export_mis():
    cases = ValuationCase.query.order_by(ValuationCase.created_at.desc()).all()
    wb = Workbook()
    ws = wb.active
    ws.title = "Master MIS"
    headers = [
        "SR NO", "DATE", "CUSTOMER NAME", "APPLICATION NO", "CONTACT NUMBER",
        "BANK", "CASE TYPE", "STATUS", "PROPERTY ADDRESS", "VISIT BY",
        "BRANCH", "SOURCE EMAIL", "ARCHIVED"
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
    for index, case in enumerate(cases, 1):
        ws.append([
            index, case.created_at.strftime("%d-%m-%Y") if case.created_at else "",
            case.customer_name, case.application_number, case.contact_number,
            case.bank_name, case.case_type, case.status, case.property_address,
            case.visit_by, case.branch_name, case.source_email,
            "Yes" if case.archived else "No"
        ])
    for column in ws.columns:
        width = min(max(len(str(cell.value or "")) for cell in column) + 2, 45)
        ws.column_dimensions[column[0].column_letter].width = width
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, as_attachment=True, download_name=f"SVAI_MIS_{datetime.now():%Y%m%d}.xlsx")


@app.route("/profile/password", methods=["POST"])
@login_required
def change_password():
    user = User.query.get(session["user_id"])
    current = request.form.get("current_password", "")
    new = request.form.get("new_password", "")
    if not user or not check_password_hash(user.password_hash, current):
        flash("Current password is incorrect.", "error")
    elif len(new) < 8:
        flash("New password must contain at least 8 characters.", "error")
    else:
        user.password_hash = generate_password_hash(new)
        db.session.commit()
        flash("Password changed.", "success")
    return redirect(url_for("dashboard"))


def scheduled_email_fetch():
    with app.app_context():
        for account in EmailAccount.query.filter_by(active=True).all():
            try:
                fetch_email_account(account)
            except Exception as exc:
                app.logger.warning("Scheduled email fetch failed for %s: %s", account.email, exc)


def start_scheduler():
    if not BackgroundScheduler:
        return
    if os.getenv("ENABLE_EMAIL_SCHEDULER", "false").lower() != "true":
        return
    scheduler = BackgroundScheduler(daemon=True)
    minutes = max(5, int(os.getenv("EMAIL_FETCH_MINUTES", "10")))
    scheduler.add_job(scheduled_email_fetch, "interval", minutes=minutes, id="email_fetch", replace_existing=True)
    scheduler.start()


with app.app_context():
    db.create_all()
start_scheduler()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")), debug=os.getenv("FLASK_DEBUG") == "1")
