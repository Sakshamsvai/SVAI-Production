import base64
import io
import json
import os
import re
import secrets
import imaplib
import smtplib
import time
import email as email_lib
import zipfile
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from functools import wraps
from html import unescape
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo

from cryptography.fernet import Fernet
from dotenv import dotenv_values, load_dotenv
from docx import Document
from flask import (
    Flask, abort, flash, jsonify, redirect, render_template, request,
    send_file, session, url_for
)
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pypdf import PdfReader
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"
load_dotenv(ENV_PATH)
# Never inherit an unrelated machine/process API key. SVAI uses only the key
# explicitly saved in its own local .env file.
local_env = dotenv_values(ENV_PATH)
if "OPENAI_API_KEY" in local_env:
    os.environ["OPENAI_API_KEY"] = local_env.get("OPENAI_API_KEY") or ""

from ai_service_openai import (
    OPENAI_MODEL, ai_enabled, build_case_profile, configure_openai,
    classify_property_photo, deterministic_email_candidate,
    document_ai_enabled,
    enrich_email_details_from_attachments, extract_property_asset,
    extract_valuation_email,
)
from report_service import fill_docx_template, fill_excel_template

try:
    from apscheduler.schedulers.background import BackgroundScheduler
except Exception:
    BackgroundScheduler = None


app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "change-this-secret-in-render")
database_url = os.getenv("DATABASE_URL", f"sqlite:///{BASE_DIR / 'svai.db'}")
if database_url.startswith("postgres://"):
    database_url = database_url.replace(
        "postgres://", "postgresql+psycopg://", 1
    )
elif database_url.startswith("postgresql://"):
    database_url = database_url.replace(
        "postgresql://", "postgresql+psycopg://", 1
    )
app.config["SQLALCHEMY_DATABASE_URI"] = database_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "50")) * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.getenv("SESSION_COOKIE_SECURE", "false").lower() == "true"
db = SQLAlchemy(app)

DOCUMENT_EXTENSIONS = {
    ".pdf", ".jpg", ".jpeg", ".png", ".webp", ".docx", ".xlsx"
}
PHOTO_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}
TEMPLATE_EXTENSIONS = {".xlsx", ".xlsm", ".docx"}
SEED_TEMPLATES_DIR = BASE_DIR / "seed_templates"
APP_TIMEZONE = ZoneInfo(os.getenv("APP_TIMEZONE", "Asia/Kolkata"))

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
    conservative_percent = db.Column(db.Float, default=100)
    distress_percent = db.Column(db.Float, default=80)
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
        try:
            Fernet(configured.encode())
            return configured.encode()
        except (TypeError, ValueError):
            # Render's generated secrets are not guaranteed to already be
            # Fernet-formatted. Derive a stable valid key from any secret.
            raw = configured.encode()
            return base64.urlsafe_b64encode(__import__("hashlib").sha256(raw).digest())
    raw = str(app.config["SECRET_KEY"]).encode()
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


def update_env_file(updates):
    """Update selected .env values while preserving unrelated configuration."""
    env_path = BASE_DIR / ".env"
    existing = env_path.read_text(encoding="utf-8").splitlines() if env_path.exists() else []
    remaining = dict(updates)
    output = []
    for line in existing:
        match = re.match(r"^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=", line)
        key = match.group(1) if match else None
        if key in remaining:
            value = str(remaining.pop(key)).replace("\\", "\\\\").replace('"', '\\"')
            output.append(f'{key}="{value}"')
        else:
            output.append(line)
    for key, raw_value in remaining.items():
        value = str(raw_value).replace("\\", "\\\\").replace('"', '\\"')
        output.append(f'{key}="{value}"')
    env_path.write_text("\n".join(output) + "\n", encoding="utf-8")


def detect_imap(email_address: str, provider: str = "auto", custom_host: str = ""):
    if custom_host:
        return custom_host
    domain = email_address.lower().split("@")[-1]
    provider = (provider or "auto").lower()
    if "gmail" in provider or domain in {"gmail.com", "googlemail.com"}:
        return "imap.gmail.com"
    if "yahoo" in provider or "yahoo" in domain:
        return "imap.mail.yahoo.com"
    raise ValueError("Only Gmail and Yahoo IMAP accounts are supported.")


def detect_smtp(email_address: str, provider: str = "auto"):
    domain = email_address.lower().split("@")[-1]
    provider = (provider or "auto").lower()
    if "gmail" in provider or domain in {"gmail.com", "googlemail.com"}:
        return "smtp.gmail.com"
    if "yahoo" in provider or "yahoo" in domain:
        return "smtp.mail.yahoo.com"
    raise ValueError("Only Gmail and Yahoo SMTP accounts are supported.")


def send_password_reset_code(account, recipient, code):
    message = EmailMessage()
    message["Subject"] = "SVAI password reset code"
    message["From"] = account.email
    message["To"] = recipient
    message.set_content(
        "Your SVAI password reset code is:\n\n"
        f"{code}\n\n"
        "This code expires in 10 minutes. If you did not request it, ignore "
        "this email."
    )
    host = detect_smtp(account.email, account.provider)
    password = decrypt_password(account.encrypted_password)
    try:
        with smtplib.SMTP_SSL(host, 465, timeout=12) as smtp:
            smtp.login(account.email, password)
            smtp.send_message(message)
            return
    except (OSError, smtplib.SMTPException):
        # Some cloud hosts block outbound SMTPS/465 while allowing the
        # standard STARTTLS submission port. Gmail and Yahoo support both.
        with smtplib.SMTP(host, 587, timeout=20) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(account.email, password)
            smtp.send_message(message)


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
                        text = re.sub(
                            r"(?is)<(?:script|style)\b[^>]*>.*?</(?:script|style)>",
                            " ",
                            text,
                        )
                        text = re.sub(
                            r"(?i)</?(?:p|div|br|tr|li|table|h[1-6])\b[^>]*>",
                            "\n",
                            text,
                        )
                        text = unescape(re.sub(r"<[^>]+>", " ", text))
                    chunks.append(text)
                except Exception:
                    continue
    else:
        payload = message.get_payload(decode=True) or b""
        chunks.append(payload.decode(message.get_content_charset() or "utf-8", errors="ignore"))
    normalized = []
    for line in "\n".join(chunks).replace("\r", "\n").splitlines():
        line = re.sub(r"[ \t\f\v]+", " ", line).strip()
        if line:
            normalized.append(line)
        elif normalized and normalized[-1] != "":
            normalized.append("")
    return "\n".join(normalized).strip()[:25000]


def ai_extract_email(subject: str, body: str, sender: str):
    return extract_valuation_email(subject, body, sender)


def is_valuation_email(subject: str, body: str, sender: str = "") -> bool:
    return deterministic_email_candidate(subject, body, sender)


def parse_iso_date(value, default):
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return default


def current_month_range():
    today = datetime.now(APP_TIMEZONE).date()
    return today.replace(day=1), today


def collect_email_attachments(message):
    attachments = []
    for part in message.walk():
        raw_name = decode_header_value(part.get_filename() or "")
        if not raw_name:
            continue
        filename = secure_filename(raw_name)
        content = part.get_payload(decode=True) or b""
        ext = Path(filename).suffix.lower()
        if not content or ext not in DOCUMENT_EXTENSIONS | {".zip"}:
            continue
        if ext == ".zip":
            try:
                with zipfile.ZipFile(io.BytesIO(content)) as bundle:
                    for member in safe_zip_members(bundle):
                        inner_name = secure_filename(Path(member.filename).name)
                        if Path(inner_name).suffix.lower() not in DOCUMENT_EXTENSIONS:
                            continue
                        attachments.append({
                            "filename": inner_name,
                            "content": bundle.read(member),
                            "mime_type": "application/octet-stream",
                        })
            except (zipfile.BadZipFile, ValueError):
                continue
        else:
            attachments.append({
                "filename": filename,
                "content": content,
                "mime_type": part.get_content_type(),
            })
    return attachments


def normalized_email_subject(subject):
    value = re.sub(
        r"(?i)^(?:(?:re|fw|fwd)(?:\[\d+\])?\s*:\s*)+",
        "",
        re.sub(r"\s+", " ", subject or "").strip(),
    )
    return value.casefold()


def normalized_application_number(value):
    """Stable key used only for matching; the displayed application stays unchanged."""
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def is_followup_email(subject, body=""):
    """Return True for report reminders/status mails, not new visits or assignments."""
    subject_text = re.sub(r"\s+", " ", subject or "").strip()
    if re.search(
        r"(?i)^(?:(?:re|fw|fwd)(?:\[\d+\])?\s*:\s*)*"
        r"(?:recall\s*:|revise(?:d)?\s+(?:technical|valuation)\s+report\b)",
        subject_text,
    ):
        return True
    latest_body = str(body or "")
    latest_body = re.split(
        r"(?im)^\s*(?:-{2,}\s*original message\s*-{2,}|from\s*:|on .{0,160} wrote\s*:)",
        latest_body,
        maxsplit=1,
    )[0]
    latest_body = re.sub(r"\s+", " ", latest_body[:2500]).strip()
    combined = f"{subject_text}\n{latest_body}"
    followup_patterns = (
        r"(?i)\b(?:please|pls|kindly)\s+(?:share|send|forward)\s+"
        r"(?:the\s+)?(?:(?:technical|valuation|final)\s+)?report\b",
        r"(?i)\b(?:technical|valuation)?\s*report\s+(?:is\s+)?"
        r"(?:awaited|pending|overdue)\b",
        r"(?i)\bawaiting\s+(?:the\s+)?(?:(?:technical|valuation)\s+)?report\b",
        r"(?i)\b(?:reminder|follow[\s-]*up)\b.{0,100}\b(?:report|status)\b",
        r"(?i)\b(?:report|case)\s+status\s+(?:required|needed|update)\b",
    )
    if not any(re.search(pattern, combined) for pattern in followup_patterns):
        return False
    new_visit_patterns = (
        r"(?i)\bsubsequent\s+visit\b",
        r"(?i)\bre[\s-]*visit\b",
        r"(?i)\b(?:part|tranche)\s+(?:valuation|visit|technical)\b",
        r"(?i)\bconstruction\s+stage\b.{0,80}\b(?:visit|required|complete)\b",
        r"(?i)\b(?:new\s+)?(?:technical|valuation)\s+(?:case\s+)?assignment\b",
    )
    return not any(re.search(pattern, latest_body) for pattern in new_visit_patterns)


def _message_already_recorded(case, unique_id):
    if not case or not unique_id:
        return False
    if case.source_message_id == unique_id:
        return True
    stored = safe_json(case.extracted_json)
    return any(
        item.get("message_id") == unique_id
        for item in stored.get("followup_emails", [])
        if isinstance(item, dict)
    )


def existing_case_for_message(account, unique_id, subject, received):
    direct = ValuationCase.query.filter_by(source_message_id=unique_id).first()
    if direct:
        return direct
    recorded = ValuationCase.query.filter(
        ValuationCase.source_email == account.email,
        ValuationCase.extracted_json.contains(unique_id),
    ).all()
    for item in recorded:
        if _message_already_recorded(item, unique_id):
            return item
    start = datetime.combine(received.date(), datetime.min.time())
    end = start + timedelta(days=1)
    candidates = ValuationCase.query.filter(
        ValuationCase.source_email == account.email,
        ValuationCase.email_received_at >= start,
        ValuationCase.email_received_at < end,
    ).all()
    normalized = normalized_email_subject(subject)
    return next(
        (
            item for item in candidates
            if normalized_email_subject(item.email_subject) == normalized
        ),
        None,
    )


def existing_case_for_application(application_number, exclude_case=None, subject=""):
    key = normalized_application_number(application_number)
    if not key:
        return None
    candidates = ValuationCase.query.filter_by(archived=False).order_by(
        db.func.coalesce(
            ValuationCase.email_received_at, ValuationCase.created_at
        ).asc(),
        ValuationCase.id.asc(),
    ).all()
    matches = [
        item for item in candidates
        if (not exclude_case or item.id != exclude_case.id)
        and normalized_application_number(item.application_number) == key
    ]
    thread_subject = normalized_email_subject(subject)
    if thread_subject:
        threaded = next(
            (
                item for item in matches
                if normalized_email_subject(item.email_subject) == thread_subject
            ),
            None,
        )
        if threaded:
            return threaded
    return matches[0] if matches else None


def apply_followup_to_existing_case(
    target, details, attachments, subject, received, unique_id
):
    if _message_already_recorded(target, unique_id):
        return False
    fill_if_missing = {
        "customer_name": details.get("customer_name", ""),
        "contact_number": details.get("contact_number", ""),
        "property_address": details.get("property_address", ""),
        "bank_name": details.get("bank_name", ""),
        "branch_name": details.get("branch_name", ""),
    }
    for field, value in fill_if_missing.items():
        if value and not (getattr(target, field, "") or "").strip():
            setattr(target, field, str(value).strip())
    stored = safe_json(target.extracted_json)
    stored.setdefault("followup_emails", []).append({
        "message_id": unique_id,
        "subject": subject,
        "received_at": received.isoformat() if received else "",
        "action": "Merged into existing MIS case; no new row created",
    })
    target.extracted_json = json.dumps(stored, ensure_ascii=False)
    db.session.commit()
    store_email_attachments(target, attachments)
    return True


def email_case_status(details):
    if details.get("portal_case"):
        return "Portal Pending"
    if details.get("correction_mail"):
        return "Correction Pending"
    required = (
        details.get("application_number"),
        details.get("customer_name"),
        details.get("case_type"),
    )
    return "New - Email" if all(required) else "Email Parsed - Review"


def apply_email_details(case, details, account, subject, received, unique_id):
    existing_row = bool(case.id)
    values = {
        "application_number": details.get("application_number", ""),
        "customer_name": details.get("customer_name", ""),
        "contact_number": details.get("contact_number", ""),
        "property_address": details.get("property_address", ""),
        "bank_name": details.get("bank_name", "") or account.bank_name or "",
        "branch_name": details.get("branch_name", ""),
        "case_type": details.get("case_type", ""),
    }
    invalid_existing = {
        "application_number": {"applicant", "application", "app", "case", "lead"},
        "customer_name": {"to be reviewed", "applicant", "customer", "pending"},
    }
    for field, value in values.items():
        current = (getattr(case, field, "") or "").strip()
        invalid_current = current.casefold() in invalid_existing.get(field, set())
        if value and (not existing_row or not current or invalid_current):
            setattr(case, field, str(value).strip())
        elif invalid_current:
            setattr(case, field, "")
    case.source_email = account.email
    case.source_message_id = case.source_message_id or unique_id
    case.email_subject = case.email_subject or subject
    case.email_received_at = case.email_received_at or received
    email_managed_statuses = {
        "", "New", "New - Email", "Email Parsed - Review",
        "Correction Pending", "Portal Pending",
    }
    if not existing_row or (case.status or "") in email_managed_statuses:
        case.status = email_case_status(details)
    case.archived = False
    stored = safe_json(case.extracted_json)
    stored["email"] = details
    case.extracted_json = json.dumps(stored, ensure_ascii=False)


def store_email_attachments(case, attachments):
    existing = {
        (asset.filename.casefold(), len(asset.content or b""))
        for asset in FileAsset.query.filter_by(case_id=case.id).all()
    }
    stored = 0
    for attachment in attachments:
        key = (attachment["filename"].casefold(), len(attachment["content"]))
        if key in existing:
            continue
        store_asset(
            case.id, "document", attachment["filename"], attachment["content"],
            attachment["mime_type"], source_kind="property_document",
            process_ai=False,
        )
        existing.add(key)
        stored += 1
    return stored


def apply_application_correction(details, existing_case, attachments):
    if not (
        details.get("correction_mail")
        and details.get("application_number")
        and details.get("customer_name")
    ):
        return False
    query = ValuationCase.query.filter(
        db.func.lower(ValuationCase.customer_name)
        == details["customer_name"].strip().lower(),
        ValuationCase.archived.is_(False),
    )
    if existing_case:
        query = query.filter(ValuationCase.id != existing_case.id)
    target = query.order_by(
        db.func.coalesce(ValuationCase.email_received_at, ValuationCase.created_at).desc()
    ).first()
    if not target:
        return False
    target.application_number = details["application_number"]
    stored = safe_json(target.extracted_json)
    stored.setdefault("corrections", []).append({
        "application_number": details["application_number"],
        "customer_name": details["customer_name"],
        "reason": "Application number correction received by email",
    })
    target.extracted_json = json.dumps(stored, ensure_ascii=False)
    target.status = "Email Correction Applied - Review"
    store_email_attachments(target, attachments)
    if existing_case:
        existing_case.archived = True
        existing_case.status = f"Correction Applied to Case #{target.id}"
    db.session.commit()
    return True


def fetch_email_account(account: EmailAccount, start_date=None, end_date=None):
    created = 0
    updated = 0
    ignored = 0
    start_default, end_default = current_month_range()
    start_date = start_date or start_default
    end_date = end_date or end_default
    if start_date > end_date:
        raise ValueError("From date cannot be after To date.")
    host = detect_imap(account.email, account.provider, account.imap_host or "")
    password = decrypt_password(account.encrypted_password)
    mail = imaplib.IMAP4_SSL(host, 993)
    warning = ""
    try:
        mail.login(account.email, password)
        mail.select("INBOX")
        since = start_date.strftime("%d-%b-%Y")
        before = (end_date + timedelta(days=1)).strftime("%d-%b-%Y")
        status, payload = mail.search(None, "SINCE", since, "BEFORE", before)
        if status != "OK":
            return {"created": 0, "ignored": 0, "message": "Inbox search failed"}
        ids = payload[0].split()[-2000:]
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
            followup_mail = is_followup_email(subject, body)
            unique_id = message.get("Message-ID") or f"{account.email}:{msg_id.decode()}"
            received = None
            try:
                received = email_lib.utils.parsedate_to_datetime(message.get("Date")).replace(
                    tzinfo=None
                )
            except Exception:
                received = datetime.utcnow()
            existing_case = existing_case_for_message(
                account, unique_id, subject, received
            )
            if not is_valuation_email(subject, body, sender) and not followup_mail:
                if existing_case and existing_case.source_email:
                    existing_case.archived = True
                    existing_case.status = "Ignored - Not Valuation Email"
                    db.session.commit()
                ignored += 1
                continue

            details = ai_extract_email(subject, body, sender)
            if not details.get("is_valuation", False) and not followup_mail:
                if existing_case and existing_case.source_email:
                    existing_case.archived = True
                    existing_case.status = "Ignored - Not Valuation Email"
                    db.session.commit()
                ignored += 1
                continue

            attachments = collect_email_attachments(message)
            existing_assets = (
                FileAsset.query.filter_by(case_id=existing_case.id).all()
                if existing_case else []
            )
            details = enrich_email_details_from_attachments(
                details,
                [
                    (asset.filename, asset.extracted_text or "")
                    for asset in existing_assets
                ],
            )
            existing_attachment_keys = {
                (asset.filename.casefold(), len(asset.content or b""))
                for asset in existing_assets
            }
            new_attachment_texts = []
            for item in attachments:
                key = (item["filename"].casefold(), len(item["content"]))
                if key in existing_attachment_keys:
                    continue
                new_attachment_texts.append((
                    item["filename"],
                    extract_basic_text(item["filename"], item["content"]),
                ))
            details = enrich_email_details_from_attachments(
                details, new_attachment_texts
            )
            details.pop("ai_error", None)
            if followup_mail:
                if existing_case and _message_already_recorded(
                    existing_case, unique_id
                ):
                    continue
                target = existing_case_for_application(
                    details.get("application_number"), existing_case, subject
                )
                if not target and existing_case and not existing_case.archived:
                    target = existing_case
                if target:
                    if apply_followup_to_existing_case(
                        target, details, attachments, subject, received, unique_id
                    ):
                        updated += 1
                else:
                    ignored += 1
                continue
            if not details.get("is_valuation", False):
                ignored += 1
                continue
            if apply_application_correction(details, existing_case, attachments):
                updated += 1
                continue

            case = existing_case or ValuationCase()
            apply_email_details(case, details, account, subject, received, unique_id)
            db.session.add(case)
            db.session.commit()
            store_email_attachments(case, attachments)
            if existing_case:
                updated += 1
            else:
                created += 1
        account.last_fetch_at = datetime.utcnow()
        db.session.commit()
        return {
            "created": created, "updated": updated, "ignored": ignored,
            "message": f"Fetched {start_date:%d-%m-%Y} to {end_date:%d-%m-%Y}",
            "warning": warning,
        }
    except imaplib.IMAP4.abort as exc:
        warning = f"Mailbox connection ended early: {exc}"
        account.last_fetch_at = datetime.utcnow()
        db.session.commit()
        return {
            "created": created,
            "updated": updated,
            "ignored": ignored,
            "message": f"Partial fetch {start_date:%d-%m-%Y} to {end_date:%d-%m-%Y}",
            "warning": warning,
        }
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
            output = [p.text for p in document.paragraphs if p.text.strip()]
            for table in document.tables:
                for row in table.rows:
                    line = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                    if line:
                        output.append(line)
            return "\n".join(output)[:50000]
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


def ai_extract_document(filename: str, content: bytes, existing_text: str = "", source_kind="property_document"):
    return extract_property_asset(filename, content, existing_text, source_kind)


def store_asset(
    case_id, asset_type, filename, content, mime_type=None, category=None,
    source_kind="property_document", process_ai=True,
):
    filename = secure_filename(filename) or f"file_{secrets.token_hex(4)}"
    text = ""
    extraction = {}
    if asset_type in {"document", "visit_data"}:
        text = extract_basic_text(filename, content)
        if process_ai:
            extraction = ai_extract_document(filename, content, text, source_kind)
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
    max_files = int(os.getenv("MAX_ZIP_FILES", "250"))
    max_uncompressed = int(os.getenv("MAX_ZIP_UNCOMPRESSED_MB", "100")) * 1024 * 1024
    accepted = 0
    total_size = 0
    for member in zf.infolist():
        path = Path(member.filename)
        if member.is_dir() or ".." in path.parts or path.is_absolute():
            continue
        accepted += 1
        total_size += member.file_size
        if accepted > max_files or total_size > max_uncompressed:
            raise ValueError("ZIP is too large after extraction.")
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
        "conservative_percent": f("conservative_percent", 100),
        "distress_percent": f("distress_percent", 80),
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


def numeric_from_value(value, default=0):
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"-?\d[\d,]*(?:\.\d+)?", str(value or ""))
    if not match:
        return float(default)
    try:
        return float(match.group(0).replace(",", ""))
    except ValueError:
        return float(default)


def valuation_defaults_from_profile(profile):
    """Apply the valuer's standing rules without inventing source facts."""
    profile = profile or {}
    stage_text = " ".join(str(profile.get(key) or "") for key in (
        "construction_stage", "construction_quality", "structure_type", "remarks",
    )).casefold()
    if any(token in stage_text for token in ("plinth", "foundation", "dpc")):
        construction_rate, conservative_percent = 300.0, 30.0
    elif any(token in stage_text for token in (
        "without plaster", "unplastered", "brick work", "brickwork", "bare brick",
    )):
        construction_rate, conservative_percent = 700.0, 70.0
    elif any(token in stage_text for token in (
        "complete", "completed", "plaster", "finished", "ready",
    )):
        construction_rate, conservative_percent = 1000.0, 100.0
    else:
        construction_rate, conservative_percent = 0.0, 100.0

    age_years = numeric_from_value(profile.get("property_age_years"))
    if not age_years:
        construction_year = numeric_from_value(profile.get("construction_year"))
        current_year = datetime.now(APP_TIMEZONE).year
        if 1900 <= construction_year <= current_year:
            age_years = float(current_year - int(construction_year))
    if age_years:
        profile["property_age_years"] = age_years
        profile["residual_age_years"] = max(0.0, 60.0 - age_years)

    govt_land_rate = numeric_from_value(profile.get("govt_land_rate"))
    market_land_rate = numeric_from_value(profile.get("land_rate"))
    if not market_land_rate and govt_land_rate:
        market_land_rate = govt_land_rate * 2

    return {
        "land_area": numeric_from_value(profile.get("land_area_as_per_docs")),
        "land_rate": market_land_rate,
        "builtup_area": numeric_from_value(profile.get("builtup_area_as_per_site")),
        "construction_rate": (
            numeric_from_value(profile.get("construction_rate"))
            or construction_rate
        ),
        "age_years": age_years,
        "depreciation_percent": 0.0,
        "govt_land_rate": govt_land_rate,
        "govt_construction_rate": numeric_from_value(
            profile.get("govt_construction_rate")
        ),
        "conservative_percent": conservative_percent,
        "distress_percent": 80.0,
    }


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
def protect_requests():
    session.setdefault("_csrf_token", secrets.token_urlsafe(32))
    if request.method == "POST":
        submitted = request.form.get("_csrf_token") or request.headers.get("X-CSRF-Token", "")
        if not secrets.compare_digest(str(submitted), str(session["_csrf_token"])):
            abort(400, description="Invalid or missing CSRF token.")


@app.context_processor
def template_security():
    return {"csrf_token": session.get("_csrf_token", "")}


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
    if SEED_TEMPLATES_DIR.exists():
        seeds = {
            "DCB.xlsx": "DCB Bank",
            "SBFC.xlsx": "SBFC Finance",
            "Laxmi India.xlsx": "Laxmi India Finance",
            "Ummeed.docx": "Ummeed Housing Finance",
        }
        added = False
        for filename, bank_name in seeds.items():
            source = SEED_TEMPLATES_DIR / filename
            if not source.exists():
                continue
            exists = FileAsset.query.filter_by(asset_type="template", filename=filename).first()
            if exists:
                continue
            mime_type = (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                if source.suffix.lower() == ".docx"
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            db.session.add(FileAsset(
                asset_type="template", category=bank_name, filename=filename,
                mime_type=mime_type, content=source.read_bytes(),
            ))
            added = True
        if added:
            db.session.commit()


@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "database": "connected",
        "ai_enabled": ai_enabled(),
        "document_processing_mode": (
            "Paid ChatGPT + Local Fallback"
            if document_ai_enabled() else "Free Local"
        ),
        "ai_provider": "OpenAI",
        "ai_model": OPENAI_MODEL,
        "time": datetime.utcnow().isoformat(),
    })


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = User.query.filter_by(email=request.form.get("email", "").strip().lower()).first()
        if user and check_password_hash(user.password_hash, request.form.get("password", "")):
            session.clear()
            session["_csrf_token"] = secrets.token_urlsafe(32)
            session["user_id"] = user.id
            session["role"] = user.role
            session["name"] = user.name
            return redirect(url_for("dashboard"))
        flash("Email or password is incorrect.", "error")
    return render_template("login.html")


@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        now = time.time()
        if now - float(session.get("password_reset_last_sent", 0) or 0) < 60:
            flash("Naya code mangne se pehle 60 seconds wait karein.", "error")
            return render_template("forgot_password.html")
        address = request.form.get("email", "").strip().lower()
        user = User.query.filter_by(email=address).first()
        if not user:
            flash("Registered SVAI email nahi mili.", "error")
            return render_template("forgot_password.html")
        account = EmailAccount.query.filter_by(
            email=address, active=True
        ).first() or EmailAccount.query.filter_by(active=True).order_by(
            EmailAccount.created_at.asc()
        ).first()
        recovery_code = os.getenv("ADMIN_RECOVERY_CODE", "").strip()
        recovery_available = (
            address == os.getenv(
                "ADMIN_EMAIL", "sakshamvaluer@yahoo.com"
            ).strip().lower()
            and len(recovery_code) >= 12
        )
        if not account:
            if recovery_available:
                session["password_reset"] = {
                    "user_id": user.id,
                    "email": user.email,
                    "code_hash": generate_password_hash(recovery_code),
                    "expires_at": now + 900,
                    "attempts": 0,
                    "recovery": True,
                }
                session["password_reset_last_sent"] = now
                flash(
                    "Linked mailbox available nahi hai. Administrator recovery "
                    "code se password reset karein.",
                    "success",
                )
                return redirect(url_for("reset_password"))
            flash(
                "Password reset code bhejne ke liye pehle se linked Gmail/Yahoo "
                "account chahiye. Local administrator se reset karwayein.",
                "error",
            )
            return render_template("forgot_password.html")
        code = f"{secrets.randbelow(1_000_000):06d}"
        try:
            send_password_reset_code(account, user.email, code)
        except Exception:
            if recovery_available:
                session["password_reset"] = {
                    "user_id": user.id,
                    "email": user.email,
                    "code_hash": generate_password_hash(recovery_code),
                    "expires_at": now + 900,
                    "attempts": 0,
                    "recovery": True,
                }
                session["password_reset_last_sent"] = now
                flash(
                    "Email service unavailable hai. Administrator recovery "
                    "code se password reset karein.",
                    "success",
                )
                return redirect(url_for("reset_password"))
            flash(
                "Reset code email nahi ho paya. Internet aur linked Gmail/Yahoo "
                "app password check karke dobara try karein.",
                "error",
            )
            return render_template("forgot_password.html")
        session["password_reset"] = {
            "user_id": user.id,
            "email": user.email,
            "code_hash": generate_password_hash(code),
            "expires_at": now + 600,
            "attempts": 0,
        }
        session["password_reset_last_sent"] = now
        flash("6-digit reset code registered email par bhej diya gaya hai.", "success")
        return redirect(url_for("reset_password"))
    return render_template("forgot_password.html")


@app.route("/reset-password", methods=["GET", "POST"])
def reset_password():
    state = session.get("password_reset")
    if not state or float(state.get("expires_at", 0) or 0) < time.time():
        session.pop("password_reset", None)
        flash("Reset code expire ho gaya. Naya code mangayein.", "error")
        return redirect(url_for("forgot_password"))
    if request.method == "POST":
        code = request.form.get("code", "").strip()
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")
        if int(state.get("attempts", 0)) >= 5:
            session.pop("password_reset", None)
            flash("Bahut zyada galat attempts. Naya reset code mangayein.", "error")
            return redirect(url_for("forgot_password"))
        if not check_password_hash(state.get("code_hash", ""), code):
            state["attempts"] = int(state.get("attempts", 0)) + 1
            session["password_reset"] = state
            flash("Reset code sahi nahi hai.", "error")
            return render_template(
                "reset_password.html", reset_email=state.get("email", "")
            )
        if len(new_password) < 8:
            flash("Naya password kam se kam 8 characters ka hona chahiye.", "error")
            return render_template(
                "reset_password.html", reset_email=state.get("email", "")
            )
        if new_password != confirm_password:
            flash("Dono naye passwords same nahi hain.", "error")
            return render_template(
                "reset_password.html", reset_email=state.get("email", "")
            )
        user = db.session.get(User, int(state["user_id"]))
        if not user:
            session.pop("password_reset", None)
            flash("Reset request valid nahi rahi. Dobara try karein.", "error")
            return redirect(url_for("forgot_password"))
        user.password_hash = generate_password_hash(new_password)
        db.session.commit()
        session.clear()
        flash("Password reset ho gaya. Ab naye password se login karein.", "success")
        return redirect(url_for("login"))
    return render_template(
        "reset_password.html", reset_email=state.get("email", "")
    )


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/settings", methods=["GET", "POST"])
@login_required
def settings_page():
    if request.method == "POST":
        api_key = request.form.get("openai_api_key", "").strip()
        new_password = request.form.get("new_password", "")
        paid_document_mode = (
            request.form.get("enable_paid_document_ai", "") == "true"
        )
        changed = False

        if api_key:
            if not (api_key.startswith("sk-") and len(api_key) > 20):
                flash("OpenAI API key sahi nahi lag rahi. Key sk- se shuru honi chahiye.", "error")
                return render_template(
                    "settings.html",
                    ai_enabled=ai_enabled(),
                    ai_model=OPENAI_MODEL,
                    document_ai_enabled=document_ai_enabled(),
                )
            update_env_file({"OPENAI_API_KEY": api_key})
            configure_openai(api_key)
            changed = True

        if paid_document_mode and not ai_enabled():
            flash(
                "Paid ChatGPT document reading on karne se pehle valid OpenAI "
                "API key save karein. Free Local Mode abhi bhi chalta rahega.",
                "error",
            )
            return render_template(
                "settings.html",
                ai_enabled=ai_enabled(),
                ai_model=OPENAI_MODEL,
                document_ai_enabled=document_ai_enabled(),
            )
        requested_mode = "true" if paid_document_mode else "false"
        current_mode = (
            "true"
            if os.getenv("OPENAI_DOCUMENT_EXTRACTION", "false").lower() == "true"
            else "false"
        )
        if requested_mode != current_mode:
            update_env_file({"OPENAI_DOCUMENT_EXTRACTION": requested_mode})
            os.environ["OPENAI_DOCUMENT_EXTRACTION"] = requested_mode
            changed = True

        if new_password:
            if len(new_password) < 8:
                flash("Naya password kam se kam 8 characters ka hona chahiye.", "error")
                return render_template(
                    "settings.html",
                    ai_enabled=ai_enabled(),
                    ai_model=OPENAI_MODEL,
                    document_ai_enabled=document_ai_enabled(),
                )
            user = db.session.get(User, session["user_id"])
            user.password_hash = generate_password_hash(new_password)
            db.session.commit()
            changed = True

        if changed:
            flash("Settings save ho gayi hain.", "success")
        else:
            flash("Koi nayi setting enter nahi ki gayi.", "error")
        return redirect(url_for("settings_page"))

    return render_template(
        "settings.html",
        ai_enabled=ai_enabled(),
        ai_model=OPENAI_MODEL,
        document_ai_enabled=document_ai_enabled(),
    )


def filter_cases_by_dates(query, start_date, end_date):
    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date + timedelta(days=1), datetime.min.time())
    return query.filter(db.or_(
        db.and_(
            ValuationCase.email_received_at.isnot(None),
            ValuationCase.email_received_at >= start_dt,
            ValuationCase.email_received_at < end_dt,
        ),
        db.and_(
            ValuationCase.email_received_at.is_(None),
            ValuationCase.created_at >= start_dt,
            ValuationCase.created_at < end_dt,
        ),
    ))


@app.route("/")
@login_required
def dashboard():
    search = request.args.get("q", "").strip()
    include_archived = request.args.get("archived") == "1"
    default_from, default_to = current_month_range()
    date_from = parse_iso_date(request.args.get("from"), default_from)
    date_to = parse_iso_date(request.args.get("to"), default_to)
    if date_from > date_to:
        date_from, date_to = date_to, date_from
    query = ValuationCase.query
    if not include_archived:
        query = query.filter_by(archived=False)
    query = filter_cases_by_dates(query, date_from, date_to)
    if search:
        pattern = f"%{search}%"
        query = query.filter(db.or_(
            ValuationCase.application_number.ilike(pattern),
            ValuationCase.customer_name.ilike(pattern),
            ValuationCase.property_address.ilike(pattern),
            ValuationCase.bank_name.ilike(pattern),
        ))
    cases = query.order_by(
        db.func.coalesce(ValuationCase.email_received_at, ValuationCase.created_at).desc()
    ).limit(2000).all()
    stats = {
        "cases": ValuationCase.query.filter_by(archived=False).count(),
        "documents": FileAsset.query.filter_by(asset_type="document").count(),
        "photos": FileAsset.query.filter_by(asset_type="photo").count(),
        "reports": FileAsset.query.filter_by(asset_type="report").count(),
    }
    return render_template(
        "dashboard.html", cases=cases, stats=stats, search=search,
        include_archived=include_archived, date_from=date_from, date_to=date_to,
        ai_enabled=ai_enabled(), ai_model=OPENAI_MODEL,
    )


@app.route("/cases/new", methods=["GET", "POST"])
@login_required
def new_case():
    if request.method == "POST":
        application_number = request.form.get("application_number", "").strip()
        if not application_number:
            flash("Application number enter karein.", "error")
            return render_template("new_case.html")
        existing = ValuationCase.query.filter(
            db.func.lower(ValuationCase.application_number) == application_number.lower(),
            ValuationCase.archived.is_(False),
        ).first()
        if existing:
            flash("Ye application number pehle se hai; existing file khol di gayi.", "success")
            return redirect(url_for("case_detail", case_id=existing.id))
        portal_case = request.form.get("portal_case") == "1"
        case = ValuationCase(
            application_number=application_number,
            customer_name=request.form.get("customer_name", "").strip(),
            contact_number=request.form.get("contact_number", "").strip(),
            property_address=request.form.get("property_address", "").strip(),
            bank_name=request.form.get("bank_name", "").strip(),
            branch_name=request.form.get("branch_name", "").strip(),
            case_type=request.form.get("case_type", "").strip(),
            visit_by=request.form.get("visit_by", "").strip(),
            status="Portal Pending" if portal_case else "Files Pending",
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
            extraction[asset.id] = data
    templates = FileAsset.query.filter(
        db.or_(
            FileAsset.asset_type == "template",
            db.and_(FileAsset.asset_type == "case_template", FileAsset.case_id == case_id),
        )
    ).order_by(FileAsset.created_at.desc()).all()
    return render_template(
        "case_detail.html", case=case, valuation=valuation, assets=assets,
        extraction=extraction, templates=templates, ai_enabled=ai_enabled(),
        document_ai_enabled=document_ai_enabled(),
        ai_model=OPENAI_MODEL, case_profile=safe_json(case.extracted_json),
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
    return handle_case_upload(case_id, request.form.get("upload_kind", "visit"))


def handle_case_upload(case_id, upload_kind):
    case = ValuationCase.query.get_or_404(case_id)
    uploaded = request.files.getlist("files")
    count = 0
    page_number = 0
    if upload_kind == "documents":
        default_type = "document"
        source_kind = "property_document"
    elif upload_kind == "visit_form":
        default_type = "visit_data"
        source_kind = "visit_data"
    elif upload_kind == "visit":
        default_type = "photo"
        source_kind = "visit_data"
    else:
        default_type = "photo"
        source_kind = "visit_data"
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
                        if upload_kind == "documents":
                            asset_type = "document"
                        elif upload_kind == "visit_form":
                            asset_type = "visit_data"
                        elif inner_ext in PHOTO_EXTENSIONS:
                            asset_type = "photo"
                        else:
                            asset_type = default_type
                        page_number += 1
                        store_asset(
                            case_id, asset_type, inner_name, inner,
                            category=(
                                f"Visit Form Page {page_number}"
                                if upload_kind == "visit_form" else None
                            ),
                            source_kind=source_kind, process_ai=False,
                        )
                        count += 1
            except (zipfile.BadZipFile, ValueError) as exc:
                message = str(exc) if isinstance(exc, ValueError) else "File is not a valid ZIP."
                flash(f"{filename}: {message}", "error")
        elif ext in PHOTO_EXTENSIONS:
            if upload_kind == "documents":
                asset_type = "document"
            elif upload_kind == "visit_form":
                asset_type = "visit_data"
            else:
                asset_type = "photo"
            page_number += 1
            store_asset(
                case_id, asset_type, filename, content, item.mimetype,
                category=(
                    f"Visit Form Page {page_number}"
                    if upload_kind == "visit_form" else None
                ),
                source_kind=source_kind, process_ai=False,
            )
            count += 1
        elif ext in DOCUMENT_EXTENSIONS:
            page_number += 1
            store_asset(
                case_id, default_type, filename, content, item.mimetype,
                category=(
                    f"Visit Form Page {page_number}"
                    if upload_kind == "visit_form" else None
                ),
                source_kind=source_kind, process_ai=False,
            )
            count += 1
    if count:
        case.status = "Files Uploaded - AI Pending"
        db.session.commit()
    flash(f"{count} file(s) processed.", "success")
    return redirect(url_for("case_detail", case_id=case_id))


def quick_asset_type(filename):
    ext = Path(filename).suffix.lower()
    name = Path(filename).stem.casefold()
    if ext in PHOTO_EXTENSIONS:
        return "photo", "visit_data"
    if any(token in name for token in (
        "visit", "engineer", "site data", "site_data", "inspection", "sketch",
    )):
        return "visit_data", "visit_data"
    return "document", "property_document"


@app.route("/cases/<int:case_id>/upload/all", methods=["POST"])
@login_required
def upload_all_case_files(case_id):
    case = ValuationCase.query.get_or_404(case_id)
    count = 0
    for item in request.files.getlist("files"):
        if not item or not item.filename:
            continue
        filename = secure_filename(item.filename)
        content = item.read()
        ext = Path(filename).suffix.lower()
        if ext == ".zip":
            try:
                with zipfile.ZipFile(io.BytesIO(content)) as bundle:
                    for member in safe_zip_members(bundle):
                        inner_name = secure_filename(Path(member.filename).name)
                        if Path(inner_name).suffix.lower() not in DOCUMENT_EXTENSIONS:
                            continue
                        asset_type, source_kind = quick_asset_type(inner_name)
                        store_asset(
                            case.id, asset_type, inner_name, bundle.read(member),
                            source_kind=source_kind, process_ai=False,
                        )
                        count += 1
            except (zipfile.BadZipFile, ValueError) as exc:
                message = str(exc) if isinstance(exc, ValueError) else "File is not a valid ZIP."
                flash(f"{filename}: {message}", "error")
        elif ext in DOCUMENT_EXTENSIONS:
            asset_type, source_kind = quick_asset_type(filename)
            store_asset(
                case.id, asset_type, filename, content, item.mimetype,
                source_kind=source_kind, process_ai=False,
            )
            count += 1

    template = request.files.get("template")
    if template and template.filename:
        filename = secure_filename(template.filename)
        if Path(filename).suffix.lower() in TEMPLATE_EXTENSIONS:
            db.session.add(FileAsset(
                case_id=case.id,
                asset_type="case_template",
                category=case.bank_name,
                filename=filename,
                mime_type=template.mimetype,
                content=template.read(),
            ))
            count += 1
        else:
            flash("Report format sirf XLSX, XLSM ya DOCX hona chahiye.", "error")

    if count:
        case.status = (
            "Portal Pending"
            if case.status == "Portal Pending"
            else "Files Uploaded - Review Pending"
        )
        db.session.commit()
    flash(f"{count} file(s) upload ho gayi. Ab details check karke report banayein.", "success")
    return redirect(url_for("case_detail", case_id=case.id))


@app.route("/cases/<int:case_id>/upload/visit", methods=["POST"])
@login_required
def upload_visit_files(case_id):
    return handle_case_upload(case_id, "visit")


@app.route("/cases/<int:case_id>/upload/visit-form", methods=["POST"])
@login_required
def upload_visit_form_files(case_id):
    return handle_case_upload(case_id, "visit_form")


@app.route("/cases/<int:case_id>/upload/documents", methods=["POST"])
@login_required
def upload_property_documents(case_id):
    return handle_case_upload(case_id, "documents")


@app.route("/cases/<int:case_id>/upload/template", methods=["POST"])
@login_required
def upload_case_template(case_id):
    case = ValuationCase.query.get_or_404(case_id)
    item = request.files.get("template")
    if not item or not item.filename:
        flash("Select a bank valuation template.", "error")
        return redirect(url_for("case_detail", case_id=case_id))
    filename = secure_filename(item.filename)
    if Path(filename).suffix.lower() not in TEMPLATE_EXTENSIONS:
        flash("Only .xlsx, .xlsm or .docx report formats are supported.", "error")
        return redirect(url_for("case_detail", case_id=case_id))
    db.session.add(FileAsset(
        case_id=case.id, asset_type="case_template",
        category=request.form.get("bank_name", "").strip() or case.bank_name,
        filename=filename, mime_type=item.mimetype, content=item.read(),
    ))
    case.status = "Report Format Uploaded"
    db.session.commit()
    flash("Case-specific valuation format uploaded.", "success")
    return redirect(url_for("case_detail", case_id=case_id))


def valuation_as_dict(valuation):
    if not valuation:
        return {}
    fields = [
        "land_area", "land_rate", "builtup_area", "construction_rate", "age_years",
        "depreciation_percent", "govt_land_rate", "govt_construction_rate",
        "conservative_percent", "distress_percent", "land_value",
        "gross_building_value", "depreciation_amount", "net_building_value",
        "market_value", "conservative_value", "distress_value", "govt_value", "remarks",
    ]
    return {field: getattr(valuation, field) for field in fields}


@app.route("/cases/<int:case_id>/process-ai", methods=["POST"])
@login_required
def process_case_ai(case_id):
    case = ValuationCase.query.get_or_404(case_id)
    paid_mode = document_ai_enabled()
    assets = FileAsset.query.filter_by(case_id=case_id).all()
    document_extractions = []
    visit_extractions = []
    processed = 0
    for asset in assets:
        if asset.asset_type == "photo":
            result = classify_property_photo(asset.filename, asset.content)
            category = result.get("category", "Other Site Photo")
            asset.category = category
            asset.extraction_json = json.dumps(result, ensure_ascii=False)
            processed += 1
        elif asset.asset_type in {"document", "visit_data"}:
            source_kind = "visit_data" if asset.asset_type == "visit_data" else "property_document"
            asset.extracted_text = asset.extracted_text or extract_basic_text(asset.filename, asset.content)
            extraction = ai_extract_document(
                asset.filename, asset.content, asset.extracted_text or "", source_kind
            )
            asset.extraction_json = json.dumps(extraction, ensure_ascii=False)
            processed += 1
        db.session.add(asset)
    db.session.commit()

    for asset in assets:
        extraction = safe_json(asset.extraction_json)
        if not extraction:
            continue
        if asset.asset_type == "document":
            document_extractions.append(extraction)
        elif asset.asset_type == "visit_data":
            visit_extractions.append(extraction)
    current = safe_json(case.extracted_json)
    email_data = current.get("email", current if "application_number" in current else {})
    valuation = Valuation.query.filter_by(case_id=case_id).first()
    case_profile = build_case_profile(
        email_data, document_extractions, visit_extractions, valuation_as_dict(valuation)
    )
    previous_profile = current.get("case_profile", {})
    if previous_profile.get("source_reviewed"):
        for field in SOURCE_REVIEW_FIELDS:
            if field in previous_profile:
                case_profile[field] = previous_profile[field]
        case_profile["source_reviewed"] = True
        case_profile["source_reviewed_at"] = previous_profile.get(
            "source_reviewed_at", ""
        )
        case_profile["survey_khasra_plot_no"] = (
            case_profile.get("survey_khasra_plot_no_as_per_docs")
            or case_profile.get("survey_khasra_plot_no_as_per_site")
            or ""
        )
    case.extracted_json = json.dumps(
        {"email": email_data, "case_profile": case_profile}, ensure_ascii=False, default=str
    )
    valuation = valuation or Valuation(case_id=case_id)
    draft_sources = valuation_defaults_from_profile(case_profile)
    for field, source_value in draft_sources.items():
        current_value = getattr(valuation, field, 0)
        standing_rule_field = field in {
            "depreciation_percent", "conservative_percent", "distress_percent",
        }
        old_generic_default = (
            (field == "conservative_percent" and current_value == 80)
            or (field == "distress_percent" and current_value == 70)
        )
        if (
            source_value not in ("", None)
            and (not current_value or standing_rule_field or old_generic_default)
        ):
            setattr(valuation, field, numeric_from_value(source_value))
    recalculated = valuation_calculation({
        "land_area": valuation.land_area,
        "land_rate": valuation.land_rate,
        "builtup_area": valuation.builtup_area,
        "construction_rate": valuation.construction_rate,
        "age_years": valuation.age_years,
        "depreciation_percent": valuation.depreciation_percent,
        "govt_land_rate": valuation.govt_land_rate,
        "govt_construction_rate": valuation.govt_construction_rate,
        "conservative_percent": valuation.conservative_percent,
        "distress_percent": valuation.distress_percent,
        "remarks": valuation.remarks or case_profile.get("remarks", ""),
    })
    for field, value in recalculated.items():
        setattr(valuation, field, value)
    db.session.add(valuation)
    if not case.application_number:
        case.application_number = str(case_profile.get("application_number", ""))
    if not case.customer_name or case.customer_name == "To be reviewed":
        case.customer_name = (
            case_profile.get("customer_name") or case_profile.get("applicant_name")
            or ""
        )
    if not case.property_address:
        case.property_address = (
            case_profile.get("property_address_as_per_site")
            or case_profile.get("property_address_as_per_docs") or ""
        )
    if not case.contact_number:
        case.contact_number = str(case_profile.get("contact_number", ""))
    case.status = (
        "AI Processed - Review Data"
        if paid_mode else "Local Files Processed - Review Data"
    )
    db.session.commit()
    flash(
        (
            f"ChatGPT + local fallback processed {processed} file(s)."
            if paid_mode else
            f"Free local mode processed {processed} file(s)."
        )
        + " Review extracted data and valuation before report generation.",
        "success",
    )
    return redirect(url_for("case_detail", case_id=case_id))


SOURCE_REVIEW_FIELDS = (
    "owner_name",
    "title_document_number",
    "registration_number",
    "registration_date",
    "land_tenure",
    "approving_authority",
    "plan_details",
    "construction_permission",
    "property_address_as_per_docs",
    "survey_khasra_plot_no_as_per_docs",
    "land_area_as_per_docs",
    "builtup_area_as_per_docs",
    "north_boundary_as_per_docs",
    "south_boundary_as_per_docs",
    "east_boundary_as_per_docs",
    "west_boundary_as_per_docs",
    "property_usage_as_per_docs",
    "property_address_as_per_site",
    "survey_khasra_plot_no_as_per_site",
    "land_area_as_per_site",
    "builtup_area_as_per_site",
    "north_boundary_as_per_site",
    "south_boundary_as_per_site",
    "east_boundary_as_per_site",
    "west_boundary_as_per_site",
    "property_usage_as_per_site",
    "road_width",
    "road_type",
    "latitude",
    "longitude",
    "occupancy",
    "person_met",
    "visit_engineer",
    "visit_date",
    "structure_type",
    "construction_quality",
    "construction_stage",
    "construction_year",
    "property_age_years",
    "residual_age_years",
    "number_of_floors",
    "floor_wise_usage",
    "room_configuration",
)


@app.route("/cases/<int:case_id>/source-review", methods=["POST"])
@login_required
def save_source_review(case_id):
    case = ValuationCase.query.get_or_404(case_id)
    stored = safe_json(case.extracted_json)
    profile = dict(stored.get("case_profile") or stored.get("email") or stored)
    for field in SOURCE_REVIEW_FIELDS:
        if field in request.form:
            profile[field] = request.form.get(field, "").strip()
    profile["survey_khasra_plot_no"] = (
        profile.get("survey_khasra_plot_no_as_per_docs")
        or profile.get("survey_khasra_plot_no_as_per_site")
        or ""
    )
    profile["source_reviewed"] = True
    profile["source_reviewed_at"] = datetime.now(APP_TIMEZONE).isoformat()
    email_data = stored.get("email", {})
    case.extracted_json = json.dumps(
        {"email": email_data, "case_profile": profile},
        ensure_ascii=False,
        default=str,
    )
    case.property_address = (
        profile.get("property_address_as_per_site")
        or profile.get("property_address_as_per_docs")
        or case.property_address
        or ""
    )
    case.status = "Source Data Reviewed"
    db.session.commit()
    flash(
        "Documents aur Actual Site ki reviewed details save ho gayi. "
        "Ab report isi separation se banegi.",
        "success",
    )
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
    if case.status == "Portal Pending":
        flash("Ye online portal case hai; local report pending rakhi gayi hai.", "error")
        return redirect(url_for("case_detail", case_id=case_id))
    valuation = Valuation.query.filter_by(case_id=case_id).first()
    if not valuation:
        valuation = Valuation(case_id=case_id, remarks="DRAFT: valuation figures pending review")
        db.session.add(valuation)
        db.session.commit()
    assets = FileAsset.query.filter_by(case_id=case_id).all()
    template_id = request.form.get("template_id")
    stored = safe_json(case.extracted_json)
    profile = dict(stored.get("case_profile") or stored.get("email") or stored)
    profile.update({
        "application_number": case.application_number,
        "customer_name": case.customer_name,
        "contact_number": case.contact_number,
        "property_address": case.property_address,
        "bank_name": case.bank_name,
        "branch_name": case.branch_name,
        "case_type": case.case_type,
        "visit_by": case.visit_by,
        "report_date": datetime.now(APP_TIMEZONE).strftime("%d-%m-%Y"),
        **valuation_as_dict(valuation),
    })
    photo_assets = [
        {
            "filename": asset.filename,
            "category": asset.category,
            "content": asset.content,
        }
        for asset in assets if asset.asset_type == "photo"
    ]
    visit_images = [
        asset for asset in assets
        if (
            asset.asset_type == "visit_data"
            and Path(asset.filename).suffix.lower() in PHOTO_EXTENSIONS
            and (asset.category or "").startswith("Visit Form Page ")
        )
    ]
    visit_images.sort(
        key=lambda asset: numeric_from_value(asset.category or "", asset.id)
    )
    if len(visit_images) >= 4:
        for asset, category in (
            (visit_images[-2], "Site Sketch"),
            (visit_images[-1], "Location Map"),
        ):
            photo_assets.append({
                "filename": asset.filename,
                "category": category,
                "content": asset.content,
            })
    output = None
    extension = ""
    mime_type = ""
    template = None
    if template_id:
        try:
            numeric_template_id = int(template_id)
        except (TypeError, ValueError):
            numeric_template_id = 0
        template = FileAsset.query.filter(
            FileAsset.id == numeric_template_id,
            db.or_(
                FileAsset.asset_type == "template",
                db.and_(FileAsset.asset_type == "case_template", FileAsset.case_id == case_id),
            ),
        ).first()
    if template is None:
        template = FileAsset.query.filter_by(
            case_id=case_id, asset_type="case_template"
        ).order_by(FileAsset.created_at.desc()).first()
    if template is None:
        flash(
            "Pehle bank ka original Excel/Word valuation format upload aur select karein. "
            "SVAI generic format nahi banayega.",
            "error",
        )
        return redirect(url_for("case_detail", case_id=case_id))
    try:
        extension = Path(template.filename).suffix.lower()
        if extension == ".docx":
            output = fill_docx_template(
                template.content,
                profile,
                photo_assets,
                template_name=template.filename,
                bank_name=template.category or case.bank_name,
            )
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            output = fill_excel_template(
                template.content, profile, photo_assets,
                template_name=template.filename,
                bank_name=template.category or case.bank_name,
            )
            if extension == ".xlsm":
                mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12"
            else:
                extension = ".xlsx"
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    except Exception as exc:
        flash(
            f"Original bank format ko safe tarike se fill nahi kiya ja saka; "
            f"generic/ टूटा report नहीं बनाया गया. {exc}",
            "error",
        )
        return redirect(url_for("case_detail", case_id=case_id))
    report_name = (
        f"SVAI_DRAFT_{case.application_number or case.id}_"
        f"{case.customer_name or 'Pending_Name'}{extension}"
    )
    report = FileAsset(
        case_id=case_id, asset_type="report", category=case.bank_name,
        filename=secure_filename(report_name), mime_type=mime_type,
        content=output,
    )
    db.session.add(report)
    case.status = "Draft Report Generated"
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
        password = re.sub(r"\s+", "", request.form.get("password", ""))
        provider = request.form.get("provider", "").strip().lower()
        domain = address.rsplit("@", 1)[-1] if "@" in address else ""
        gmail_address = domain in {"gmail.com", "googlemail.com"}
        yahoo_address = domain == "yahoo.com" or domain.startswith("yahoo.")
        if provider not in {"gmail", "yahoo"} or not (gmail_address or yahoo_address):
            flash("Only Gmail and Yahoo email IDs are supported.", "error")
            return redirect(url_for("email_accounts"))
        if (provider == "gmail" and not gmail_address) or (provider == "yahoo" and not yahoo_address):
            flash("Selected provider does not match the email ID.", "error")
            return redirect(url_for("email_accounts"))
        if len(password) != 16:
            flash("Enter the 16-character Gmail/Yahoo app password (spaces are ignored).", "error")
            return redirect(url_for("email_accounts"))
        custom_host = ""
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
    date_from, date_to = current_month_range()
    return render_template(
        "email_accounts.html", accounts=accounts, date_from=date_from, date_to=date_to,
        ai_enabled=ai_enabled(), ai_model=OPENAI_MODEL,
    )


@app.route("/email-accounts/<int:account_id>/fetch", methods=["POST"])
@login_required
def fetch_one_email(account_id):
    account = EmailAccount.query.get_or_404(account_id)
    default_from, default_to = current_month_range()
    date_from = parse_iso_date(request.form.get("from"), default_from)
    date_to = parse_iso_date(request.form.get("to"), default_to)
    try:
        result = fetch_email_account(account, date_from, date_to)
        flash(
            f"{result['created']} new, {result.get('updated', 0)} corrected valuation "
            f"case(s); {result['ignored']} unrelated email(s) ignored."
            + (f" {result.get('warning')}" if result.get("warning") else ""),
            "error" if result.get("warning") else "success",
        )
    except Exception as exc:
        flash(f"Email fetch failed: {exc}", "error")
    return redirect(url_for("email_accounts"))


@app.route("/email-accounts/fetch-all", methods=["POST"])
@login_required
def fetch_all_emails():
    total = 0
    updated = 0
    errors = []
    warnings = []
    ignored = 0
    default_from, default_to = current_month_range()
    date_from = parse_iso_date(request.form.get("from"), default_from)
    date_to = parse_iso_date(request.form.get("to"), default_to)
    for account in EmailAccount.query.filter_by(active=True).all():
        try:
            result = fetch_email_account(account, date_from, date_to)
            total += result["created"]
            updated += result.get("updated", 0)
            ignored += result["ignored"]
            if result.get("warning"):
                warnings.append(f"{account.email}: {result['warning']}")
        except Exception as exc:
            errors.append(f"{account.email}: {exc}")
    flash(
        f"{total} valuation case(s) added; {updated} existing case(s) corrected; "
        f"{ignored} unrelated email(s) ignored. "
        f"Range: {date_from:%d-%m-%Y} to {date_to:%d-%m-%Y}."
        + (f" Warnings: {'; '.join(warnings)}" if warnings else "")
        + (f" Errors: {'; '.join(errors)}" if errors else ""),
        "success" if not errors and not warnings else "error",
    )
    return redirect(url_for(
        "dashboard", **{"from": date_from.isoformat(), "to": date_to.isoformat()}
    ))


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
    default_from, default_to = current_month_range()
    date_from = parse_iso_date(request.args.get("from"), default_from)
    date_to = parse_iso_date(request.args.get("to"), default_to)
    query = filter_cases_by_dates(ValuationCase.query, date_from, date_to)
    cases = query.order_by(
        db.func.coalesce(ValuationCase.email_received_at, ValuationCase.created_at)
    ).all()
    mis_template = SEED_TEMPLATES_DIR / "MIS format.xlsx"
    if mis_template.exists():
        wb = load_workbook(mis_template)
        ws = wb["ALL BANK"] if "ALL BANK" in wb.sheetnames else wb.active
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "ALL BANK"
        ws.append([
            "SR NO", "Date", "CUSTOMER NAME", "APPLICATION NO", "CONTACT NUMBER",
            "BANK", "CASE TYPE", "STATUS", "ADDRESS", "VISIT BY", "BRANCH",
            "Pending", "K.M",
        ])
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="FFE699")
    for index, case in enumerate(cases, 1):
        stored = safe_json(case.extracted_json)
        profile = stored.get("case_profile") or stored.get("email") or stored
        case_date = case.email_received_at or case.created_at
        ws.append([
            index, case_date if case_date else "",
            case.customer_name, case.application_number, case.contact_number,
            case.bank_name, case.case_type, case.status, case.property_address,
            case.visit_by, case.branch_name, profile.get("pending", ""),
            profile.get("distance_from_branch", profile.get("km", "")),
        ])
        ws.cell(index + 1, 2).number_format = "d-mmm-yy"
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(
        out, as_attachment=True,
        download_name=f"SVAI_MIS_{date_from:%Y%m%d}_{date_to:%Y%m%d}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


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
