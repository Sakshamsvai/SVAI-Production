from flask import Blueprint, request, jsonify
import imaplib
import email
from email.header import decode_header
import re
from datetime import datetime
from bs4 import BeautifulSoup

mis_email_bp = Blueprint('mis_email', __name__)

linked_accounts = []
mis_database = []

IMAP_SERVERS = {
    "Yahoo Mail": "imap.mail.yahoo.com",
    "Gmail": "imap.gmail.com",
    "Outlook / Hotmail": "outlook.office365.com",
    "Auto Detect": "imap.mail.yahoo.com"
}

@mis_email_bp.route('/api/get-mis', methods=['GET', 'POST'])
def get_mis():
    return jsonify({"success": True, "mis": mis_database})

@mis_email_bp.route('/api/link-email', methods=['POST'])
def link_email():
    data = request.get_json(force=True) or {}
    user_email = str(data.get('email', '')).strip()
    app_password = str(data.get('password', '')).strip()
    provider = str(data.get('provider', 'Auto Detect')).strip()

    if not user_email or not app_password:
        return jsonify({"success": False, "message": "Email ID aur App Password dono bharna zaroori hai!"})

    imap_host = IMAP_SERVERS.get(provider, "imap.mail.yahoo.com")
    if "gmail.com" in user_email.lower():
        imap_host = "imap.gmail.com"
    elif "yahoo" in user_email.lower():
        imap_host = "imap.mail.yahoo.com"

    linked_accounts.clear()
    linked_accounts.append({
        "email": user_email,
        "password": app_password,
        "server": imap_host,
        "provider": provider
    })

    return jsonify({"success": True, "message": f"✅ {user_email} Linked Successfully!"})

def smart_clean_text(raw_html_or_text):
    if not raw_html_or_text:
        return ""
    try:
        soup = BeautifulSoup(raw_html_or_text, "html.parser")
        return soup.get_text(separator=" ")
    except Exception:
        return raw_html_or_text

def deep_parse_email(subject, clean_text, sender):
    full = f"{subject} {clean_text}"
    
    # 1. Customer Name
    cust_name = None
    patterns_name = [
        r'Applicant Name[\s:-]*([A-Z\s]{3,30})(?:\n|\r|Mobile|App|Case)',
        r'Case of\s*-\s*([A-Z\s]{3,30})(?:\n|\r|//|App|Mobile)',
        r'Technical Report Initiate of\s*-\s*([A-Z\s]{3,30})(?:\n|\r|//|App)'
    ]
    for p in patterns_name:
        m = re.search(p, full, re.IGNORECASE)
        if m and len(m.group(1).strip()) > 2:
            cust_name = m.group(1).strip()
            break

    # 2. Application ID
    app_no = None
    app_match = re.search(r'(?:Application No|App id|App No)[\s:-]*(\d{6,12})', full, re.IGNORECASE)
    if app_match:
        app_no = app_match.group(1).strip()

    # 3. Mobile Number
    mobile = "N/A"
    mob_match = re.search(r'(?:Mobile No|Mobile Number)[\s:-]*([6-9]\d{9})', full, re.IGNORECASE)
    if mob_match:
        mobile = mob_match.group(1).strip()

    # 4. Bank / Company
    bank_name = "SMFG / Grihashakti" if "grihashakti" in sender.lower() or "smfg" in full.lower() else "Bank / FI"

    # 5. Case Type
    case_type = "Valuation Case"
    type_match = re.search(r'Case Type[\s:-]*([^\n\r/]+)', full, re.IGNORECASE)
    if type_match:
        case_type = type_match.group(1).strip()
    elif "RESALE" in full.upper():
        case_type = "HL RESALE PURCHASE"

    return cust_name, app_no, mobile, bank_name, case_type

@mis_email_bp.route('/api/fetch-emails', methods=['POST'])
def fetch_emails():
    if not linked_accounts:
        return jsonify({"success": False, "message": "Pehle Email Link Karein!"})

    acc = linked_accounts[0]
    fetched_count = 0

    try:
        mail = imaplib.IMAP4_SSL(acc['server'])
        mail.login(acc['email'], acc['password'])
        mail.select("inbox")

        status, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()

        recent_ids = email_ids[-15:] if len(email_ids) >= 15 else email_ids
        recent_ids.reverse()

        for e_id in recent_ids:
            _, msg_data = mail.fetch(e_id, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    
                    # Decode Subject
                    subj_hdr = decode_header(msg["Subject"])[0]
                    subject = subj_hdr[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(subj_hdr[1] if subj_hdr[1] else 'utf-8', errors='ignore')
                    
                    sender = msg.get("From", "")
                    
                    # Extract Body
                    body_content = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() in ["text/plain", "text/html"]:
                                body_content += " " + part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    else:
                        body_content = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

                    clean_body = smart_clean_text(body_content)
                    cust_name, app_no, mobile, bank_name, case_type = deep_parse_email(subject, clean_body, sender)

                    # IGNORE JUNK / OTP / UNKNOWN ENTRIES
                    if not cust_name or not app_no or "OTP" in subject:
                        continue

                    # STRICT DUPLICATE BLOCKER
                    is_duplicate = any(m['application_no'] == app_no or (m['customer_name'] == cust_name and m['application_no'] == app_no) for m in mis_database)
                    
                    if not is_duplicate:
                        mis_database.append({
                            "sr_no": len(mis_database) + 1,
                            "date": datetime.now().strftime("%Y-%m-%d"),
                            "customer_name": cust_name,
                            "application_no": app_no,
                            "contact_number": mobile,
                            "bank": bank_name,
                            "status": "New Request",
                            "address": case_type
                        })
                        fetched_count += 1

        mail.logout()
        return jsonify({
            "success": True, 
            "message": f"⚡ SMART CLEANED: {fetched_count} Valid Requests Added into MIS (Duplicates & Junk Filtered)!",
            "mis": mis_database
        })

    except Exception as e:
        return jsonify({"success": False, "message": f"❌ Error: {str(e)}"})
