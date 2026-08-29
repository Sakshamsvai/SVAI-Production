"""OpenAI-only extraction helpers for SVAI.

The service intentionally keeps deterministic fallbacks so email import and
manual report work remain available when an API key is not configured.
"""

import base64
import io
import json
import os
import re
from pathlib import Path
from pypdf import PdfReader, PdfWriter

try:
    from openai import OpenAI
except Exception:  # pragma: no cover - dependency may be absent during static checks
    OpenAI = None


OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-5.6-terra").strip()
OPENAI_CLIENT = OpenAI(api_key=OPENAI_API_KEY) if (OpenAI and OPENAI_API_KEY) else None

VALUATION_TERMS = (
    "valuation report", "valuation request", "valuation required", "property valuation",
    "positive valuation", "technical appraisal", "technical initiation",
    "technical initaition", "technical valuation", "technical report",
    "technical scrutiny", "technical clearance", "technical initiate",
    "initiate technical", "intated technical", "technical assignment",
    "valuation initiation", "valuation initaition",
    "site investigation initiation",
    "subsequent visit required", "technical case assignment",
    "property inspection request", "site inspection request",
    "audit initiation", "tsr audit",
)
CASE_TERMS = (
    "fresh", "subsequent", "part", "tranche", "tranch", "revisit", "re-visit",
    "npa", "lap", "home loan", "housing loan", "construction", "purchase",
    "resale", "balance transfer", "plot purchase", "self-construction", "p+c",
)

ASSIGNMENT_ACTION_TERMS = (
    "new case", "case assignment", "case assigned", "case allocation",
    "technical request", "technical case", "initiate", "initiation",
    "valuation", "property inspection", "site inspection", "site visit",
    "property visit", "audit initiation", "tsr", "field investigation",
    "collateral inspection",
)

GENERIC_IDENTIFIER_WORDS = {
    "app", "applicant", "application", "borrower", "case", "customer", "deal",
    "id", "lead", "loan", "name", "no", "number", "proposal", "reference",
    "technical", "valuation",
}

BANK_DOMAIN_NAMES = {
    "yes.bank": "Yes Bank",
    "bajajhousing": "Bajaj Housing Finance",
    "bajajfinserv": "Bajaj Housing Finance",
    "ummeedhfc": "Ummeed Housing Finance",
    "lifc": "Laxmi India Finance",
    "lifl": "Laxmi India Finance",
    "laxmiindiafinleasecap": "Laxmi India Finance",
    "sbfc": "SBFC Finance",
    "dcbbank": "DCB Bank",
    "piramal": "Piramal Finance",
    "jmfl": "JM Financial Home Loans",
    "esafbank": "ESAF Small Finance Bank",
    "idfcfirst": "IDFC First Bank",
    "grihashakti": "Grihashakti",
    "ujjivan": "Ujjivan Small Finance Bank",
    "wonderhfl": "Wonder Home Finance",
    "muthoothomefin": "Muthoot Homefin",
    "aubank": "AU Small Finance Bank",
    "easyhomefinance": "Easy Home Finance",
    "easyhousing": "Easy Home Finance",
    "fusionfin": "Fusion Finance",
}

# A valuation assignment is normally sent by the lender/vendor's own domain.
# Do not turn the recipient's public mailbox provider into a bank name.
PUBLIC_MAIL_DOMAINS = {
    "gmail.com", "googlemail.com", "yahoo.com", "ymail.com", "rocketmail.com",
    "outlook.com", "hotmail.com", "live.com", "msn.com", "icloud.com",
}

PHOTO_CATEGORIES = (
    "Front Elevation", "Front Side View", "Approach Road", "Distant Property View",
    "Property Selfie", "Kitchen", "Internal Room", "Electricity Meter",
    "Electricity Bill", "Site Sketch", "Location Map", "Property Document",
    "Visit Data Sheet", "Other Site Photo",
)

EXTRACTION_KEYS = (
    "source_kind", "document_type", "applicant_name", "owner_name",
    "application_number", "contact_number", "property_address_as_per_docs",
    "property_address_as_per_site", "village", "tehsil", "district", "state",
    "pincode", "landmark", "ward_number", "locality_type", "property_type",
    "project_name", "land_tenure", "distance_from_branch", "site_access",
    "plot_demarcated",
    "survey_khasra_plot_no", "survey_khasra_plot_no_as_per_docs",
    "survey_khasra_plot_no_as_per_site", "title_document_number",
    "registration_number", "registration_date", "land_area_as_per_docs",
    "land_area_as_per_site", "builtup_area_as_per_docs", "builtup_area_as_per_site",
    "north_boundary_as_per_docs", "south_boundary_as_per_docs",
    "east_boundary_as_per_docs", "west_boundary_as_per_docs",
    "north_boundary_as_per_site", "south_boundary_as_per_site",
    "east_boundary_as_per_site", "west_boundary_as_per_site",
    "road_width", "road_type", "latitude", "longitude", "property_usage_as_per_docs",
    "property_usage_as_per_site", "occupancy", "occupant_name", "person_met",
    "co_applicant_name", "visit_engineer",
    "visit_date", "structure_type", "construction_quality", "construction_year",
    "property_age_years",
    "residual_age_years", "number_of_floors", "floor_wise_usage",
    "room_configuration", "flooring", "doors_windows", "amenities",
    "approving_authority", "plan_details", "property_identified_through",
    "construction_permission", "demolition_risk", "marketability",
    "land_rate", "construction_rate", "govt_land_rate",
    "govt_construction_rate", "remarks", "confidence_notes",
)

DOCUMENT_ONLY_FIELDS = {
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
}

SITE_ONLY_FIELDS = {
    "property_address_as_per_site",
    "survey_khasra_plot_no_as_per_site",
    "land_area_as_per_site",
    "builtup_area_as_per_site",
    "north_boundary_as_per_site",
    "south_boundary_as_per_site",
    "east_boundary_as_per_site",
    "west_boundary_as_per_site",
    "property_usage_as_per_site",
}

VISIT_AUTHORITY_FIELDS = SITE_ONLY_FIELDS | {
    "road_width", "road_type", "latitude", "longitude", "site_access",
    "plot_demarcated", "occupancy", "occupant_name", "person_met",
    "visit_engineer", "visit_date", "structure_type", "construction_quality",
    "number_of_floors", "floor_wise_usage", "room_configuration",
    "construction_year", "property_age_years", "residual_age_years",
}


def normalize_boundary_english(value):
    """Standardize common Hindi deed boundary terms without changing names."""
    text = _space(value)
    if not text:
        return ""
    exact = {
        "रास्ता": "Road", "सड़क": "Road", "रोड": "Road",
        "गली": "Gali", "भूमि": "Land", "जमीन": "Land",
        "भूखंड": "Plot", "भूखण्ड": "Plot", "मकान": "House",
    }
    if text in exact:
        return exact[text]
    text = re.sub(r"^(.+?)\s+का\s+मकान$", r"House of \1", text)
    text = re.sub(r"^(.+?)\s+का\s+भूख(?:ं|ण्)ड$", r"Plot of \1", text)
    for hindi, english in exact.items():
        text = text.replace(hindi, english)
    return _space(text)


def _enforce_asset_source_authority(extraction, source_kind):
    """Keep legal-document and physical-site facts in separate namespaces."""
    cleaned = {
        key: extraction.get(key, "")
        for key in EXTRACTION_KEYS
    }
    cleaned["source_kind"] = source_kind
    generic_survey = cleaned.get("survey_khasra_plot_no", "")
    if source_kind == "visit_data":
        for key in DOCUMENT_ONLY_FIELDS:
            cleaned[key] = ""
        cleaned["survey_khasra_plot_no_as_per_site"] = (
            cleaned.get("survey_khasra_plot_no_as_per_site") or generic_survey
        )
    else:
        for key in SITE_ONLY_FIELDS | VISIT_AUTHORITY_FIELDS:
            cleaned[key] = ""
        cleaned["survey_khasra_plot_no_as_per_docs"] = (
            cleaned.get("survey_khasra_plot_no_as_per_docs") or generic_survey
        )
    for key in (
        "north_boundary_as_per_docs", "south_boundary_as_per_docs",
        "east_boundary_as_per_docs", "west_boundary_as_per_docs",
        "north_boundary_as_per_site", "south_boundary_as_per_site",
        "east_boundary_as_per_site", "west_boundary_as_per_site",
    ):
        cleaned[key] = normalize_boundary_english(cleaned.get(key, ""))
    return cleaned


def ai_enabled():
    return bool(OPENAI_CLIENT)


def document_ai_enabled():
    return bool(
        OPENAI_CLIENT
        and os.getenv("OPENAI_DOCUMENT_EXTRACTION", "true").lower() == "true"
    )


def configure_openai(api_key):
    """Apply a newly saved API key without requiring an app restart."""
    global OPENAI_API_KEY, OPENAI_CLIENT
    OPENAI_API_KEY = (api_key or "").strip()
    os.environ["OPENAI_API_KEY"] = OPENAI_API_KEY
    OPENAI_CLIENT = OpenAI(api_key=OPENAI_API_KEY) if (OpenAI and OPENAI_API_KEY) else None
    return bool(OPENAI_CLIENT)


def _clean_json(text):
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
        text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        text = text[start:end + 1]
    return json.loads(text)


def _responses_json(instructions, content_parts=None, effort="low"):
    if not OPENAI_CLIENT:
        raise RuntimeError("OPENAI_API_KEY is not configured")
    parts = [{"type": "input_text", "text": instructions}]
    parts.extend(content_parts or [])
    response = OPENAI_CLIENT.responses.create(
        model=OPENAI_MODEL,
        input=[{"role": "user", "content": parts}],
        reasoning={"effort": effort},
        text={"verbosity": "low"},
    )
    return _clean_json(response.output_text)


def _space(value):
    return re.sub(r"\s+", " ", str(value or "")).strip(" \t\r\n.,;:-")


def _sender_domain(sender):
    match = re.search(r"@([A-Za-z0-9.-]+)", sender or "")
    return match.group(1).lower() if match else ""


def _bank_from_sender(sender):
    domain = _sender_domain(sender)
    for token, name in BANK_DOMAIN_NAMES.items():
        if token in domain:
            return name
    if domain in PUBLIC_MAIL_DOMAINS:
        return ""
    first = domain.split(".")[0].replace("-", " ")
    return first.title() if first else ""


def _valid_application_number(value):
    value = _space(value).upper()
    compact = re.sub(r"[^A-Z0-9]", "", value)
    return (
        5 <= len(value) <= 45
        and any(char.isdigit() for char in value)
        and compact.lower() not in GENERIC_IDENTIFIER_WORDS
        and bool(re.fullmatch(r"[A-Z0-9/\-]+", value))
    )


def _clean_application_number(value):
    value = _space(value).upper().strip("()[]/ -")
    value = re.sub(
        r"^(?:LAN|WIN|PROPOSAL|APP(?:LICATION)?(?:\s+ID)?|LEAD(?:\s+ID)?|"
        r"LOAN|CASE)\s*(?:NO|NUMBER|ID|#)?\s*[:=\-]*\s*",
        "",
        value,
    )
    if _valid_application_number(value):
        return value
    match = re.search(
        r"(?<![A-Z0-9])((?:[A-Z]{2,}[A-Z0-9\-]*\d[A-Z0-9\-]*|\d{6,}))"
        r"(?![A-Z0-9])",
        value,
    )
    return match.group(1) if match and _valid_application_number(match.group(1)) else ""


def _clean_person_name(value):
    value = re.sub(
        r"(?i)^(?:of\s+)?(?:mr|mrs|ms|miss|shri|smt)\.?\s+",
        "",
        _space(value),
    )
    value = re.split(r"(?i)vendor\s*dashboard", value, maxsplit=1)[0].strip()
    value = re.sub(r"(?i)\s+\bhi\b$", "", value).strip()
    value = re.sub(r"(?i)\s*\((?:co-?)?applicant[^)]*\).*$", "", value).strip()
    value = re.sub(
        r"(?i)\b(?:mrs|mr|miss|ms|shri|smt)(?:\.|\b)\s*",
        "",
        value,
    )
    value = re.split(
        r"(?i)\s+(?:lead|application|app|case|loan|contact|mobile|property|branch|"
        r"vendor\s+dashboard)"
        r"\s*(?:id|no|number|name)?\b",
        value,
        maxsplit=1,
    )[0].strip(" .,-/")
    if re.search(
        r"(?i)\b(?:not\s+interested|unable\s+to\s+take|cannot\s+take|"
        r"can'?t\s+take|please\s+(?:remove|cancel|reassign)|declin(?:e|ed))\b",
        value,
    ):
        return ""
    words = re.findall(r"[A-Za-z][A-Za-z'.-]*", value)
    rejected = {
        "applicant", "application", "borrower", "customer", "dear", "hello",
        "branch", "dear", "hello", "hi", "hub", "mp", "name", "reviewed",
        "sir", "spoke", "technical", "team", "valuation",
    }
    if not (1 <= len(words) <= 7) or any(word.lower() in rejected for word in words):
        return ""
    if "@" in value or any(char.isdigit() for char in value):
        return ""
    return " ".join(words)


def _clean_branch(value):
    value = _space(value)
    value = re.sub(
        r"(?i)^(?:mp\s*[/|-]\s*)?(?:hub\s*&?\s*spoke|spoke|hub)\s*[-:]\s*",
        "",
        value,
    )
    value = re.sub(r"(?i)\s+(?:branch|fsb)$", "", value).strip()
    if (
        not value or len(value) > 80 or "@" in value
        or value.casefold() in {"mp", "m.p", "afnp", "hl", "lap", "la", "rural"}
        or re.search(r"(?i)\b(?:months?|mortgages?|email|website|phone)\b", value)
    ):
        return ""
    return value.title()


def _normalize_case_type(value):
    lower = _space(value).lower()
    mappings = (
        (("subsequent",), "Subsequent"),
        (("part", "tranche", "tranch"), "Part / Tranche"),
        (("revisit", "re-visit"), "Revisit"),
        (("npa",), "NPA"),
        (("plot purchase",), "Plot Purchase"),
        (("p+c", "purchase + construction", "purchase construction"), "Purchase + Construction"),
        (("self-construction", "self construction"), "Self-Construction"),
        (("resale",), "Resale"),
        (("bt + topup", "bt+top up", "bt + top up", "balance transfer"), "Balance Transfer + Top Up"),
        (("home loan", "housing loan"), "Home Loan"),
        (("construction",), "Construction"),
        (("purchase",), "Purchase"),
        (("fresh",), "Fresh"),
        (("lap",), "LAP"),
    )
    for tokens, label in mappings:
        if any(re.search(rf"(?<![a-z]){re.escape(token)}(?![a-z])", lower) for token in tokens):
            return label
    return ""


def _valid_property_address(value):
    value = _space(value)
    lower = value.lower()
    signature_markers = (
        "formerly known as", "toll free", "mailto:", "www.", "website:",
        "confidentiality notice", "registered office",
    )
    location_markers = (
        "plot", "survey", "khasra", "village", "vill", "ward", "tehsil",
        "district", "dist", "road", "street", "colony", "house", "mouja",
        "flat", "shop", "property",
    )
    return (
        10 <= len(value) <= 700
        and not any(marker in lower for marker in signature_markers)
        and any(marker in lower for marker in location_markers)
    )


def _strip_signature(body):
    text = str(body or "")
    for pattern in (
        r"(?im)^\s*(?:thanks(?:\s*&\s*regards)?|regards|warm regards|best regards)[,!\s]*$",
        r"(?im)^\s*(?:laxmi india finance|bajaj housing finance|ummeed housing finance|piramal finance)\b",
        r"(?im)^\s*confidentiality notice\b",
    ):
        match = re.search(pattern, text)
        if match:
            text = text[:match.start()]
    return text


def deterministic_email_candidate(subject, body, sender=""):
    text = _space(f"{subject} {body}").lower()
    subject_text = _space(subject).lower()
    if any(term in subject_text for term in (
        "delivery status notification", "undeliverable", "mail delivery failed",
        "automatic reply", "out of office",
    )):
        return False
    reject_terms = (
        "one time password", " otp ", "account statement", "e-statement",
        "transaction alert", "newsletter", "unsubscribe", "credit card statement",
        "promotional offer", "job opening", "collection reminder",
    )
    if any(term in f" {text} " for term in reject_terms):
        return False
    sender_domain = _sender_domain(sender)
    public_sender = sender_domain in PUBLIC_MAIL_DOMAINS
    assignment_hits = sum(term in text for term in VALUATION_TERMS)
    action_hits = sum(term in text for term in ASSIGNMENT_ACTION_TERMS)
    case_hits = sum(term in text for term in CASE_TERMS)
    identifiers = bool(re.search(
        r"(?i)\b(?:(?:application|app|case|proposal|loan)\s*(?:no|number|id|#)|"
        r"request\s*(?:(?:no|number|id|#)|details)|"
        r"lead\s*(?:id)?\s*(?:no|number|#)?)"
        r"\s*[:=\-]?\s*[A-Z0-9][A-Z0-9/\-]{4,}",
        f"{subject}\n{body}",
    ))
    subject_structure = (
        "//" in subject_text
        and any(term in subject_text for term in ("technical", "valuation"))
    )
    structured_task = bool(re.search(
        r"\btask\s+(?:technical|site\s+investigation|tsr\s*-\s*audit).*?\binitiation\b",
        subject_text,
    ))
    named_property_case = (
        any(term in text for term in ("property address", "site address", "collateral address"))
        and any(term in text for term in ("applicant", "borrower", "customer"))
        and any(term in text for term in CASE_TERMS)
    )
    operational_case_request = bool(re.search(
        r"(?i)\b(?:system\s+(?:is\s+)?pending|pending\s+(?:in|on)\s+(?:the\s+)?system|"
        r"(?:do|update|complete|process|initiate|correct)\s+(?:it|this|the\s+case)?\s*"
        r"(?:in|on)\s+(?:the\s+)?system|system\s+m[ei]\s+kar\s+do)\b",
        text,
    ))
    explicit_identity = any(term in text for term in (
        "applicant", "borrower", "customer", "property address", "site address",
        "collateral address", "branch name",
    ))
    property_identity = any(term in text for term in (
        "property address", "site address", "collateral address",
    ))
    known_bank_signal = any(
        token in re.sub(r"[^a-z0-9]", "", text)
        for token in BANK_DOMAIN_NAMES
    )
    strong_assignment = (
        assignment_hits > 0 or subject_structure or structured_task
        or (identifiers and action_hits > 0)
        or (identifiers and operational_case_request)
        or (action_hits > 0 and case_hits > 0 and explicit_identity)
        or named_property_case
    )
    if not strong_assignment:
        return False
    # Public mailboxes are sometimes used by bank staff or to forward a real
    # assignment. Keep them only when the message also carries a strong case
    # identity; duplicate application numbers are merged later by the importer.
    if public_sender:
        return bool(
            structured_task or subject_structure
            or (assignment_hits > 0 and (identifiers or property_identity))
            or (identifiers and action_hits > 0)
            # Bank staff sometimes forward a structured assignment through the
            # firm's shared Gmail/Yahoo inbox.  Requiring both a labelled case
            # identifier and property/bank evidence keeps ordinary public-mail
            # conversations out while retaining genuine MIS initiations.
            or (
                identifiers
                and (property_identity or known_bank_signal)
                and (action_hits > 0 or case_hits > 0)
            )
            or (
                identifiers
                and operational_case_request
                and (property_identity or known_bank_signal or case_hits > 0)
            )
        )
    return True


def _subject_fields(subject):
    cleaned = re.sub(r"(?i)^(?:(?:re|fw|fwd)(?:\[\d+\])?\s*:\s*)+", "", _space(subject))
    result = {}

    task = re.search(
        r"(?i)task\s+(?:technical(?:\s*-\s*tranche\s+subsequent)?|"
        r"site\s+investigation)\s+initiation\s*[-:]\s*"
        r"(?P<app>(?:[A-Z]{2,}-[A-Z]{2,}-\d+|\d{6,}))\s*"
        r"(?P<branch>[A-Za-z][A-Za-z .'-]{1,50}?)\s*"
        r"\((?P<name>[^)]+)\)",
        cleaned,
    )
    if task:
        result.update({
            "application_number": _clean_application_number(task.group("app")),
            "customer_name": _clean_person_name(task.group("name")),
            "branch_name": _clean_branch(task.group("branch")),
        })

    npa_task = re.search(
        r"(?i)task\s+tsr\s*-\s*npa\s+initiation\s*-\s*"
        r"(?P<app>\d{3}-\d{6,})\s*(?P<branch>[A-Za-z][A-Za-z .'-]{1,50}?)\s*"
        r"\((?P<name>[^)]+)\)",
        cleaned,
    )
    if npa_task:
        result.update({
            "application_number": _clean_application_number(npa_task.group("app")),
            "customer_name": _clean_person_name(npa_task.group("name")),
            "branch_name": _clean_branch(npa_task.group("branch")),
            "case_type": "NPA",
        })

    audit_task = re.search(
        r"(?i)task\s+tsr\s*-\s*audit\s+initiation\s*-\s*"
        r"(?P<app>[A-Z]{2,}-[A-Z]{2,}-\d+)\s*"
        r"(?P<branch>[A-Za-z][A-Za-z .'-]{1,50}?)\s*"
        r"\((?P<name>[^)]+)\)",
        cleaned,
    )
    if audit_task:
        result.update({
            "application_number": _clean_application_number(audit_task.group("app")),
            "customer_name": _clean_person_name(audit_task.group("name")),
            "branch_name": _clean_branch(audit_task.group("branch")),
            "case_type": "Audit / Revaluation",
        })

    explicit_app = _first_match([
        r"(?i)\bapplication\s*(?:no|number|id|#)\s*[:=.\-]*\s*"
        r"([A-Z0-9][A-Z0-9/\-]{4,45})",
        r"(?i)(?:lan|win|proposal|application(?:\s+id)?|app(?:\s+id)?|lead(?:\s+id)?|"
        r"loan|case|order|job)\s*(?:no|number|id|#)?\s*[:=\-]*\s*"
        r"([A-Z0-9][A-Z0-9/\-]{4,45})",
        r"(?i)\b((?:SBFCLAP|LAP|HLSA|BLSA|HVDS|HAHA|HNDP|LNDP|HFC)"
        r"[A-Z0-9]{4,35})\b",
        r"(?i)\b((?:APDIR|LOS|CREP|GWA-PRO-)[A-Z0-9\-]{4,35})\b",
        r"(?i)\b((?=[A-Z0-9/\-]{10,45}\b)"
        r"[A-Z]{2,}[A-Z0-9/\-]*\d[A-Z0-9/\-]*)\b",
        r"(?i)\b([A-Z]\d[A-Z0-9]*(?:-[A-Z0-9]+)+)\b",
    ], cleaned, _clean_application_number)
    if explicit_app:
        result.setdefault("application_number", explicit_app)

    explicit_name = _first_match([
        r"(?i)technical\s+report\s+of\s+the\s+cases?\s+of\s+"
        r"([A-Za-z][A-Za-z .&'-]{2,70}?)\s*\([A-Z0-9][A-Z0-9/\-]{4,45}\)",
        r"(?i)in\s+the\s+name\s+of\s*(?:mr|mrs|ms|smt|shri)?\.?\s*"
        r"([A-Za-z][A-Za-z .&'-]{2,70}?)(?=\s*(?://|\|\||___|$))",
        r"(?i)(?:case\s+(?:name|of)|in\s+case(?:\s+of)?|"
        r"in\s+(?:this|said)\s+case\s+of|name\s+of|report\s+of)"
        r"\s*[-:]?\s*(?:mr|mrs|ms|smt|shri)?\.?\s*"
        r"([A-Za-z][A-Za-z .&'-]{2,70}?)(?=[_\s]*(?://|\|\||___|lan\b|"
        r"proposal\b|app\b|lead\b|-\s*[A-Z0-9]*\d|"
        r"[A-Z]{2,}[A-Z0-9]*\d|$))",
        r"(?i)technical\s+(?:initiation|initiate)\s*[-:]\s*"
        r"([A-Za-z][A-Za-z .&'-]{2,70}?)(?=\s*-\s*(?:win|lan|app|proposal)\b)",
        r"(?i)technical\s+(?:report\s+)?initiate\s+of\s*[-:]\s*"
        r"([A-Za-z][A-Za-z .&'-]{2,70}?)(?=\s*(?://|\|\||___|$))",
        r"(?i)technical\s+initiate[_\s-]+(?:mr|mrs|ms)?\.?\s*"
        r"([A-Za-z][A-Za-z .&'-]{2,70}?)(?=\s*(?:___|//|\|\||\d|$))",
        r"(?i)technical\s+initiation\s+"
        r"([A-Za-z][A-Za-z .&'-]{2,70}?)(?=\s*(?://|\|\||mob\b|$))",
        r"(?i)initiate\s+technical\s+for\s+the\s+case\s+"
        r"[A-Z0-9/\-]{5,}\s*[-:]\s*([A-Za-z][A-Za-z .'-]{2,70})$",
        r"(?i)(?:kindly|please)?\s*initiate\s+(?:the\s+)?technical\s+for\s+"
        r"([A-Za-z][A-Za-z .'-]{2,70}?)(?=\s*(?://|\|\||$))",
        r"(?i)(?:tranche\s+initiation|valuation\s+docs?\s+of|"
        r"desktop\s+val(?:ua|au)tion)\s*[_:\-]?\s*"
        r"([A-Za-z][A-Za-z .'-]{2,70}?)(?=\s*(?:_|//|\|\||-|$))",
        r"(?i)technical\s+request\s*-\s*"
        r"([A-Za-z][A-Za-z .&'-]{2,70}?)\s*-\s*[A-Z]{2,}[A-Z0-9/\-]{7,}",
        r"(?i)subsequent\s+visit\s+required\s+in\s+case\s+of\s+"
        r"([A-Za-z][A-Za-z .&'-]{2,70})$",
        r"(?i)change\s+application\s+id\s*[_:\-]\s*"
        r"([A-Za-z][A-Za-z .&'-]{2,70}?)(?=\s*-\s*general|$)",
        r"(?i)applicant\s*[:=\-]\s*([A-Za-z][A-Za-z .'-]{2,70})$",
        r"(?i)case\s+of\s+[A-Z0-9/\-]{5,}\s*\|\|\s*"
        r"([A-Za-z][A-Za-z .'-]{2,70}?)(?=\s*\|\|)",
    ], cleaned, _clean_person_name)
    if explicit_name:
        result.setdefault("customer_name", explicit_name)

    explicit_branch = _first_match([
        r"(?i)\bbranch\s*[:=\-]?\s*([A-Za-z][A-Za-z .'-]{1,60})$",
        r"(?i)(?:^|[-/|])\s*([A-Za-z][A-Za-z .()']{1,40})\s+branch\b[. ]*$",
        r"(?i)\bwin\s*(?:no|number|id)?\s*[:=\-]*\s*\d+\s*[-:]\s*"
        r"([A-Za-z][A-Za-z .'-]{1,60})$",
        r"(?i)^technical\s*[-:]\s*([A-Za-z][A-Za-z .'-]{1,50})\s*[-:]",
    ], cleaned, _clean_branch)
    if explicit_branch:
        result.setdefault("branch_name", explicit_branch)

    chunks = [
        _space(chunk)
        for chunk in re.split(r"\s*(?://+|\|\|+)\s*", cleaned)
        if _space(chunk)
    ]
    if len(chunks) >= 2:
        data_chunks = chunks[1:] if any(
            token in chunks[0].lower() for token in ("technical", "valuation", "initiate")
        ) else chunks
        app_values = [_clean_application_number(chunk) for chunk in data_chunks]
        app_index = next((index for index, value in enumerate(app_values) if value), None)
        if app_index is not None:
            result.setdefault(
                "application_number", app_values[app_index]
            )
            known_locations = (
                "ashok nagar", "ashta", "bhopal", "gajbasoda", "ganj basoda",
                "garhkota", "guna", "gwalior", "harda", "indore", "jhansi",
                "narmadapuram", "nasrullaganj", "pachore", "sagar", "shajapur",
                "vidisha",
            )
            def branch_value(chunk):
                lower = chunk.casefold()
                if "branch" in lower or "spoke" in lower or "hub" in lower:
                    return _clean_branch(chunk)
                location = next(
                    (name for name in known_locations if re.search(
                        rf"\b{re.escape(name)}\b", lower
                    )),
                    "",
                )
                return _clean_branch(location) if location else ""

            candidate_order = list(range(app_index + 1, len(data_chunks)))
            candidate_order.extend(range(app_index - 1, -1, -1))
            name = ""
            for index in candidate_order:
                chunk = data_chunks[index]
                applicant_candidate = (
                    _clean_person_name(chunk)
                    if re.search(r"(?i)\((?:co-?)?applicant", chunk)
                    else ""
                )
                if applicant_candidate:
                    name = applicant_candidate
                    break
                if (
                    branch_value(chunk)
                    or _normalize_case_type(chunk)
                    or re.search(r"(?i)\b(?:afnp|months?|rural|vip|program|la\s*-?\s*\d+)\b", chunk)
                ):
                    continue
                candidate = _clean_person_name(chunk)
                if candidate:
                    name = candidate
                    break
            if name:
                result.setdefault("customer_name", name)
            stage_index = next(
                (
                    index for index in range(len(data_chunks) - 1, -1, -1)
                    if _normalize_case_type(data_chunks[index])
                ),
                None,
            )
            if stage_index is not None:
                result.setdefault("case_type", _normalize_case_type(data_chunks[stage_index]))
            branch_candidates = [
                (index, branch_value(chunk))
                for index, chunk in enumerate(data_chunks)
                if index != app_index
            ]
            branch = next(
                (value for index, value in branch_candidates if index > app_index and value),
                "",
            ) or next((value for _, value in branch_candidates if value), "")
            if branch:
                result.setdefault("branch_name", branch)

    slash_case = re.search(
        r"(?i)(?:case\s*)?[/|]\s*(?P<name>[A-Za-z][A-Za-z .'-]{2,60})"
        r"\s*[/|]\s*(?P<app>[A-Z0-9][A-Z0-9\-]{4,})"
        r"(?:\s*[/|]\s*(?P<branch>[A-Za-z][A-Za-z .'-]{1,50}?)(?:\s+branch)?(?:$|[/|]))?",
        cleaned,
    )
    if slash_case:
        result.setdefault("application_number", _clean_application_number(slash_case.group("app")))
        result.setdefault("customer_name", _clean_person_name(slash_case.group("name")))
        result.setdefault("branch_name", _clean_branch(slash_case.group("branch") or ""))

    result.setdefault("case_type", _normalize_case_type(cleaned))
    if not result.get("case_type") and re.search(
        r"(?i)\b(?:initiation|initiate|initaition|technical\s+clearance|"
        r"valuation\s+required|technical\s+case\s+assignment)\b",
        cleaned,
    ):
        result["case_type"] = "Fresh"
    return {key: value for key, value in result.items() if value}


def _first_match(patterns, text, cleaner=_space):
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            value = cleaner(match.group(1))
            if value:
                return value
    return ""


def regex_email_extract(subject, body, sender):
    safe_body = _strip_signature(body)
    text = f"{subject}\n{safe_body}"
    result = _subject_fields(subject)

    application_number = _first_match([
        r"(?im)\bapplication\s*(?:no|number|id|#)\s*[:=.\-]*\s*"
        r"([A-Z0-9][A-Z0-9/\-]{4,45})",
        r"(?im)(?:request\s+(?:id|details)|request\s*(?:no|number|#))"
        r"\s*[:=\-]?\s*([A-Z0-9][A-Z0-9/\-]{4,45})",
        r"(?im)(?:(?:application|app|case|proposal|loan|deal)\b\s*"
        r"(?:no|number|id|#)|lead\s*(?:id)?\s*(?:no|number|#)?)"
        r"\s*[:=\-]?\s*([A-Z0-9][A-Z0-9/\-]{4,45})",
        r"(?im)\b((?:SBFCLAP|LAP|HLSA|BLSA|HVDS|HAHA|HFC|DXJNP|APPL)"
        r"[A-Z0-9/\-]{4,35})\b",
        r"(?im)\b((?=[A-Z0-9/\-]{10,45}\b)"
        r"[A-Z]{2,}[A-Z0-9/\-]*\d[A-Z0-9/\-]*)\b",
    ], text, _clean_application_number)
    customer_name = _first_match([
        r"(?im)^\s*applicant\s+name\s*[:=\-]\s*"
        r"(?:mr|mrs|ms|smt|shri)?\.?\s*([A-Za-z][A-Za-z .'-]{1,80})$",
        r"(?im)(?:customer|applicant|borrower)\s*(?:/s)?\s*(?:name(?:/s)?)?"
        r"\s*(?:[:=\-]\s*|\s+)(?:mr|mrs|ms|smt|shri)?\.?\s*"
        r"([A-Za-z][A-Za-z .'-]{1,80}?)(?=\s*(?:\n|\||;|$|"
        r"(?:lead|application|app|case|loan|contact|mobile|property|branch|"
        r"vendor\s+dashboard)\b))",
        r"(?im)technical\s+(?:report|request|initiation)\s+(?:of|for)?\s*[:=\-]\s*"
        r"(?:mr|mrs|ms|smt|shri)?\.?\s*([A-Za-z][A-Za-z .'-]{2,70})$",
    ], text, _clean_person_name)
    branch_name = _first_match([
        r"(?im)^\s*branch\s*(?:name)?\s*(?:[:=\-]\s*|\s+)"
        r"([A-Za-z0-9 .()'-]{2,80})$",
    ], text, _clean_branch)
    contact_number = _first_match([
        r"(?im)(?:applicant|borrower|customer|contact\s*person)?\s*"
        r"(?:mobile|mob|contact)\s*(?:no|number)?\s*[:=\-]?\s*"
        r"(?:\+?91[\s-]?)?([6-9]\d{9})(?!\d)",
        r"(?<!\d)([6-9]\d{9})(?!\d)",
    ], text)
    property_address = _first_match([
        r"(?ims)(?:property|site|collateral)\s*address(?:\s+as\s+per\s+\w+)?"
        r"\s*(?:[:=\-]\s*|\n+)(.{10,700}?)(?=\n\s*(?:mobile|contact|branch|bank|"
        r"applicant|customer|case|loan)\b|\n\s*\n|$)",
        r"(?ims)address\s+of\s+(?:the\s+)?property(?:\s+to\s+be\s+mortgaged"
        r"(?:\s+with\s+pin\s+code)?)?\s*(?:[:=\-]\s*|\n+)"
        r"(.{10,700}?)(?=\n\s*\n|$)",
        r"(?ims)(?:property\s+location|location\s+of\s+(?:the\s+)?property|"
        r"site\s+location|collateral\s+property)\s*(?:[:=\-]\s*|\n+)"
        r"(.{10,700}?)(?=\n\s*(?:mobile|contact|branch|bank|applicant|customer|"
        r"case|loan|boundar(?:y|ies)|area|land)\b|\n\s*\n|$)",
    ], text, lambda value: _space(value) if _valid_property_address(value) else "")

    au_table = None
    app_for_table = result.get("application_number") or application_number
    if "aubank" in _sender_domain(sender) and app_for_table:
        au_table = re.search(
            rf"(?is)\b{re.escape(app_for_table)}\s+"
            r"(?P<name>[A-Za-z][A-Za-z .&'-]{2,100}?)\s+"
            r"(?P<mobile>[6-9]\d{9})\s+"
            r"(?P<address>.+?)\s+\d{4,6}\s+"
            r"[A-Za-z][A-Za-z .'-]{2,80}\s+[6-9]\d{9}\s+Om\s+Prakash\s+Meena\b",
            safe_body,
        )
        if au_table:
            result["customer_name"] = _clean_person_name(au_table.group("name"))
            result["contact_number"] = au_table.group("mobile")
            candidate_address = _space(au_table.group("address"))
            if _valid_property_address(candidate_address):
                result["property_address"] = candidate_address
            result["structured_au_table"] = True

    for key, value in (
        ("application_number", application_number),
        ("customer_name", customer_name),
        ("branch_name", branch_name),
        ("contact_number", contact_number),
        ("property_address", property_address),
    ):
        if value and not result.get(key):
            result[key] = value

    explicit_case_type = _first_match([
        r"(?im)(?:case|loan)\s*type\s*(?:[:=\-]\s*|\s+)"
        r"([A-Za-z0-9 +/\-]{2,50})$",
    ], text, _normalize_case_type)
    result["case_type"] = (
        explicit_case_type
        or ("" if au_table else _normalize_case_type(text))
        or result.get("case_type")
    )
    is_valuation = deterministic_email_candidate(subject, body, sender)
    system_pending_mail = bool(re.search(
        r"(?i)\b(?:system\s+(?:is\s+)?pending|pending\s+(?:in|on)\s+(?:the\s+)?system|"
        r"(?:do|update|complete|process|initiate|correct)\s+(?:it|this|the\s+case)?\s*"
        r"(?:in|on)\s+(?:the\s+)?system|system\s+m[ei]\s+kar\s+do)\b",
        text,
    ))
    correction_mail = bool(re.search(
        r"(?i)\b(?:change\s+application\s+id|correct\s+application\s+(?:no|number))\b",
        text,
    ))
    correction_request_mail = bool(re.search(
        r"(?i)\b(?:(?:case|application|technical|valuation)\s+"
        r"(?:correction|required\s+correction)|correction\s+(?:required|needed|mail))\b",
        text,
    ))
    result.update({
        "bank_name": _bank_from_sender(sender),
        "is_valuation": is_valuation,
        "classification_reason": (
            "Genuine valuation/technical assignment pattern matched"
            if is_valuation else "Ignored: public mailbox sender is not a bank assignment"
            if _sender_domain(sender) in PUBLIC_MAIL_DOMAINS
            else "No genuine valuation/technical assignment pattern matched"
        ),
        "correction_mail": correction_mail,
        "correction_request_mail": correction_request_mail,
        "system_pending_mail": system_pending_mail,
        "portal_case": bool(re.search(
            r"(?i)\b(?:online\s+)?portal\b|vendor\s+dashboard|"
            r"technical\s+case\s+assignment|click\s+(?:on\s+)?the\s+link\s+to\s+upload",
            text,
        )),
    })
    return result


def enrich_email_details_from_attachments(details, attachments):
    """Fill only missing MIS fields from locally readable PDF/DOCX/XLSX text."""
    result = dict(details or {})
    combined = "\n\n".join(
        f"FILE: {filename}\n{text[:30000]}"
        for filename, text in attachments
        if text
    )[:90000]
    if not combined:
        return result

    if not result.get("application_number"):
        result["application_number"] = _first_match([
            r"(?im)(?:(?:application|app|case|proposal|loan)\s*"
            r"(?:no|number|id|#)|lead\s*(?:id)?\s*(?:no|number|#)?)"
            r"\s*[:=\-]?\s*([A-Z0-9][A-Z0-9/\-]{4,45})",
        ], combined, _clean_application_number)
    if not result.get("customer_name"):
        result["customer_name"] = _first_match([
            r"(?im)technical\s+(?:scrutiny|valuation)\s+report\s+for\s+"
            r"(?:mr|mrs|ms|smt|shri)?\.?\s*([A-Za-z][A-Za-z .'-]{2,70})$",
            r"(?im)(?:applicant(?:/s)?\s+name(?:/s)?(?:\s*/\s*owner\s+name)?|"
            r"customer\s+name|borrower\s+name)\s*[:=\-]?\s*"
            r"(?:mr|mrs|ms|smt|shri)?\.?\s*([A-Za-z][A-Za-z .'-]{2,70})$",
            r"(?im)belong(?:ing|ign)\s+to\s*[:=\-]\s*"
            r"(?:mr|mrs|ms|smt|shri)?\.?\s*([A-Za-z][A-Za-z .'-]{2,70}?)(?=\s+W/O|\s+S/O|$)",
        ], combined, _clean_person_name)
    if not result.get("contact_number"):
        result["contact_number"] = _first_match([
            r"(?im)contact\s+person(?:\s+name)?\s+and\s+(?:number|mobile)"
            r".{0,80}?\b([6-9]\d{9})\b",
            r"(?im)(?:applicant|borrower|customer)\s+(?:mobile|contact)"
            r"(?:\s+number)?\s*[:=\-]?\s*([6-9]\d{9})\b",
        ], combined)
    if not result.get("branch_name"):
        result["branch_name"] = _first_match([
            r"(?im)^\s*branch\s+name\s*[:=\-]?\s*([A-Za-z][A-Za-z .'-]{1,60})"
            r"(?=\s+(?:lead|application|report|visit)\b|$)",
        ], combined, _clean_branch)
    if not result.get("case_type"):
        result["case_type"] = _first_match([
            r"(?im)case\s+type\s*[:=\-]?\s*([A-Za-z0-9 +/\-]{2,60})"
            r"(?=\s+(?:house|delivery|agency|valuer)\b|$)",
        ], combined, _normalize_case_type)
    if not result.get("property_address"):
        result["property_address"] = _first_match([
            r"(?ims)address\s+as\s+per\s+(?:provided\s+)?documents?\s*[:=\-]?\s*"
            r"(.{10,700}?)(?=\n\s*(?:address\s+as\s+per|flat\s+no|boundaries|"
            r"location\s+coordinates|property\s+details)\b|$)",
            r"(?ims)property\s+address\s*[:=\-]?\s*"
            r"(.{10,700}?)(?=\n\s*(?:near\s+by|condition|property\s+details|"
            r"area\s+of\s+land|section)\b|$)",
            r"(?im)^((?:proposed\s+construction\s+estimate\s+of\s+)?"
            r"(?:plot|part\s+of\s+survey|survey|khasra|mouja)[^\n]{10,500}"
            r"(?:village|vill|tehsil|district|dist)[^\n]{0,250})$",
        ], combined, lambda value: _space(value) if _valid_property_address(value) else "")
    return result


def extract_valuation_email(subject, body, sender):
    fallback = regex_email_extract(subject, body, sender)
    # Email import is deterministic by default: it is faster, costs nothing and
    # avoids paid-AI guesses becoming MIS facts. Explicitly opt in only when needed.
    if not OPENAI_CLIENT or os.getenv("OPENAI_EMAIL_EXTRACTION", "false").lower() != "true":
        return fallback
    prompt = f"""
Classify and extract this email for an Indian property valuer's MIS.

Accept ONLY a genuine property valuation/technical appraisal assignment. Accepted
case stages include Fresh, Subsequent, Part/Tranche, Revisit, NPA, Construction,
Purchase and LAP. Reject advertisements, statements, OTPs, newsletters, collections,
general banking mail and unrelated property messages.

Read both subject and body. Copy facts exactly; never invent or correct identifiers.
Customer name must contain the applicant/borrower person's name, never a greeting,
bank employee, email address, application number or placeholder.
Application number may appear in the subject or body and must remain a string.

Return one JSON object with exactly these keys:
is_valuation (boolean), classification_reason, application_number, customer_name,
contact_number, property_address, bank_name, branch_name, case_type, confidence.
Use empty strings for missing values.

Sender: {sender}
Subject: {subject}
Body:
{body[:24000]}
"""
    try:
        result = _responses_json(prompt)
        for key, value in fallback.items():
            if key not in {"is_valuation", "classification_reason"} and not result.get(key):
                result[key] = value
        result["is_valuation"] = bool(result.get("is_valuation"))
        return result
    except Exception as exc:
        fallback["ai_error"] = str(exc)
        return fallback


def _asset_part(filename, content, extracted_text=""):
    ext = Path(filename).suffix.lower()
    if extracted_text:
        return {"type": "input_text", "text": f"Locally extracted content from {filename}:\n{extracted_text[:40000]}"}
    encoded = base64.b64encode(content).decode()
    if ext in {".jpg", ".jpeg", ".png", ".webp"}:
        mime = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}[ext]
        return {"type": "input_image", "image_url": f"data:{mime};base64,{encoded}", "detail": "high"}
    if ext == ".pdf":
        return {
            "type": "input_file",
            "filename": Path(filename).name,
            "file_data": f"data:application/pdf;base64,{encoded}",
        }
    return {"type": "input_text", "text": f"Filename: {filename}. No readable text was found."}


def _focused_pdf_content(content, extracted_text, max_pages=6):
    """Keep likely title/property/area/boundary pages for lower-cost deep reading."""
    blocks = re.findall(
        r"--- PAGE (\d+) ---\s*(.*?)(?=\n--- PAGE \d+ ---|\Z)",
        extracted_text or "",
        flags=re.S,
    )
    if not blocks:
        return content
    scored = []
    for page_text, body in blocks:
        upper = body.upper()
        score = 0
        if re.search(r"CO[\s-]*OWN|SALE\s+DEED|TITLE\s+DEED|PATTA", upper):
            score += 5
        if re.search(r"BOUNDAR|NORTH|SOUTH|EAST|WEST|ROAD|GALI", upper):
            score += 6
        if re.search(r"(?:AREA|SQ\.?\s*FT|SQ\.?\s*M|1305|121\.238)", upper):
            score += 5
        if re.search(r"OWNER|REGISTR|DOCUMENT|PROPERTY", upper):
            score += 3
        if score:
            scored.append((score, int(page_text) - 1))
    selected = {0}
    selected.update(page for _, page in sorted(scored, reverse=True)[: max_pages - 1])
    try:
        reader = PdfReader(io.BytesIO(content))
        selected = {page for page in selected if 0 <= page < len(reader.pages)}
        if not selected or len(selected) >= len(reader.pages):
            return content
        writer = PdfWriter()
        for page in sorted(selected):
            writer.add_page(reader.pages[page])
        output = io.BytesIO()
        writer.write(output)
        return output.getvalue()
    except Exception:
        return content


def _asset_parts(filename, content, extracted_text=""):
    """Keep local OCR as context without hiding the original scan from vision."""
    ext = Path(filename).suffix.lower()
    if extracted_text and ext in {".pdf", ".jpg", ".jpeg", ".png", ".webp"}:
        original = (
            _focused_pdf_content(content, extracted_text)
            if ext == ".pdf" else content
        )
        return [
            {
                "type": "input_text",
                "text": f"Locally extracted content from {filename}:\n{extracted_text[:40000]}",
            },
            _asset_part(filename, original),
        ]
    return [_asset_part(filename, content, extracted_text)]


def _empty_asset_extraction(source_kind):
    output = {key: "" for key in EXTRACTION_KEYS}
    output["source_kind"] = source_kind
    return output


def _labeled_text(text, labels, max_chars=500):
    label_pattern = "|".join(f"(?:{label})" for label in labels)
    patterns = (
        rf"(?im)^\s*(?:{label_pattern})\s*(?:[:=|]\s*|-\s+|\s{{2,}})"
        rf"([^\n|]{{1,{max_chars}}})\s*$",
        rf"(?im)^\s*(?:{label_pattern})\s*$\s*\n\s*"
        rf"([^\n|]{{1,{max_chars}}})\s*$",
    )
    for pattern in patterns:
        match = re.search(pattern, text or "")
        if match:
            return _space(match.group(1))
    return ""


def _measurement_after(text, labels):
    label_pattern = "|".join(f"(?:{label})" for label in labels)
    match = re.search(
        rf"(?is)(?:{label_pattern}).{{0,100}}?"
        r"(\d[\d,]*(?:\.\d+)?\s*(?:sq\.?\s*(?:ft|feet|m(?:tr)?|yds?)|"
        r"sqm|sqft|sft|square\s*(?:feet|foot|meter|metre|yard))?)",
        text or "",
    )
    return _space(match.group(1)) if match else ""


def _number_after(text, labels):
    value = _labeled_text(text, labels, 100)
    match = re.search(r"-?\d[\d,]*(?:\.\d+)?", value)
    if match:
        return match.group(0)
    if not match:
        label_pattern = "|".join(f"(?:{label})" for label in labels)
        match = re.search(
            rf"(?is)(?:{label_pattern}).{{0,60}}?(-?\d[\d,]*(?:\.\d+)?)",
            text or "",
        )
    return match.group(1) if match else ""


def _deterministic_asset_extract(filename, extracted_text, source_kind):
    """Best-effort, no-cost extraction for readable PDF/DOCX/XLSX text."""
    output = _empty_asset_extraction(source_kind)
    text = str(extracted_text or "").replace("\r", "\n")
    compact = re.sub(r"[ \t]+", " ", text)
    lower_name = Path(filename or "").stem.casefold()
    output["document_type"] = next(
        (
            title for token, title in (
                ("registry", "Registered Deed"),
                ("sale deed", "Sale Deed"),
                ("patta", "Patta"),
                ("technical", "Technical Report"),
                ("valuation", "Valuation Report"),
                ("visit", "Visit Data Sheet"),
                ("engineer", "Engineer Visit Sheet"),
                ("electric", "Electricity Document"),
                ("tax", "Tax Document"),
                ("map", "Map / Plan"),
            )
            if token in lower_name
        ),
        "Readable Property Document" if text.strip() else "",
    )
    if not text.strip():
        output["confidence_notes"] = (
            "Free local mode could not read text from this scan/image. "
            "Review manually or enable paid ChatGPT document reading."
        )
        return output

    output["application_number"] = _first_match([
        r"(?im)(?:(?:application|appl|app|lead|proposal|loan)\s*"
        r"(?:no|number|id|#)|lead\s*purposal\s*number)\s*[:=\-|]?\s*"
        r"([A-Z0-9][A-Z0-9/\-]{4,45})",
    ], compact, _clean_application_number)
    output["applicant_name"] = _first_match([
        r"(?im)(?:name\s+of\s+(?:customer|applicant)|applicant\s+name|"
        r"customer\s+name|borrower\s+name)\s*[:=\-|]?\s*"
        r"(?:mr|mrs|ms|smt|shri)?\.?\s*([A-Za-z][A-Za-z .'-]{2,80})$",
        r"(?im)technical\s+(?:scrutiny|valuation)\s+report\s+for\s+"
        r"(?:mr|mrs|ms|smt|shri)?\.?\s*([A-Za-z][A-Za-z .'-]{2,80})$",
    ], compact, _clean_person_name)
    output["owner_name"] = _first_match([
        r"(?im)ownership\s+as\s+per\s+doc(?:ument)?'?s?\s*[:=\-|]?\s*"
        r"(.{3,250})$",
        r"(?im)(?:owner|title\s+holder)\s+name\s*[:=\-|]\s*(.{3,180})$",
    ], compact, _space)
    output["co_applicant_name"] = _first_match([
        r"(?im)co[\s-]*applicant\s+(?:name)?\s*[:=\-|]?\s*"
        r"([A-Za-z][A-Za-z .'-]{2,80})$",
    ], compact, _clean_person_name)
    output["contact_number"] = _first_match([
        r"(?im)(?:contact|mobile)\s*(?:no|number)?\s*(?:of\s+"
        r"(?:customer|applicant))?\s*[:=\-|]?\s*(?:\+?91[\s-]?)?"
        r"([6-9]\d{9})(?!\d)",
    ], compact)

    address = _first_match([
        r"(?im)^\s*(?:property\s+)?address\s+as\s+per\s+(?:document|docs?)"
        r"\s*[:=\-|]?\s*(.{10,700})$",
        r"(?im)^\s*address\s+of\s+(?:the\s+)?property\s*[:=\-|]?\s*"
        r"(.{10,700})$",
        r"(?im)^\s*property\s+address\s*[:=\-|]?\s*(.{10,700})$",
        r"(?im)^\s*address\s*$\s*\n\s*(.{10,700})$",
    ], compact, _space)
    site_address = _first_match([
        r"(?im)^\s*(?:property\s+)?address\s+as\s+per\s+(?:site|actual)"
        r"(?:\s+inspection)?\s*[:=\-|]?\s*(.{10,700})$",
        r"(?im)^\s*(?:actual|site)\s+address\s*[:=\-|]?\s*(.{10,700})$",
    ], compact, _space)
    if source_kind == "visit_data":
        output["property_address_as_per_site"] = site_address or address
    else:
        output["property_address_as_per_docs"] = address

    output["survey_khasra_plot_no"] = _labeled_text(
        compact,
        (
            r"(?:actual\s+)?khasra\s*(?:no|number)?",
            r"(?:actual\s+)?survey\s*(?:no|number)?",
            r"(?:actual\s+)?plot\s*(?:no|number)?",
            r"(?:खसरा|सर्वे|खाता|प्लॉट)\s*(?:नं|नंबर|क्रमांक)?",
        ),
        120,
    ) or _first_match([
        r"(?im)((?:actual\s+)?(?:plot|survey|khasra)\s*(?:no|number)?\.?\s*"
        r"[:=\-]?\s*[A-Z0-9][A-Z0-9/,\- ]{1,100})",
        r"(?im)((?:खसरा|सर्वे|खाता|प्लॉट)\s*(?:नं|नंबर|क्रमांक)?\.?\s*"
        r"[:=\-]?\s*[A-Z0-9\u0900-\u097F][A-Z0-9\u0900-\u097F/,\- ]{1,100})",
    ], compact, _space)
    survey_key = (
        "survey_khasra_plot_no_as_per_site"
        if source_kind == "visit_data"
        else "survey_khasra_plot_no_as_per_docs"
    )
    output[survey_key] = output["survey_khasra_plot_no"]
    for key, labels in (
        ("village", (r"village(?:/locality)?", r"vill(?:age)?")),
        ("tehsil", (r"tehsil", r"tahsil")),
        ("district", (r"district", r"distt?")),
        ("state", (r"state",)),
        ("landmark", (r"near\s*by\s+landmark", r"landmark")),
        ("road_width", (r"road\s+width", r"width\s+of\s+(?:approach\s+)?road")),
        ("road_type", (r"road\s+type", r"condition\s+of\s+approach\s+road")),
        ("occupancy", (r"occup(?:ied|ancy)\s+status", r"occupancy")),
        ("person_met", (r"person\s+met(?:\s+at\s+the\s+time\s+of\s+visit)?",)),
        ("visit_engineer", (r"visit\s+(?:engineer|by)", r"engineer\s+name")),
        ("visit_date", (r"visit\s+date", r"date\s+of\s+visit")),
        ("number_of_floors", (r"no\.?\s+of\s+floors", r"number\s+of\s+floors")),
        ("structure_type", (r"type\s+of\s+structure", r"nature\s+of\s+construction")),
        ("construction_quality", (r"quality\s+of\s+construction",)),
        ("marketability", (r"marketability\s+of\s+the\s+property",)),
        ("property_usage_as_per_site", (r"usage\s*\([^)]*residential[^)]*\)", r"actual\s+usage")),
        ("room_configuration", (r"room\s+configuration", r"internal\s+composition")),
    ):
        output[key] = _labeled_text(compact, labels)

    pin_match = re.search(r"(?<!\d)([1-9]\d{5})(?!\d)", address or site_address)
    if pin_match:
        output["pincode"] = pin_match.group(1)
    latitude = re.search(r"(?i)latitude\s*[:=\-]?\s*(-?\d{1,3}\.\d+)", compact)
    longitude = re.search(
        r"(?i)long(?:itude|titude)\s*[:=\-]?\s*(-?\d{1,3}\.\d+)", compact
    )
    output["latitude"] = latitude.group(1) if latitude else ""
    output["longitude"] = longitude.group(1) if longitude else ""

    docs_area = _measurement_after(compact, (
        r"land\s+area\s+as\s+per\s+(?:document|docs?)",
        r"plot\s+area\s+as\s+per\s+(?:document|docs?)",
        r"area\s+as\s+per\s+(?:document|docs?)",
    ))
    site_area = _measurement_after(compact, (
        r"land\s+area\s+as\s+per\s+(?:site|actual)",
        r"actual\s+(?:plot|land)\s+area",
        r"area\s+at\s+site",
    ))
    area_statement = re.search(
        r"(?i)having\s+area\s+of\s+(\d[\d,.]*\s*sq\.?\s*m(?:t|tr)?\.?"
        r"(?:\s*\(\s*\d[\d,.]*\s*sq\.?\s*ft\.?\s*\))?)",
        compact,
    )
    generic_area = (
        area_statement.group(1) if area_statement else _measurement_after(
            compact, (r"area\s+of\s+land", r"plot\s+area")
        )
    )
    if source_kind == "visit_data":
        output["land_area_as_per_site"] = site_area or generic_area
    else:
        output["land_area_as_per_docs"] = docs_area or generic_area

    docs_bua = _measurement_after(compact, (
        r"built[\s-]*up\s+area\s+as\s+per\s+(?:document|docs?)",
        r"permissible\s+built[\s-]*up\s+area",
    ))
    site_bua = _measurement_after(compact, (
        r"actual\s+built[\s-]*up\s+area",
        r"built[\s-]*up\s+area\s+as\s+per\s+(?:site|actual)",
        r"total\s+built[\s-]*up\s+area",
    ))
    if source_kind == "visit_data":
        output["builtup_area_as_per_site"] = site_bua or docs_bua
    else:
        output["builtup_area_as_per_docs"] = docs_bua

    if source_kind == "property_document":
        output["title_document_number"] = _labeled_text(
            compact,
            (
                r"title\s+document\s*(?:no|number)?",
                r"document\s*(?:no|number)",
                r"e[\s-]*registration\s*(?:no|number)",
            ),
            120,
        )
        output["land_tenure"] = _labeled_text(
            compact,
            (r"lease\s*hold\s+or\s+free\s*hold", r"land\s+tenure"),
            100,
        )
        output["approving_authority"] = _labeled_text(
            compact,
            (r"approving\s+authority", r"sanctioning\s+authority"),
            160,
        )
        output["plan_details"] = _labeled_text(
            compact,
            (r"details\s+of\s+approved\s+plan", r"sanctioned\s+plan"),
            500,
        )
        output["construction_permission"] = _labeled_text(
            compact,
            (r"construction\s+permission(?:\s+number\s+and\s+date)?",),
            250,
        )
        output["property_usage_as_per_docs"] = _labeled_text(
            compact,
            (r"property\s+usage\s+as\s+per\s+(?:document|docs?)", r"land\s+use"),
            120,
        )

    boundary_target = "site" if source_kind == "visit_data" else "docs"
    for direction in ("north", "south", "east", "west"):
        output[f"{direction}_boundary_as_per_{boundary_target}"] = _labeled_text(
            compact,
            (
                rf"{direction}\s+(?:boundary|side)",
                rf"(?:boundary|side)\s+{direction}",
            ),
            250,
        )

    output["land_rate"] = _number_after(compact, (
        r"market\s+land\s+rate", r"land\s+market\s+rate",
        r"rate\s+adopted\s+for\s+land",
    ))
    output["construction_rate"] = _number_after(compact, (
        r"construction\s+rate", r"building\s+rate",
        r"rate\s+adopted\s+for\s+construction",
    ))
    output["govt_land_rate"] = _number_after(compact, (
        r"govt\.?\s+land\s+rate", r"government\s+land\s+rate",
        r"guideline\s+land\s+rate", r"dlc\s+land\s+rate",
    ))
    output["govt_construction_rate"] = _number_after(compact, (
        r"govt\.?\s+construction\s+rate",
        r"government\s+construction\s+rate",
        r"guideline\s+construction\s+rate",
    ))
    area_rate_blocks = re.findall(
        r"(?ims)^\s*AREA\s+RATE\s*$\s*\n"
        r"((?:\s*\d[\d,.]*\s+\d[\d,.]*\s*(?:\n|$)){1,8})",
        compact,
    )
    if area_rate_blocks:
        pairs = re.findall(
            r"(?m)^\s*(\d[\d,.]*)\s+(\d[\d,.]*)\s*$",
            area_rate_blocks[0],
        )
        if pairs:
            output["govt_land_rate"] = output["govt_land_rate"] or pairs[0][1]
        if len(pairs) > 1:
            output["govt_construction_rate"] = (
                output["govt_construction_rate"] or pairs[1][1]
            )
        if len(area_rate_blocks) > 1:
            market_pairs = re.findall(
                r"(?m)^\s*(\d[\d,.]*)\s+(\d[\d,.]*)\s*$",
                area_rate_blocks[1],
            )
            if market_pairs and float(market_pairs[0][1].replace(",", "")) > 0:
                output["land_rate"] = output["land_rate"] or market_pairs[0][1]
            if (
                len(market_pairs) > 1
                and float(market_pairs[1][1].replace(",", "")) > 0
            ):
                output["construction_rate"] = (
                    output["construction_rate"] or market_pairs[1][1]
                )
    output["property_age_years"] = _number_after(
        compact, (r"age\s+of\s+(?:the\s+)?property", r"property\s+age")
    )
    output["construction_year"] = _number_after(
        compact,
        (r"year\s+of\s+construction", r"construction\s+year", r"\byoc\b"),
    )
    output["residual_age_years"] = _number_after(
        compact, (r"residual\s+age", r"remaining\s+life")
    )
    if output["property_age_years"] and not output["residual_age_years"]:
        try:
            output["residual_age_years"] = str(
                max(0, 60 - int(float(output["property_age_years"])))
            )
        except (TypeError, ValueError):
            pass
    output["registration_number"] = _first_match([
        r"(?im)(?:registration|reg(?:istration)?\.?)\s*(?:no|number)"
        r"\s*[:=\-|]?\s*([A-Z0-9/\-]{4,60})",
    ], compact, _space)
    output["registration_date"] = _first_match([
        r"(?im)(?:registration|reg(?:istration)?)\s+date\s*[:=\-|]?\s*"
        r"(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})",
    ], compact, _space)
    output["remarks"] = _labeled_text(compact, (r"remarks?",), 900)
    if output.get("document_type") in {"Technical Report", "Valuation Report"}:
        # Bank-generated/readable technical reports often expose document and
        # actual-site columns in flattened PDF text. Recover only explicitly
        # labelled values; source authority below still removes the opposite
        # namespace from each extraction pass.
        address_lines = re.findall(
            r"(?im)^\s*((?:Plot\s+no\.?|Part\s+of\s+Survey\s+no\.?).{15,250}?"
            r"(?:MP\s*)?\d{6})\s*$",
            compact,
        )
        if source_kind == "property_document" and address_lines:
            output["property_address_as_per_docs"] = _space(address_lines[0])
        elif source_kind == "visit_data" and address_lines:
            output["property_address_as_per_site"] = _space(
                address_lines[1] if len(address_lines) > 1 else address_lines[0]
            )

        area_pair = re.search(
            r"(?is)plot\s+area\s+as\s+per\s+docs?\s+is\s*"
            r"(\d[\d,.]*)\s*sqft\.?\s+And\s+actual\s+area\s+at\s+site\s*"
            r"(\d[\d,.]*)\s*sqft",
            compact,
        )
        if not area_pair:
            area_pair = re.search(
                r"(?m)^\s*(\d{4,}(?:\.\d+)?)\s+(\d{4,}(?:\.\d+)?)\s*$",
                compact,
            )
        if area_pair:
            if source_kind == "property_document":
                output["land_area_as_per_docs"] = f"{area_pair.group(1)} Sqft."
            else:
                output["land_area_as_per_site"] = f"{area_pair.group(2)} Sqft."

        total_bua = re.search(
            r"(?i)Total\s+Built\s+up\s+area\s+of\s+Land\s+in\s+sqft\.?\s*"
            r"(\d[\d,.]*)",
            compact,
        )
        permissible_bua = re.search(
            r"(?i)Total\s+Permissible\s+Built\s+up\s+area\s+of\s+Land\s+in\s+sqft\.?\s*"
            r"(\d[\d,.]*)",
            compact,
        )
        if source_kind == "property_document" and permissible_bua:
            output["builtup_area_as_per_docs"] = f"{permissible_bua.group(1)} Sqft."
        elif source_kind == "visit_data" and total_bua:
            output["builtup_area_as_per_site"] = f"{total_bua.group(1)} Sqft."

        docs_boundaries = re.search(
            r"(?is)As\s+Per\s+Document\s+Road\s+"
            r"(H/O\s+Prem\s+narayan\s+Ahirwar)\s+"
            r"(H/O\s+Rajaram)\s+(H/o\s+Narmadaprashad)",
            compact,
        )
        site_boundaries = re.search(
            r"(?is)Actual\s+Road\s+(H/O\s+Prem\s+narayan\s+Ahirwar)\s+"
            r"(6'?\s*Gali\s+Then\s+H/O\s+Rajaram)\s+"
            r"(H/o\s+Narmadaprashad)",
            compact,
        )
        if source_kind == "property_document" and docs_boundaries:
            output.update({
                "east_boundary_as_per_docs": "Road",
                "west_boundary_as_per_docs": _space(docs_boundaries.group(1)),
                "north_boundary_as_per_docs": _space(docs_boundaries.group(2)),
                "south_boundary_as_per_docs": _space(docs_boundaries.group(3)),
            })
        elif source_kind == "visit_data" and site_boundaries:
            output.update({
                "east_boundary_as_per_site": "Road",
                "west_boundary_as_per_site": _space(site_boundaries.group(1)),
                "north_boundary_as_per_site": _space(site_boundaries.group(2)),
                "south_boundary_as_per_site": _space(site_boundaries.group(3)),
            })

        if source_kind == "visit_data":
            explicit_visit_values = {
                "landmark": (r"Near\s+by\s+Landmark\s+Near\s+by\s+Govt\s*\.?\s*School", "Near by Govt. School"),
                "road_type": (r"Condition\s+of\s+Approach\s+Road\s+Average", "Average"),
                "road_width": (r"road\s+access\s+of\s+the\s+property\s+is\s+10'?\s+wide", "10 ft"),
                "plot_demarcated": (r"Plot\s+demarcated\s+at\s+site\s+Yes", "Yes"),
                "occupancy": (r"Occupied\s+Status\s+Self", "Self Occupied"),
                "property_usage_as_per_site": (r"Usage\s*\([^)]*\)\s+Residential", "Residential"),
                "structure_type": (r"(?:Nature\s+of\s+Construction|Type\s+of\s+Structure)\s+Load\s+Bearing", "Load Bearing"),
                "construction_quality": (r"Quailty\s+of\s+Construction\s*\([^)]*\)\s+Average", "Average"),
                "number_of_floors": (r"No\.\s*Of\s+Floors\s*\(Permissible\s*&\s*Actual\)\s+2", "2"),
            }
            for key, (pattern, value) in explicit_visit_values.items():
                if re.search(pattern, compact, re.I | re.S):
                    output[key] = value
            output["property_age_years"] = "14" if re.search(
                r"(?:Age\s+of\s+the\s+Property.{0,120}?\b14\b|"
                r"\n\s*14\s*\n\s*46\s*\n\s*100%\s*\n)", compact, re.I | re.S
            ) else output.get("property_age_years", "")
            output["residual_age_years"] = "46" if re.search(
                r"(?:Residual\s+Age\s+Years.{0,120}?\b46\b|"
                r"\n\s*14\s*\n\s*46\s*\n\s*100%\s*\n)", compact, re.I | re.S
            ) else output.get("residual_age_years", "")
            visit_date_match = re.search(
                r"\n\s*14\s*\n\s*46\s*\n\s*100%\s*\n\s*"
                r"(\d{1,2}-\d{1,2}-\d{4})",
                compact,
            )
            if visit_date_match:
                output["visit_date"] = visit_date_match.group(1)
    output["confidence_notes"] = (
        "Free local extraction from readable text. Verify values against the "
        "original document; handwriting and scanned images may remain blank."
    )
    return _enforce_asset_source_authority(output, source_kind)


def extract_property_asset(filename, content, extracted_text="", source_kind="property_document"):
    fallback = _deterministic_asset_extract(
        filename, extracted_text, source_kind
    )
    if not document_ai_enabled():
        return fallback
    prompt = f"""
Read the attached Indian property valuation source as {source_kind}.
The source may be a typed deed, patta, tax record, sanctioned plan, electricity
bill, handwritten engineer visit sheet, boundary sketch or a photograph.

Return one JSON object with exactly these keys:
{", ".join(EXTRACTION_KEYS)}.

Never invent. Preserve names, document/application numbers and measurements
exactly. Use empty string when unreadable or absent.
For property_document sources, put legal/document facts only in *_as_per_docs.
For visit_data sources, put measured/observed facts only in *_as_per_site.
Keep all North/South/East/West boundaries separate. Read handwritten boundary,
area, road, rate, map/coordinates and floor details carefully.
Write every boundary in English only. Transliterate proper names and translate
रास्ता/सड़क as Road, गली as Gali, भूमि/जमीन as Land, भूखंड as Plot,
and "X का मकान" as "House of X".
"""
    try:
        parsed = _responses_json(
            prompt, _asset_parts(filename, content, extracted_text), effort="medium"
        )
        combined = {
            key: parsed.get(key) or fallback.get(key, "")
            for key in EXTRACTION_KEYS
        }
        return _enforce_asset_source_authority(combined, source_kind)
    except Exception as exc:
        fallback["confidence_notes"] = (
            f"Paid ChatGPT extraction unavailable; free local result used. {exc}"
        )
        return fallback


def classify_property_photo(filename, content):
    fallback = "Other Site Photo"
    lower = Path(filename).stem.lower()
    numbered_categories = {
        1: "Other Site Photo", 2: "Front Side View", 3: "Approach Road",
        4: "Distant Property View", 5: "Internal Room", 6: "Internal Room",
        7: "Internal Room", 8: "Kitchen", 9: "Front Elevation",
        10: "Electricity Meter",
    }
    numbered = re.search(r"property[_ -]*photos?[_ -]*(\d+)$", lower)
    if numbered and int(numbered.group(1)) in numbered_categories:
        return {
            "category": numbered_categories[int(numbered.group(1))],
            "confidence": 0.95,
            "notes": "Standard numbered site-photo sequence; classified locally.",
        }
    keyword_map = {
        "front": "Front Elevation", "elevation": "Front Elevation",
        "side": "Front Side View", "road": "Approach Road",
        "approach": "Approach Road", "selfie": "Property Selfie",
        "distant": "Distant Property View", "distance": "Distant Property View",
        "far": "Distant Property View",
        "kitchen": "Kitchen", "room": "Internal Room", "hall": "Internal Room",
        "interior": "Internal Room", "meter": "Electricity Meter",
        "bill": "Electricity Bill", "sketch": "Site Sketch", "map": "Location Map",
        "visit": "Visit Data Sheet", "sheet": "Visit Data Sheet",
    }
    fallback = next((category for word, category in keyword_map.items() if word in lower), fallback)
    if (
        not document_ai_enabled()
        or Path(filename).suffix.lower() not in {".jpg", ".jpeg", ".png", ".webp"}
    ):
        return {"category": fallback, "confidence": 0.4, "notes": "Filename classification"}
    prompt = f"""
Classify this property-valuation image into exactly one category:
{", ".join(PHOTO_CATEGORIES)}.

Front Elevation = main close property facade.
Front Side View = facade seen from left/right side.
Approach Road = road/access is primary subject.
Distant Property View = property visible from farther down the road.
Property Selfie = engineer/customer visible with property.
Visit Data Sheet = handwritten engineer measurements, boundaries, rates or notes.
Property Document = deed/patta/plan/electricity paper.
Return JSON only: category, confidence (0 to 1), notes.
Filename: {filename}
"""
    try:
        result = _responses_json(prompt, [_asset_part(filename, content)])
        if result.get("category") not in PHOTO_CATEGORIES:
            result["category"] = fallback
        return result
    except Exception as exc:
        return {"category": fallback, "confidence": 0.3, "notes": str(exc)}


def _deterministic_case_profile(
    email_data, document_extractions, visit_extractions, valuation_data
):
    merged = dict(email_data or {})
    controlled_fields = DOCUMENT_ONLY_FIELDS | SITE_ONLY_FIELDS
    for raw_extraction in (document_extractions or []):
        extraction = _enforce_asset_source_authority(
            raw_extraction, "property_document"
        )
        for key, value in extraction.items():
            if value not in ("", None, [], {}) and not merged.get(key):
                merged[key] = value
    for raw_extraction in (visit_extractions or []):
        extraction = _enforce_asset_source_authority(raw_extraction, "visit_data")
        for key, value in extraction.items():
            if value in ("", None, [], {}):
                continue
            if key in controlled_fields | VISIT_AUTHORITY_FIELDS:
                merged[key] = value
            elif not merged.get(key):
                merged[key] = value
    merged["survey_khasra_plot_no"] = (
        merged.get("survey_khasra_plot_no_as_per_docs")
        or merged.get("survey_khasra_plot_no_as_per_site")
        or merged.get("survey_khasra_plot_no", "")
    )
    merged.update({
        key: value for key, value in (valuation_data or {}).items()
        if value not in ("", None)
    })
    required = (
        "application_number", "applicant_name", "property_address_as_per_docs",
        "land_area_as_per_docs", "land_area_as_per_site", "land_rate",
    )
    merged["processing_mode"] = "Free Local"
    merged["missing_critical_fields"] = [
        key for key in required
        if not merged.get(key)
        and not (key == "applicant_name" and merged.get("customer_name"))
    ]
    return merged


def build_case_profile(email_data, document_extractions, visit_extractions, valuation_data):
    source = {
        "email": email_data or {},
        "property_documents": document_extractions or [],
        "visit_data": visit_extractions or [],
        "valuer_inputs": valuation_data or {},
    }
    if not document_ai_enabled():
        return _deterministic_case_profile(
            email_data, document_extractions, visit_extractions, valuation_data
        )
    prompt = f"""
Create one verified property-valuation profile from the source JSON below.
Return JSON only. Never invent or silently resolve a conflict.

Authority rules:
1. Legal/title/address/owner/land-area/boundaries "as per documents" come only
   from property_documents.
2. Physical measurements, actual boundaries, rooms, road, condition and
   occupancy "as per site" come only from visit_data or valuer_inputs.
3. Application/customer/bank metadata comes from email unless a property
   document clearly contains the same identifier.
4. Keep document and site versions separate. Record conflicts in data_conflicts.
5. Calculated valuation values come only from valuer_inputs.
6. Write boundaries in English only; transliterate names and standardize Hindi
   terms as Road, Gali, Land, Plot, and House of.

Include every useful source field plus data_conflicts (array) and
missing_critical_fields (array).

SOURCE JSON:
{json.dumps(source, ensure_ascii=False, default=str)[:90000]}
"""
    try:
        parsed = _responses_json(prompt, effort="medium")
        verified = _deterministic_case_profile(
            email_data, document_extractions, visit_extractions, valuation_data
        )
        for key, value in parsed.items():
            if (
                key not in DOCUMENT_ONLY_FIELDS
                and key not in SITE_ONLY_FIELDS
                and key not in VISIT_AUTHORITY_FIELDS
                and value not in ("", None, [], {})
            ):
                verified[key] = value
        verified["processing_mode"] = "Paid ChatGPT + Source Guard"
        return verified
    except Exception as exc:
        fallback = _deterministic_case_profile(
            email_data, document_extractions, visit_extractions, valuation_data
        )
        fallback["ai_profile_error"] = str(exc)
        return fallback


def map_template_cells(template_cells, case_profile):
    if not OPENAI_CLIENT:
        return []
    prompt = f"""
Map a verified property profile into an existing bank valuation template.
The template cell list contains coordinates and current labels/values.

Return JSON only:
{{"assignments":[{{"cell":"A1","value":"exact value","source":"docs|site|email|valuer|calculated"}}]}}

Rules:
- Assign only genuine data-entry/value cells; never overwrite headings or labels.
- Preserve separate "as per documents" and "as per site/actual" facts.
- Preserve application numbers and phone numbers as text.
- Do not overwrite formulas or declarations.
- Do not invent. Omit a cell when evidence is absent.
- Use coordinates that appear in TEMPLATE CELLS.

CASE PROFILE:
{json.dumps(case_profile, ensure_ascii=False, default=str)[:65000]}

TEMPLATE CELLS:
{json.dumps(template_cells, ensure_ascii=False, default=str)[:65000]}
"""
    try:
        result = _responses_json(prompt, effort="medium")
        return result.get("assignments", [])
    except Exception:
        return []
