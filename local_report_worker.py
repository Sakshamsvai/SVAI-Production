import io
import json
import os
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from pypdf import PdfReader

from ai_service_openai import (
    build_case_profile,
    classify_property_photo,
    extract_property_asset,
)
from report_service import fill_docx_template, fill_excel_template


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 250 * 1024 * 1024


@app.before_request
def restrict_browser_origin():
    origin = request.headers.get("Origin", "")
    allowed = (
        not origin
        or origin == "https://svai-valuation-app.onrender.com"
        or origin.startswith("http://127.0.0.1")
    )
    if not allowed:
        return jsonify({"error": "Origin not allowed."}), 403


@app.after_request
def allow_svai_browser(response):
    origin = request.headers.get("Origin", "")
    if origin == "https://svai-valuation-app.onrender.com" or origin.startswith("http://127.0.0.1"):
        response.headers["Access-Control-Allow-Origin"] = origin
        response.headers["Vary"] = "Origin"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    response.headers["Access-Control-Allow-Private-Network"] = "true"
    response.headers["Access-Control-Expose-Headers"] = "Content-Disposition"
    return response


@app.route("/health", methods=["GET", "OPTIONS"])
def health():
    return jsonify({"status": "ok", "mode": "SVAI Local Report Worker"})


def pdf_photos(filename, content, asset_type):
    if Path(filename).suffix.lower() != ".pdf":
        return []
    name = Path(filename).stem.casefold()
    visit_source = asset_type == "visit_data" or any(
        token in name for token in ("visit", "inspection", "site_data", "site data")
    )
    if not visit_source:
        return []
    output = []
    reader = PdfReader(io.BytesIO(content))
    for page_index, page in enumerate(reader.pages):
        images = list(page.images)
        if not images:
            continue
        largest = max(images, key=lambda item: len(item.data or b""))
        if len(largest.data or b"") < 10_000:
            continue
        category = "Front Elevation" if page_index == 0 else "Other Site Photo"
        output.append({
            "filename": f"{Path(filename).stem}_page_{page_index + 1}_{largest.name}",
            "category": category,
            "content": largest.data,
        })
    return output


def save_to_local_reports(content, filename):
    downloads = Path(os.getenv(
        "SVAI_LOCAL_REPORT_DIR",
        str(Path(__file__).resolve().parent / "Generated Reports Local"),
    ))
    downloads.mkdir(parents=True, exist_ok=True)
    safe_name = Path(filename).name or "SVAI_Local_Report.xlsx"
    target = downloads / safe_name
    counter = 2
    while target.exists():
        target = downloads / f"{Path(safe_name).stem} ({counter}){Path(safe_name).suffix}"
        counter += 1
    target.write_bytes(content)
    return target


@app.route("/generate", methods=["POST", "OPTIONS"])
def generate():
    if request.method == "OPTIONS":
        return ("", 204)
    profile = json.loads(request.form.get("profile", "{}"))
    manifest = json.loads(request.form.get("manifest", "[]"))
    template = request.files.get("template")
    if not template or not template.filename:
        return jsonify({"error": "Bank template missing."}), 400
    template_content = template.read()
    photos = []
    document_extractions = []
    visit_extractions = []
    use_paid_ai = request.form.get("paid_ai") == "1"
    source_dir = Path(__file__).resolve().parent / "Generated Reports Local" / (
        f"{Path(request.form.get('report_name') or 'SVAI_Report').stem} Sources"
    )
    source_dir.mkdir(parents=True, exist_ok=True)
    for item in manifest:
        upload = request.files.get(item["field"])
        if not upload:
            continue
        content = upload.read()
        asset_type = item.get("asset_type", "")
        extension = Path(item.get("filename", "")).suffix.lower()
        filename = item.get("filename") or upload.filename
        (source_dir / f"{item.get('field', 'asset')}_{Path(filename).name}").write_bytes(content)
        if use_paid_ai and asset_type in {"document", "visit_data"}:
            extraction = extract_property_asset(
                filename, content, "",
                "visit_data" if asset_type == "visit_data" else "property_document",
            )
            if asset_type == "visit_data":
                visit_extractions.append(extraction)
            else:
                document_extractions.append(extraction)
        filename_key = Path(filename).stem.casefold().replace("_", " ").replace("-", " ")
        category_key = str(item.get("category") or "").casefold()
        is_google_map = "google map" in filename_key or "google map" in category_key
        is_mp_kisan = any(token in filename_key or token in category_key for token in (
            "mp kisan", "mp kishan", "kisan app", "kishan app",
        ))
        is_report_visit_image = is_google_map or is_mp_kisan
        if asset_type == "photo" or (
            asset_type == "visit_data"
            and extension in {".jpg", ".jpeg", ".png", ".webp"}
            and is_report_visit_image
        ):
            photo_category = item.get("category") or "Other Site Photo"
            if use_paid_ai and asset_type == "photo":
                photo_category = classify_property_photo(filename, content).get(
                    "category", photo_category
                )
            photos.append({
                "filename": filename,
                "category": (
                    "Google Map" if is_google_map else
                    "MP Kisan" if is_mp_kisan else
                    photo_category
                ),
                "content": content,
            })
    if use_paid_ai:
        profile = build_case_profile(
            profile, document_extractions, visit_extractions, {}
        )
    extension = Path(template.filename).suffix.lower()
    if extension == ".docx":
        output = fill_docx_template(
            template_content, profile, photos,
            template_name=template.filename,
            bank_name=profile.get("bank_name", ""),
        )
        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    else:
        output = fill_excel_template(
            template_content, profile, photos,
            template_name=template.filename,
            bank_name=profile.get("bank_name", ""),
        )
        if extension == ".xlsm":
            mime = "application/vnd.ms-excel.sheet.macroEnabled.12"
        else:
            extension = ".xlsx"
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    filename = request.form.get("report_name") or f"SVAI_Local_Report{extension}"
    saved_path = save_to_local_reports(output, filename)
    response = send_file(
        io.BytesIO(output), as_attachment=True, download_name=filename, mimetype=mime
    )
    response.headers["X-SVAI-Local-Saved"] = str(saved_path)
    response.headers["Access-Control-Expose-Headers"] = (
        "Content-Disposition, X-SVAI-Local-Saved"
    )
    return response


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.getenv("SVAI_LOCAL_WORKER_PORT", "8765")), threaded=False)
