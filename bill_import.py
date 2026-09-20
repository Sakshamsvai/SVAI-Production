"""Conservative bill label extraction and bounded, in-memory ZIP input."""
import hashlib
import io
import os
import re
import zipfile
from pathlib import PurePosixPath


def labeled_value(lines, labels):
    for index, line in enumerate(lines):
        # Match labels at the start, never substring matches in service text.
        for label in sorted(labels, key=len, reverse=True):
            pattern = r"^\s*" + r"[\s._/-]*".join(map(re.escape, label))
            match = re.match(pattern + r"\b\s*[:=]?\s*(.*)$", line, re.I)
            if match:
                value = match.group(1).strip(" :")
                if not value and index + 1 < len(lines):
                    candidate = lines[index + 1].strip()
                    if label in {"invoicenumber", "invoiceno", "billnumber", "billno"}:
                        value = valid_invoice(candidate)
                    elif label in {"invoicedate", "billdate", "date"}:
                        if re.fullmatch(r"[\d/\- .A-Za-z]+", candidate) and re.search(r"\d", candidate):
                            value = candidate
                    elif re.fullmatch(r"[₹\d,. ]+", candidate):
                        value = candidate
                if value:
                    # A GST rate is not the tax amount: CGST @ 9% : 900.
                    if label in {"cgst", "sgst", "igst", "cgstamount", "sgstamount", "igstamount"}:
                        value = re.sub(r"@?\s*\d+(?:\.\d+)?\s*%", "", value).strip(" :")
                    return value.split(" : ")[0].strip()
    return ""


def valid_invoice(value):
    value = str(value or "").strip()
    if value.upper().startswith("REVIEW-"):
        return ""
    if not re.search(r"\d", value) or len(value) > 120:
        return ""
    if re.search(r"details|service|description|taxable|total|date", value, re.I):
        return ""
    return value


def bill_inputs(uploads, read_upload):
    max_files = int(os.getenv("MAX_BILL_IMPORT_FILES", "2000"))
    max_total = int(os.getenv("MAX_BILL_IMPORT_MB", "1024")) * 1024 * 1024
    max_single = int(os.getenv("MAX_BILL_FILE_MB", "100")) * 1024 * 1024
    count = total = 0
    for upload in uploads:
        content = read_upload(upload)
        if PurePosixPath(upload.filename).suffix.lower() != ".zip":
            count += 1
            total += len(content)
            if count > max_files or total > max_total or len(content) > max_single:
                raise ValueError("Billing import size limit exceeded; upload a smaller batch.")
            yield upload.filename, content
            continue
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                members = archive.infolist()
                if len(members) > max_files * 3:
                    raise ValueError("ZIP contains too many entries.")
                for member in members:
                    path = PurePosixPath(member.filename.replace("\\", "/"))
                    if member.is_dir() or path.suffix.lower() not in {".pdf", ".xlsx"}:
                        continue
                    if path.is_absolute() or ".." in path.parts or member.flag_bits & 1:
                        raise ValueError("Unsafe or password-protected ZIP entry.")
                    count += 1
                    total += member.file_size
                    if count > max_files or total > max_total or member.file_size > max_single:
                        raise ValueError("Billing import size limit exceeded; upload a smaller batch.")
                    yield str(path), archive.read(member)
        except (zipfile.BadZipFile, RuntimeError) as exc:
            raise ValueError("ZIP could not be read.") from exc


def review_identifier(content):
    return "REVIEW-" + hashlib.sha256(content).hexdigest()[:24]
