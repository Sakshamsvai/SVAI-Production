# SVAI Production — OpenAI Edition

Production deployment: https://svai-valuation-app.onrender.com

SVAI is an end-to-end property valuation workflow for Saksham Associate:

Gmail/Yahoo → valuation-only MIS → visit ZIP/photos → property documents →
free local extraction (optional paid ChatGPT) → manual valuation review →
original bank Excel/Word report.

All AI work uses OpenAI through the Responses API; no second AI provider is
configured.

## What is ready

- Gmail and Yahoo linking with encrypted 16-character app passwords. The
  connection is saved, so Fetch does not ask for the code again.
- Forgot Password sends a six-digit, 10-minute reset code through a previously
  linked Gmail/Yahoo mailbox.
- Email date range defaults to the first day of the current month through today.
- User-selectable From/To date filters and filtered MIS Excel download.
- Valuation-only email classification for Fresh, Subsequent, Part/Tranche,
  Revisit, NPA, LAP, Construction, Purchase and related technical cases.
- Customer name and application number extraction from email subject and body.
- Separate, non-ambiguous uploads for:
  - Property Documents (registry, patta, title papers and legal maps);
  - Visit Form (1/2/3/4/5 handwritten engineer pages and MP Kisan screenshots
    selected together);
  - Site Photos (elevation, side, road, selfie, kitchen and interiors);
  - original bank valuation format in XLSX, XLSM or DOCX.
- Free Local Mode separates readable "as per documents" facts from "actual at
  site" visit facts. A source-wise correction screen keeps document address,
  khasra, areas and boundaries separate from actual visit values. Paid ChatGPT
  reading can be enabled later for scans/photos.
- Photo classification for front elevation, side views, road/distant view,
  property selfie, kitchen, interiors, meter/bill, sketch and map.
- Seeded DCB, SBFC, Laxmi India and Ummeed report formats.
- Original MIS format with the operational 13 columns.
- Manual review/correction before the final report is sent to the bank.

## Local setup

1. Run the easy setup (no PowerShell input is required):

```powershell
powershell -ExecutionPolicy Bypass -File .\START_LOCAL.ps1 -EasyStart
```

The script generates local secrets, installs dependencies and opens the app.
After login, use **Settings** in the top menu to save the OpenAI API key and
choose a new SVAI password. The key is never displayed back in the browser.

Open `http://127.0.0.1:8000`.

Default local login:

- Email: `sakshamvaluer@yahoo.com`
- Temporary password: `ChangeMe123!`

Change the password immediately after the first login.

## Dashboard and real-data workflow

- The highlighted **Dashboard / MIS** button is available on every logged-in page.
- MIS uses only genuine valuation/technical assignment emails; statements, OTPs,
  promotions and unrelated attachments are ignored.
- Existing imported emails are corrected again whenever **Fetch Valuation Emails**
  is used, so improved subject/body/attachment rules fix old rows too.
- Report reminders and follow-up mails (for example, "Please share report") are
  attached to the existing application and never create another MIS row. Genuine
  Subsequent, Revisit and Part/Tranche assignments remain separate cases.
- A manual report file now starts with only the application number. The next page
  has separate Property Documents, Visit Form, Site Photos and bank-format
  uploads so handwritten pages cannot be confused with property photos.
- Online-portal cases can be kept in **Portal Pending** without a fake local report.
- Report generation requires the bank's uploaded original XLSX/XLSM/DOCX format.
  SVAI fills that same workbook/document and never falls back to a generic report.
- Excel sheet names, merged cells, row heights, column widths and print area are
  guarded; if safe in-place filling cannot be confirmed, no broken report is emitted.

ChatGPT Free/Plus and the OpenAI API are separate services. Without API billing,
SVAI still performs MIS, typed/readable PDF/DOCX/XLSX extraction, manual case
data, uploads, calculations and exact-format reports. Scanned images and
handwriting stay blank for manual review until paid document reading is enabled.

## Gmail / Yahoo setup

Enable two-step verification on the mailbox, generate an app password, and link
that 16-character password in **Gmail / Yahoo**. Do not enter the normal mailbox
password. Spaces in the displayed app password are ignored. This is a one-time
link: the encrypted connection remains saved after closing or restarting the
laptop. The code is needed again only if the mailbox app password is revoked,
changed, or the account is deliberately removed/reconnected.

The Fetch button scans the selected date range. Free deterministic extraction runs
first and fills customer, application, case type and branch from the subject/body,
then fills missing property address/details from readable PDF, DOCX or XLSX
attachments. Paid ChatGPT email extraction is disabled by default, so MIS fetching
does not stop when API billing is unavailable.

## Environment variables

- `SECRET_KEY`: long random session secret.
- `ENCRYPTION_KEY`: secret used to encrypt email app passwords.
- `ADMIN_EMAIL`: first admin email.
- `ADMIN_PASSWORD`: secure first-login password.
- `OPENAI_API_KEY`: OpenAI API key.
- `OPENAI_MODEL`: defaults to `gpt-5.6-terra`.
- `OPENAI_DOCUMENT_EXTRACTION`: `false` keeps free local processing active;
  switch it on from Settings only when API billing is available.
- `APP_TIMEZONE`: defaults to `Asia/Kolkata`.
- `DATABASE_URL`: PostgreSQL connection string in production; SQLite locally.
- `ENABLE_EMAIL_SCHEDULER`: `false` by default.
- `EMAIL_FETCH_MINUTES`: scheduler interval, minimum 5.
- `SESSION_COOKIE_SECURE`: `true` on HTTPS deployment.
- `MAX_UPLOAD_MB`: request upload limit, default 50.
- `MAX_ZIP_FILES`: ZIP file-count limit, default 250.
- `MAX_ZIP_UNCOMPRESSED_MB`: extracted ZIP size limit, default 100.

Generate a Fernet key:

```powershell
py -3.12 -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
```

## Render deployment

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\PUSH_TO_GITHUB_AND_DEPLOY.ps1
```

Configure the required values from `.env.example` in Render. Use PostgreSQL for
production because Render web-service disk is temporary. Documents, images,
templates and reports are stored in the configured database.

## Bank report mapping

The four supplied formats are seeded automatically on first use. For another
bank, upload the original format globally or directly inside a case. SVAI uses:

1. optional `{{FIELD_NAME}}` tokens;
2. deterministic label-to-value mapping;
3. protected photo slots for the mapped DCB, SBFC and Laxmi layouts.

The uploaded original is never modified. Every generation creates a new report.
XLSM macros are preserved when an XLSM template is used. If a field cannot be
matched confidently, it is left blank for manual review instead of altering the
bank layout or guessing a value.

## Review requirement

Handwriting, unclear scans, obstructed photos and conflicting documents can
produce uncertain results. The generated report is a draft: a qualified valuer
must verify customer/application details, title facts, boundaries, areas, rates,
calculations and photographs before final issue to the bank.
