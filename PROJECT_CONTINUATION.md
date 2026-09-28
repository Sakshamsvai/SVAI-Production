# SVAI Project Continuation Checkpoint

## Latest verification: 21 September 2026

- Added a permanent, duplicate-free Debit Audit at `/banking/debit-audit` and
  an `Open Debit Audit` Banking button. Existing uploaded statements backfill
  225 debit rows. Select a person/party or a month to see date-wise payments,
  UTR, narration, source statement and month totals. Different salary/remark
  wording is grouped under the same recipient name, e.g. Praveen Singh Chandel
  has seven rows across October 2025 to March 2026 totaling ₹568,982.00.
- Browser-rendered Debit Audit verified after local restart; `/health` reports
  `status=ok`, `database=connected`. Full isolated regression suite passed:
  107 tests in 112.000 seconds.

- Payment Tracking now shows month and invoice net after TDS (pending estimate: 10%
  of taxable value; confirmed bills: saved TDS). Actual Received excludes TDS.
  Each bill status lists possible receipt date, UTR, amount and a link to its
  review row; confirmed receipts retain their saved evidence. No automatic
  Received decision. Debit rows and already-used receipt references are excluded.
- Restarted local app and browser-verified authenticated HDFC page: August net
  13,284 links to 17-09-2026 / HDFCH01268336955. Local health is connected.
  Bank recognition remains pattern-based; unknown payers require review.
  Full isolated unittest suite passed: 106 tests in 83.508 seconds.

- Added the user-provided Saksham Valuer logo as
  `static/branding/saksham-valuer-logo.png`. The authenticated top header and
  login/forgot/reset pages now use it. Browser-rendered Dashboard verified the
  logo, `SAKSHAM VALUER` label and compact mobile-safe header sizing. Focused
  page-render and health tests passed.
- Repaired permanent Banking payment-history duplicates caused by an old
  identity format without the `Credit:` prefix. Created an online SQLite backup:
  `artifact_work/svai-before-payment-history-dedupe-20260921-141229.db` before
  changing live data. Startup now canonicalizes legacy keys, preserves one
  earliest source row, and prevents an old/new key pair from being saved twice.
- Browser-verified Bajaj Housing Finance Bill Audit after restart: 9 unique
  payments totaling ₹518,562.00 (previously the same 9 rows appeared twice as
  18 payments). Full isolated regression suite passed: 105 tests in 72.309
  seconds; local health returned `status=ok`, `database=connected`.
- Added and activated a conservative name-only follow-up rule. A status/document
  email without an application number attaches to an existing MIS case only when
  its customer name identifies exactly one active case (and uses bank name when
  present). It never creates a new MIS row. New Fresh/Subsequent/Revisit/Part/
  Tranche assignment wording remains outside this rule, and ambiguous same-name
  matches are not auto-merged.
- Full isolated regression suite after this change: 104 tests passed in 46.509
  seconds. The local server was restarted and `/health` again returned
  `status=ok`, `database=connected`.
- Current checkout HEAD: `f0d03a5`. Preserved existing local changes.
- Full isolated test suite: `.venv\Scripts\python.exe -m unittest discover -s tests -v`
  completed with 103 tests passing in 85.035 seconds. Deprecation and resource
  warnings remain; no test failures occurred.
- Local `/health` returned `status=ok` and `database=connected` on retry after
  an initial connection refusal. No server restart was performed in this session.
- Existing authenticated in-app browser session opened Dashboard and Billing.
  Billing displayed bank-wise ZIP/PDF/XLSX upload, bank/branch/date/review filters,
  and the bank-wise Excel export link. September had no uploaded manual bills.
- Visually checked Billing at the current narrow browser width: top PRO VALUER
  branding and vertical left navigation remain visible.
- Real-file ZIP import, populated register/edit/export browser verification, and
  public deployment verification remain pending. No live records were changed
  by this verification work.

## Latest local work: 14 September 2026

- Continued bank-wise Billing MIS request using the local `Bill Audit 2025 (1).xlsx`
  reference: bank sheets with month, bill number/date and amounts.
- Added bounded in-memory ZIP input for PDF/XLSX bills, conservative label parsing,
  duplicate skipping that preserves edits, branch metadata in an additive table,
  bank/branch/date/bill/review filters, compact rows and an Edit dialog.
- Export retains combined MIS and adds individual bank sheets. Unsupported invoice
  identifiers are blank in exports; missing fields and mismatched totals need review.
- Three targeted smoke tests passed, including synthetic ZIP import/export and
  duplicate preservation. Local synthetic page rendered and visually inspected.
- Files: `bill_import.py`, `server.py`, `templates/billing_home.html`, `static/app.css`,
  `tests/test_smoke.py`. Synthetic preview is under `artifact_work/bank-mis-preview`.
- No live data import, source-data repair, server restart or deployment was performed.
  The additive table is created by the existing startup `db.create_all()` path.
  Live authenticated ZIP/import validation remains pending. Scanned PDFs without
  extractable text go to review; OCR is not implemented by this change.

## Original July checkpoint (historical)

Saved on 27 July 2026 so work can continue after a laptop restart or in a later
Codex task. Do not place API keys, mailbox app passwords or other secrets in this
file.

## Main locations

- Working project:
  `C:\Users\Omprakash meena\Downloads\savi-main\SVAI-Production-Final`
- Clean delivery ZIP:
  `C:\Users\Omprakash meena\Downloads\savi-main\SVAI-Production-Complete-Free-Ready.zip`
- Local live database:
  `SVAI-Production-Final\instance\svai.db`
- Latest protected MIS backup:
  `SVAI-Production-Final\artifact_work\svai-before-real-mis-repair-20260727-001149.db`

The delivery ZIP intentionally excludes `.env`, the live database, mailbox
passwords, uploaded private documents, generated reports, logs, the virtual
environment and temporary repair files.

## Restart after the laptop is switched on

Open PowerShell in the working project folder and run:

```powershell
powershell -ExecutionPolicy Bypass -File .\START_LOCAL.ps1 -EasyStart
```

Then open `http://127.0.0.1:8000`.

## Completed behaviour

- OpenAI/ChatGPT is the only configured AI provider; Gemini is not used.
- Gmail and Yahoo can be linked with encrypted 16-character app passwords.
- A linked Gmail/Yahoo account stays saved after laptop/app restart. Routine
  Fetch never asks for the 16-character app password again.
- Login includes Forgot Password. A six-digit reset code is sent through a
  previously linked mailbox and expires after 10 minutes.
- MIS uses a selected From/To date range and only valuation/technical cases.
- Real subject/body/attachment parsing covers Fresh, Subsequent, Revisit,
  Part/Tranche, NPA, LAP, Purchase and Construction patterns from several banks.
- Customer, application, contact, branch, case type and property address are
  filled only when supported by the email or readable attachment.
- Existing MIS values that were manually reviewed are stable: repeat fetches
  fill missing data but do not replace correct customer/application/address,
  branch or completed status fields.
- Uncertain values remain blank for manual review; values are not invented.
- Report reminders, "Please share report", status mails and follow-ups are
  attached to the existing case and do not create another MIS row.
- Genuine Subsequent, Revisit and Part/Tranche assignments remain separate.
- The existing MIS was repaired to 205 active email cases. Fifty duplicate/reply
  rows were merged, five follow-up-only rows were hidden, and obvious junk was
  archived. Seven customer names and seven application numbers remain blank
  because the stored sources did not support a confident value.
- A report starts with the application number, then shows four explicit uploads:
  Property Documents, multi-page Visit Form (including MP Kisan screenshots),
  Site Photos and the exact bank template.
- JPEG/PNG/PDF pages uploaded through Visit Form always remain `visit_data`;
  they are never reclassified as property photos.
- Source authority is enforced: legal address/khasra/areas/boundaries come only
  from Property Documents, while actual address/khasra/areas/boundaries come
  only from Visit Form. A two-column review screen allows correction before the
  exact bank report is generated.
- Portal cases remain `Portal Pending`.
- Report generation requires the uploaded XLSX/XLSM/DOCX bank template. It fills
  known labels/tokens and photo places in the same file without splitting,
  rebuilding or generically redesigning the template.
- If safe in-place filling cannot be confirmed, no broken report is generated.
- Excel sheet names, merged cells, widths, heights, freeze panes and print area
  are guarded. XLSM macros are preserved.
- The global Dashboard / MIS button is available from logged-in pages.
- Free Local Mode reads typed/searchable PDF, DOCX and XLSX content, calculates
  valuation values and generates the exact report without OpenAI billing.
- Paid ChatGPT document/photo reading is optional and is off by default.
- A real Laxmi India test report was generated for application
  `LAPVDS100026755` / `MAHENDRA KIRAR`. It preserved the original sheet and
  merged-cell structure, placed the real internal photo and map in their
  labeled boxes, and left unsupported site-address and market-rate facts blank
  or zero for valuer review.

## Verification

- Full automated suite: 17 tests passed.
- Health endpoint returns database connected and explicitly reports either
  `Free Local` or `Paid ChatGPT + Local Fallback`.
- Real workbook QA found zero formula-error strings; the original sheet name and
  merged-cell layout were preserved.
- The clean delivery ZIP excludes `.env`, live databases, linked-mail
  credentials, uploaded private files, real generated reports, logs, virtual
  environments and temporary QA/repair files.

## Known operational notes

- Deterministic MIS email parsing works without paid OpenAI calls.
- Free Local Mode works without API billing. It cannot reliably interpret
  unclear scans or handwriting, so those facts must be checked/entered manually.
- Paid ChatGPT scan/photo reading requires OpenAI API billing/credits; ChatGPT
  Free/Plus is separate from API billing.
- A previous full mailbox refetch ended early because of a Yahoo IMAP abort and
  Gmail DNS/network failure. Use Fetch again when the connection is stable.
- Any API key pasted into chat should be revoked and replaced in SVAI Settings.

## How to continue improving

When a new bank email pattern or report format fails, provide the real subject,
the relevant body/table screenshot and the bank template. Preserve these rules:

1. do not add reminders/follow-ups as new MIS rows;
2. do keep genuine new Subsequent/Revisit/Part assignments;
3. never guess unsupported facts;
4. never restructure the uploaded bank report format.
