# SVAI Project Continuation Checkpoint

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
