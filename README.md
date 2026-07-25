# SVAI Production

Ready-to-use workflow:

Login → Bank Email → Genuine MIS → Case Search → ZIP/Documents/Photos →
Document Extraction → Photo Categorisation → Valuation → Bank Excel →
Report Download → Archive

## Local test

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\START_LOCAL.ps1
```

Open `http://127.0.0.1:8000`

Default local login:

- Email: `sakshamvaluer@yahoo.com`
- Password: `ChangeMe123!`

Change the password immediately from the dashboard.

## Push to GitHub and Render

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\PUSH_TO_GITHUB_AND_DEPLOY.ps1
```

Repository:

`https://github.com/Sakshamsvai/SVAI-Production.git`

Render service:

`https://svai-valuation-app.onrender.com`

## Required Render environment variables

In Render → svai-valuation-app → Environment, set:

- `SECRET_KEY`: long random value
- `ENCRYPTION_KEY`: generate with the command below
- `ADMIN_EMAIL`: `sakshamvaluer@yahoo.com`
- `ADMIN_PASSWORD`: your secure initial password
- `GEMINI_API_KEY`: your Google Gemini API key
- `GEMINI_MODEL`: `gemini-2.5-flash`
- `DATABASE_URL`: PostgreSQL connection string
- `ENABLE_EMAIL_SCHEDULER`: `false` initially

Generate encryption key:

```powershell
py -3.12 -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
```

## Persistent production data

SQLite works locally, but Render's web-service filesystem is temporary. For permanent MIS,
users, email settings, reports and uploaded files, add a PostgreSQL database and put its
internal connection string in `DATABASE_URL`.

All uploaded documents, photos, templates and generated reports are stored inside the
database, so they survive web-service redeploys when PostgreSQL is used.

## Bank Excel templates

Upload the bank's original `.xlsx` file from **Bank Templates**. Put these tokens in the
cells that SVAI must fill:

- `{{APPLICATION_NUMBER}}`
- `{{CUSTOMER_NAME}}`
- `{{CONTACT_NUMBER}}`
- `{{PROPERTY_ADDRESS}}`
- `{{BANK_NAME}}`
- `{{BRANCH_NAME}}`
- `{{CASE_TYPE}}`
- `{{REPORT_DATE}}`
- `{{LAND_AREA}}`
- `{{LAND_RATE}}`
- `{{LAND_VALUE}}`
- `{{BUILTUP_AREA}}`
- `{{CONSTRUCTION_RATE}}`
- `{{MARKET_VALUE}}`
- `{{CONSERVATIVE_VALUE}}`
- `{{DISTRESS_VALUE}}`
- `{{GOVT_VALUE}}`
- `{{REMARKS}}`

## Important limitations

- Email extraction uses regex and Gemini only when `GEMINI_API_KEY` is configured.
- It never inserts fixed fake applicant data.
- OCR quality depends on document clarity and Gemini availability.
- Exact bank report automation needs a tokenised bank template or a custom mapping for
  that bank's cells.
- Render Free services sleep when inactive; automatic IMAP polling runs only while the
  service is awake. The **Fetch Bank Emails** button always works when the app is open.
