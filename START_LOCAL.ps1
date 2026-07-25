$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
if (-not (Test-Path "venv")) {
  py -3.12 -m venv venv
  .\venv\Scripts\python.exe -m pip install -r requirements.txt
}
$env:SECRET_KEY="local-development-secret-change"
$env:ADMIN_EMAIL="sakshamvaluer@yahoo.com"
$env:ADMIN_PASSWORD="ChangeMe123!"
$env:DATABASE_URL="sqlite:///svai.db"
Start-Process "http://127.0.0.1:8000"
.\venv\Scripts\python.exe server.py
