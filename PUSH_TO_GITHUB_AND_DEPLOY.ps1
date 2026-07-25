$ErrorActionPreference = "Stop"
$Repo = "https://github.com/Sakshamsvai/SVAI-Production.git"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "SVAI Final - Setup, Test and GitHub Push" -ForegroundColor Cyan
Set-Location $Root

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  Write-Host "Git not found. Install Git for Windows first." -ForegroundColor Red
  Pause
  exit 1
}
if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
  Write-Host "Python launcher not found." -ForegroundColor Red
  Pause
  exit 1
}

if (-not (Test-Path ".env")) {
  Copy-Item ".env.example" ".env"
  Write-Host ".env created. Local default login will be used." -ForegroundColor Yellow
}

if (Test-Path "venv") { Remove-Item -Recurse -Force "venv" }
py -3.12 -m venv venv
.\venv\Scripts\python.exe -m pip install --upgrade pip
.\venv\Scripts\python.exe -m pip install -r requirements.txt

Write-Host "Running syntax test..." -ForegroundColor Yellow
.\venv\Scripts\python.exe -m py_compile server.py

if (-not (Test-Path ".git")) {
  git init
}
git branch -M main
$origin = git remote get-url origin 2>$null
if ($LASTEXITCODE -ne 0) {
  git remote add origin $Repo
} elseif ($origin -ne $Repo) {
  git remote set-url origin $Repo
}

git add .
git commit -m "Final SVAI production workflow" 2>$null
git push -u origin main --force

Write-Host ""
Write-Host "GitHub push completed." -ForegroundColor Green
Write-Host "Render will auto-deploy from main branch." -ForegroundColor Green
Write-Host "Open: https://svai-valuation-app.onrender.com" -ForegroundColor Cyan
Pause
