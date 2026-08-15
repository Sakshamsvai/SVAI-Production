$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root
$python = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
  throw "SVAI Python environment missing. Run START_LOCAL.ps1 once first."
}
Write-Host "SVAI Local Report Worker running on this PC..." -ForegroundColor Green
Write-Host "Keep this window open while creating reports." -ForegroundColor Yellow
& $python local_report_worker.py
