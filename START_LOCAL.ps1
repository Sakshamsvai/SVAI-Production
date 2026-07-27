[CmdletBinding()]
param(
  [switch]$EasyStart
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

function New-RandomSecret {
  $bytes = New-Object byte[] 48
  $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
  try { $rng.GetBytes($bytes) } finally { $rng.Dispose() }
  return [Convert]::ToBase64String($bytes)
}

function ConvertFrom-SecureInput([Security.SecureString]$SecureValue) {
  $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureValue)
  try {
    return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
  } finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
  }
}

function Quote-DotEnv([string]$Value) {
  $escaped = $Value.Replace("\", "\\").Replace('"', '\"')
  return '"' + $escaped + '"'
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "      SVAI ONE-CLICK SETUP AND START" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

$EnvPath = Join-Path $Root ".env"
$ApiKeyReady = $false
if (Test-Path -LiteralPath $EnvPath) {
  $apiLine = Get-Content -LiteralPath $EnvPath |
    Where-Object { $_ -match '^\s*OPENAI_API_KEY\s*=' } |
    Select-Object -First 1
  if ($apiLine) {
    $savedKey = ($apiLine -replace '^\s*OPENAI_API_KEY\s*=\s*', '').Trim().Trim('"')
    $ApiKeyReady = $savedKey.StartsWith("sk-") -and $savedKey.Length -gt 20
  }
}

if (-not $ApiKeyReady -and $EasyStart) {
  $envLines = @(
    "SECRET_KEY=$(Quote-DotEnv (New-RandomSecret))"
    "ENCRYPTION_KEY=$(Quote-DotEnv (New-RandomSecret))"
    "ADMIN_EMAIL=sakshamvaluer@yahoo.com"
    "ADMIN_PASSWORD=ChangeMe123!"
    "OPENAI_API_KEY="
    "OPENAI_MODEL=gpt-5.6-terra"
    "OPENAI_EMAIL_EXTRACTION=false"
    "OPENAI_DOCUMENT_EXTRACTION=false"
    "APP_TIMEZONE=Asia/Kolkata"
    "DATABASE_URL=sqlite:///svai.db"
    "ENABLE_EMAIL_SCHEDULER=false"
    "EMAIL_FETCH_MINUTES=10"
    "SESSION_COOKIE_SECURE=false"
    "MAX_UPLOAD_MB=50"
    "MAX_ZIP_FILES=250"
    "MAX_ZIP_UNCOMPRESSED_MB=100"
  )
  [IO.File]::WriteAllLines(
    $EnvPath,
    $envLines,
    (New-Object System.Text.UTF8Encoding($false))
  )
  Write-Host "Easy setup settings create ho gayi hain." -ForegroundColor Green
} elseif (-not $ApiKeyReady) {
  Write-Host "Step 1: OpenAI API key banani hai." -ForegroundColor Yellow
  Write-Host "Browser ab OpenAI API Keys page kholega."
  Write-Host "Login karein, 'Create new secret key' dabayein aur key Copy karein."
  Start-Process "https://platform.openai.com/api-keys"
  Write-Host ""

  do {
    $secureApiKey = Read-Host "Copy ki hui OpenAI API key yahan paste karein" -AsSecureString
    $apiKey = ConvertFrom-SecureInput $secureApiKey
    if (-not ($apiKey.StartsWith("sk-") -and $apiKey.Length -gt 20)) {
      Write-Host "Ye API key sahi nahi lag rahi. Key 'sk-' se shuru honi chahiye." -ForegroundColor Red
    }
  } until ($apiKey.StartsWith("sk-") -and $apiKey.Length -gt 20)

  do {
    $secureAdminPassword = Read-Host "SVAI login ke liye apna naya password banayein (kam se kam 8 characters)" -AsSecureString
    $adminPassword = ConvertFrom-SecureInput $secureAdminPassword
    if ($adminPassword.Length -lt 8) {
      Write-Host "Password kam se kam 8 characters ka hona chahiye." -ForegroundColor Red
    }
  } until ($adminPassword.Length -ge 8)

  $envLines = @(
    "SECRET_KEY=$(Quote-DotEnv (New-RandomSecret))"
    "ENCRYPTION_KEY=$(Quote-DotEnv (New-RandomSecret))"
    "ADMIN_EMAIL=sakshamvaluer@yahoo.com"
    "ADMIN_PASSWORD=$(Quote-DotEnv $adminPassword)"
    "OPENAI_API_KEY=$(Quote-DotEnv $apiKey)"
    "OPENAI_MODEL=gpt-5.6-terra"
    "OPENAI_EMAIL_EXTRACTION=false"
    "OPENAI_DOCUMENT_EXTRACTION=false"
    "APP_TIMEZONE=Asia/Kolkata"
    "DATABASE_URL=sqlite:///svai.db"
    "ENABLE_EMAIL_SCHEDULER=false"
    "EMAIL_FETCH_MINUTES=10"
    "SESSION_COOKIE_SECURE=false"
    "MAX_UPLOAD_MB=50"
    "MAX_ZIP_FILES=250"
    "MAX_ZIP_UNCOMPRESSED_MB=100"
  )
  [IO.File]::WriteAllLines(
    $EnvPath,
    $envLines,
    (New-Object System.Text.UTF8Encoding($false))
  )
  $apiKey = $null
  $adminPassword = $null
  Write-Host "Secure settings save ho gayi hain." -ForegroundColor Green
}

$PythonExe = $null
$UsePyLauncher = $false
if (Get-Command py -ErrorAction SilentlyContinue) {
  $CandidatePython = (Get-Command py).Source
  try {
    & $CandidatePython -3.12 -c "import sys" 2>$null
    if ($LASTEXITCODE -eq 0) {
      $PythonExe = $CandidatePython
      $UsePyLauncher = $true
    }
  } catch {}
}
if (-not $PythonExe -and (Get-Command python -ErrorAction SilentlyContinue)) {
  $CandidatePython = (Get-Command python).Source
  try {
    & $CandidatePython -c "import sys" 2>$null
    if ($LASTEXITCODE -eq 0) { $PythonExe = $CandidatePython }
  } catch {}
}
if (-not $PythonExe -and (Get-Command python3 -ErrorAction SilentlyContinue)) {
  $CandidatePython = (Get-Command python3).Source
  try {
    & $CandidatePython -c "import sys" 2>$null
    if ($LASTEXITCODE -eq 0) { $PythonExe = $CandidatePython }
  } catch {}
}
if (-not $PythonExe) {
  $BundledPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
  if (Test-Path -LiteralPath $BundledPython) {
    $PythonExe = $BundledPython
  }
}

if (-not $PythonExe) {
  Write-Host "Python nahi mila. Browser mein Python download page khul raha hai." -ForegroundColor Red
  Start-Process "https://www.python.org/downloads/windows/"
  Read-Host "Python install karne ke baad ye file dobara chalayein. Band karne ke liye Enter dabayein"
  exit 1
}

$VenvPython = Join-Path $Root ".venv\Scripts\python.exe"
$VenvReady = $false
if (Test-Path -LiteralPath $VenvPython) {
  try {
    & $VenvPython -c "import flask, openpyxl, pypdf" 2>$null
    $VenvReady = $LASTEXITCODE -eq 0
  } catch {
    $VenvReady = $false
  }
}
if (-not $VenvReady) {
  Write-Host ""
  Write-Host "Step 2: SVAI ki zaroori files install ho rahi hain..." -ForegroundColor Yellow
  $ResolvedRoot = (Resolve-Path -LiteralPath $Root).Path
  $VenvPath = Join-Path $ResolvedRoot ".venv"
  if (Test-Path -LiteralPath $VenvPath) {
    if (-not $VenvPath.StartsWith(
      $ResolvedRoot + [IO.Path]::DirectorySeparatorChar,
      [StringComparison]::OrdinalIgnoreCase
    )) {
      throw "Unsafe virtual environment path: $VenvPath"
    }
    Remove-Item -LiteralPath $VenvPath -Recurse -Force
  }
  if ($UsePyLauncher) {
    & $PythonExe -3.12 -m venv ".venv"
  } else {
    & $PythonExe -m venv ".venv"
  }
  if ($LASTEXITCODE -ne 0) { throw "Python virtual environment create nahi hua." }
  & $VenvPython -m pip install -r "requirements.txt"
  if ($LASTEXITCODE -ne 0) { throw "SVAI dependencies install nahi hui." }
}

function Test-SvaiRunning {
  try {
    $health = Invoke-RestMethod -Uri "http://127.0.0.1:8000/health" -TimeoutSec 2
    return $health.status -eq "ok"
  } catch {
    return $false
  }
}

if (-not (Test-SvaiRunning)) {
  Write-Host ""
  Write-Host "Step 3: SVAI start ho raha hai..." -ForegroundColor Yellow
  $stdoutLog = Join-Path $Root "svai-server.out.log"
  $stderrLog = Join-Path $Root "svai-server.err.log"
  Start-Process `
    -FilePath $VenvPython `
    -ArgumentList "server.py" `
    -WorkingDirectory $Root `
    -WindowStyle Hidden `
    -RedirectStandardOutput $stdoutLog `
    -RedirectStandardError $stderrLog | Out-Null

  $started = $false
  for ($attempt = 1; $attempt -le 30; $attempt++) {
    Start-Sleep -Seconds 1
    if (Test-SvaiRunning) {
      $started = $true
      break
    }
  }
  if (-not $started) {
    Write-Host "SVAI start nahi hua. Error file: $stderrLog" -ForegroundColor Red
    Read-Host "Band karne ke liye Enter dabayein"
    exit 1
  }
}

Write-Host ""
Write-Host "SVAI READY HAI" -ForegroundColor Green
Write-Host "Login Email: sakshamvaluer@yahoo.com"
Write-Host "Password: setup ke samay jo aapne banaya tha"
Write-Host "App Address: http://127.0.0.1:8000"
Write-Host ""
Start-Process "http://127.0.0.1:8000"
if (-not $EasyStart) {
  Read-Host "Browser khul gaya hai. Is PowerShell window ko band karne ke liye Enter dabayein"
}
