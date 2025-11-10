#!/usr/bin/env pwsh
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Move to repo root (scripts/..)
Set-Location (Join-Path $PSScriptRoot '..')

Write-Host "[1/5] Creating virtualenv (.venv) if missing..."
if (-not (Test-Path .\.venv\Scripts\python.exe)) {
  py -3 -m venv .venv
}

Write-Host "[2/5] Activating virtualenv..."
. .\.venv\Scripts\Activate.ps1

Write-Host "[3/5] Installing Python packages..."
$env:PIP_DISABLE_PIP_VERSION_CHECK = '1'
python -m pip install -U pip setuptools wheel
pip install playwright playwright-stealth beautifulsoup4 openpyxl pandas python-dotenv requests

Write-Host "[4/5] Installing Chromium for Playwright..."
playwright install chromium

Write-Host "[5/5] Running crawler (standalone_base2_win10.py) with visible browser..."
$env:PLAYWRIGHT_HEADLESS = '0'
python .\standalone_base2_win10.py

