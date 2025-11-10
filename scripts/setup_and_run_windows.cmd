@echo off
setlocal enableextensions

rem Move to repo root (scripts\..)
pushd "%~dp0.." >nul

echo [1/5] Creating virtualenv (.venv) if missing...
if not exist .venv\Scripts\python.exe (
  py -3 -m venv .venv
)

echo [2/5] Activating virtualenv...
call .venv\Scripts\activate.bat

echo [3/5] Installing Python packages...
python -m pip install -U pip setuptools wheel
pip install playwright playwright-stealth beautifulsoup4 openpyxl pandas python-dotenv requests

echo [4/5] Installing Chromium for Playwright...
playwright install chromium

echo [5/5] Running crawler (standalone_base2_win10.py) with visible browser...
set PLAYWRIGHT_HEADLESS=0
python standalone_base2_win10.py

popd >nul
endlocal

