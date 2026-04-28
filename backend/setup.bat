@echo off
REM One-time setup: create venv + install requirements.
REM No app code runs here, so no Doppler injection needed.

if not exist venv (
  python -m venv venv
)

call venv\Scripts\activate.bat
python -m pip install --upgrade pip wheel
pip install -r requirements.txt

if not exist .env (
  copy .env.example .env >nul
  echo Created .env from .env.example. Edit before running create_admin.py.
)

echo.
echo Setup complete. Next:
echo   1. Edit .env (or wire up Doppler — runbook 01)
echo   2. python create_admin.py        (one-time, seeds the admin row)
echo   3. start.bat                     (or: doppler run -- python run.py)
