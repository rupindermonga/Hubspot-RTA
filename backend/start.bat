@echo off
REM Local-dev start. Once Doppler is wired up (runbook 01), prefix with:
REM   "%USERPROFILE%\bin\doppler.exe" run --
REM Pattern A is safe because this file is local-dev-only (no deploy hook calls it).

call venv\Scripts\activate.bat
python run.py
