@echo off
REM Local-dev start — Pattern A (in-place Doppler wrap).
REM Safe because start.bat is local-dev-only; no deploy hook or Docker CMD calls it.
REM Phase 1 Doppler dev retrofit complete 2026-04-30.

call venv\Scripts\activate.bat
doppler run -- python run.py
