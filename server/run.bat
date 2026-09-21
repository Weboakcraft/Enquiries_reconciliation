@echo off
REM ---- OakCraft Lead Reconciler : Windows one-click start ----
cd /d "%~dp0"
where py >nul 2>nul && (set PY=py) || (set PY=python)
%PY% -c "import flask, openpyxl" 2>nul || (
  echo Installing required packages, one moment...
  %PY% -m pip install -r requirements.txt
)
%PY% app.py
pause
