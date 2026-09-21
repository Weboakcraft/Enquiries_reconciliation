#!/usr/bin/env bash
# ---- OakCraft Lead Reconciler : macOS / Linux start ----
cd "$(dirname "$0")"
python3 -c "import flask, openpyxl" 2>/dev/null || {
  echo "Installing required packages, one moment..."
  python3 -m pip install -r requirements.txt
}
exec python3 app.py
