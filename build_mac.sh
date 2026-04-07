#!/usr/bin/env bash
set -euo pipefail

python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt

python3 -m PyInstaller \
  --windowed \
  --name "Purchase KPI Dashboard" \
  --noconfirm \
  main.py

echo "Build complete: dist/Purchase KPI Dashboard.app"
