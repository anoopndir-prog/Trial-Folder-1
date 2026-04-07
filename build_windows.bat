@echo off
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

python -m PyInstaller --windowed --name "Purchase KPI Dashboard" --noconfirm main.py

echo Build complete: dist\Purchase KPI Dashboard\Purchase KPI Dashboard.exe
pause
