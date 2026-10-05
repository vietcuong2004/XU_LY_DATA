@echo off
cd /d "%~dp0"
python ui_server.py --open
if errorlevel 1 (
  echo.
  echo Khong khoi dong duoc. Kiem tra Python va chay: python -m pip install -r requirements.txt
  pause
)
