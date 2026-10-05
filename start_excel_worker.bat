@echo off
cd /d "%~dp0"
if not defined EXCEL_WORKER_TOKEN (
  echo Set EXCEL_WORKER_TOKEN before starting this worker. See EXCEL_WORKER_SETUP.md.
  pause
  exit /b 1
)
python excel_worker.py --host 127.0.0.1 --port 8767
pause
