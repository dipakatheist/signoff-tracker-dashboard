@echo off
cd /d "%~dp0"
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
  echo Installing the Excel reader dependency...
  python -m pip install -r requirements.txt
  if errorlevel 1 exit /b 1
)
python dashboard_server.py
pause