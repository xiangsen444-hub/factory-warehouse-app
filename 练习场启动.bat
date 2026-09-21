@echo off
setlocal
cd /d "%~dp0"
set "WAREHOUSE_PYTHON=python"
if exist ".venv\Scripts\python.exe" (
    set "WAREHOUSE_PYTHON=%~dp0.venv\Scripts\python.exe"
    goto checkdeps
)
where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found. Install Python 3.11+ and add it to PATH.
    pause
    exit /b 1
)
:checkdeps
"%WAREHOUSE_PYTHON%" -c "import streamlit, pandas, qrcode, PIL, openpyxl, xlrd, pymupdf" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Install dependencies first: python -m pip install -r requirements.txt
    pause
    exit /b 1
)
echo Starting warehouse system at http://localhost:8502
echo Keep this window open while using the system.
"%WAREHOUSE_PYTHON%" -m streamlit run app.py --server.port 8502
pause
