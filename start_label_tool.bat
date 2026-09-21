@echo off
setlocal
cd /d "%~dp0"
set "LABEL_PYTHON=python"
if exist ".venv\Scripts\python.exe" set "LABEL_PYTHON=%~dp0.venv\Scripts\python.exe"

echo ==========================================
echo   Label Printing Tool - standalone
echo   Not related to the warehouse system
echo ==========================================
echo.

if not exist ".venv\Scripts\python.exe" (
    where python >nul 2>nul
    if errorlevel 1 goto nopython
)

"%LABEL_PYTHON%" -c "import streamlit, PIL, qrcode, win32ui, pandas, xlrd, openpyxl" >nul 2>nul
if errorlevel 1 goto needinstall
goto run

:needinstall
echo [First run] Installing dependencies, needs internet, about 1-2 min...
"%LABEL_PYTHON%" -m pip install -r label_tool_requirements.txt
if errorlevel 1 goto installfail

:run
echo.
echo Starting... browser will open at http://localhost:8503
echo [NOTE] Keep this window open. Closing it stops the tool.
echo ==========================================
echo.
start "" /min cmd /c "timeout /t 4 /nobreak >nul & start http://localhost:8503"
"%LABEL_PYTHON%" -m streamlit run label_tool.py --server.port 8503 --server.headless true

echo.
echo Stopped. Press any key to close.
pause >nul
goto :eof

:nopython
echo [ERROR] Python not found on this computer.
echo Install Python from https://www.python.org/downloads/
echo (check "Add python.exe to PATH" during install), then run this again.
echo.
pause
exit /b 1

:installfail
echo [ERROR] Install failed, check your network and try again.
pause
exit /b 1
