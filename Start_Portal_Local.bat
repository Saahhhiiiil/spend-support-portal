@echo off
setlocal
cd /d "%~dp0"

echo ================================================
echo   Spend/Support Processing Portal - Local Mode
echo ================================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=python"
    ) else (
        echo Python was not found. Install Python 3.10 or later and enable Add Python to PATH.
        pause
        exit /b 1
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating a local virtual environment...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 goto :error
)

echo Installing or checking required packages...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :error
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo.
echo Starting the portal on this computer...
echo Keep this window open while using the portal.
echo Your browser should open at http://localhost:8501
 echo.
".venv\Scripts\python.exe" -m streamlit run app.py --server.address 127.0.0.1
if errorlevel 1 goto :error
exit /b 0

:error
echo.
echo The portal could not start. Check the error shown above.
pause
exit /b 1
