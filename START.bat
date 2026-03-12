@echo off
:: IMPORTANT: Change to the folder where this .bat file lives
cd /d "%~dp0"

title CGTMSE Portal v3
echo.
echo ================================================
echo   CGTMSE Data Collection Portal v3
echo   Folder: %~dp0
echo ================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    echo Install from https://python.org
    echo Tick "Add Python to PATH" during install.
    pause & exit /b
)

echo Installing required packages...
python -m pip install flask openpyxl werkzeug --quiet
echo.
echo ================================================
echo   Admin:  http://localhost:5000
echo   Branch: http://localhost:5000/branch
echo   Login:  admin / admin123
echo ================================================
echo.
echo Press Ctrl+C to stop.
echo.
python app.py
pause
