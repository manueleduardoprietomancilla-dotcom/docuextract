@echo off
echo.
echo ===================================
echo   DocuExtract AI - Setup
echo ===================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.10+ from https://python.org
    pause
    exit /b
)

echo [1/3] Installing dependencies...
pip install -r requirements.txt -q

echo [2/3] Starting server...
echo.
echo ===================================
echo   App ready at: http://localhost:8000
echo   Press Ctrl+C to stop
echo ===================================
echo.
python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
