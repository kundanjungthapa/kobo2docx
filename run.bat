@echo off
title Kobo2Docx Auto-Converter
echo ============================================================
echo               Kobo2Docx Automated Converter                
echo ============================================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not added to PATH.
    echo Please install Python 3.8 or higher from python.org
    pause
    exit /b
)

:: Create Virtual Environment if it doesn't exist
if not exist "venv\" (
    echo [INFO] Creating virtual environment...
    python -m venv venv
)

:: Activate Virtual Environment & Install Dependencies
call venv\Scripts\activate.bat
echo [INFO] Checking dependencies...
pip install -r requirements.txt --quiet --disable-pip-version-check

:: Run Main Converter Script
echo.
echo [INFO] Running Kobo2Docx Converter...
echo ------------------------------------------------------------
python main.py
echo ------------------------------------------------------------
echo.

:: Open Outputs Folder Automatically
if exist "outputs\" (
    echo Opening 'outputs' folder...
    start outputs
)

echo.
echo Conversion finished successfully!
pause