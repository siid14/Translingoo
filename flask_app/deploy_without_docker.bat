@echo off
echo Translingoo Deployment Script (No Docker)
echo ===============================

REM Check for Python
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Python is not installed. Please install Python first.
    exit /b 1
)

echo Step 1: Installing required packages...
pip install -r requirements.txt

REM Create symbolic link to the src directory if it doesn't exist
if not exist ".\src\" (
    echo Creating junction to src directory...
    mklink /J ".\src" "..\src"
)

REM Ensure the upload and download directories exist
if not exist ".\uploads" mkdir .\uploads
if not exist ".\downloads" mkdir .\downloads

echo Step 2: Starting the Flask application...
echo The application will be available at http://localhost:5000
echo Press Ctrl+C to stop the server.

REM Run Flask
set FLASK_APP=app.py
set FLASK_DEBUG=1
python -m flask run --host=0.0.0.0 