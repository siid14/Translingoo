@echo off
REM PDF Translator run script for Windows

echo Creating virtual environment for PDF Translator...
if not exist venv_pdf (
    python -m venv venv_pdf
)

echo Activating virtual environment...
call venv_pdf\Scripts\activate.bat

echo Installing requirements...
pip install -r requirements_pdf.txt

echo Creating necessary directories...
if not exist ..\pdf_uploads mkdir ..\pdf_uploads
if not exist ..\pdf_downloads mkdir ..\pdf_downloads
if not exist templates mkdir templates
if not exist static\css mkdir static\css

echo Starting PDF Translator app...
python pdf_translator_app.py

pause 