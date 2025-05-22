#!/bin/bash

# PDF Translator run script

# Create virtual environment if it doesn't exist
if [ ! -d "venv_pdf" ]; then
    echo "Creating virtual environment for PDF Translator..."
    python3 -m venv venv_pdf
fi

# Activate virtual environment
source venv_pdf/bin/activate

# Install requirements
echo "Installing requirements..."
pip install -r requirements_pdf.txt

# Create necessary directories
mkdir -p ../pdf_uploads
mkdir -p ../pdf_downloads
mkdir -p templates
mkdir -p static/css

# Run the app
echo "Starting PDF Translator app..."
python pdf_translator_app.py 