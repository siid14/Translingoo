#!/bin/bash

echo "Building Excel Translator GUI for macOS..."

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Python is not installed. Please install Python 3.8 or higher."
    exit 1
fi

# Install required packages
echo "Installing required packages..."
pip3 install -r gui_requirements.txt

# Build the executable
echo "Building executable..."
pyinstaller --name "Excel Translator" --windowed --onefile \
    --add-data "Dockerfile:." \
    --add-data "converter.py:." \
    --add-data "README.md:." \
    gui_wrapper.py

echo "Build completed!"
echo "The executable is located in the 'dist' folder." 