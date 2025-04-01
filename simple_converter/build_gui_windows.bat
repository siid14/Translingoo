@echo off
echo Building Excel Translator GUI for Windows...

:: Ensure Python is installed
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Python is not installed. Please install Python 3.8 or higher.
    exit /b 1
)

:: Install required packages
echo Installing required packages...
pip install -r gui_requirements.txt

:: Build the executable
echo Building executable...
pyinstaller --name "Excel Translator" --windowed --onefile --add-data "Dockerfile;." --add-data "converter.py;." --add-data "README.md;." gui_wrapper.py

echo Build completed!
echo The executable is located in the "dist" folder.
pause 