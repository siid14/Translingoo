@echo off
echo Excel Translator - Simple Version
echo ===============================

REM Check for Docker
where docker >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Docker is not installed. Please install Docker first.
    echo Visit https://docs.docker.com/get-docker/ for installation instructions.
    exit /b 1
)

REM Check for input file
if "%~1"=="" (
    echo Please provide the path to the Excel file:
    echo Usage: run.bat path\to\excelfile.xls [output_file.xlsx]
    exit /b 1
)

set INPUT_FILE=%~f1
if not exist "%INPUT_FILE%" (
    echo File not found: %INPUT_FILE%
    exit /b 1
)

REM Determine output file
if "%~2"=="" (
    REM Create default output filename
    set OUTPUT_FILE=%~d1%~p1%~n1_translated.xlsx
) else (
    set OUTPUT_FILE=%~f2
)

echo Input file: %INPUT_FILE%
echo Output will be saved to: %OUTPUT_FILE%

REM Build the Docker image
echo.
echo Building the Docker image...
docker build -t excel-translator .

if %ERRORLEVEL% neq 0 (
    echo Failed to build Docker image.
    exit /b 1
)

REM Run the container
echo.
echo Processing the Excel file...
docker run --rm -v "%~dp1:/data" excel-translator "/data/%~nx1" -o "/data/%~nx2"

if %ERRORLEVEL% neq 0 (
    echo Failed to process the Excel file.
    exit /b 1
) else (
    echo.
    echo Success! File has been translated and saved to:
    echo %OUTPUT_FILE%
) 