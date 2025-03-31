@echo off
echo Translingoo Deployment Script
echo ===============================

REM Check for Docker
where docker >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Docker is not installed. Please install Docker first.
    echo Visit https://docs.docker.com/get-docker/ for installation instructions.
    exit /b 1
)

echo Step 1: Preparing to build the Docker image...

REM Create a temporary directory for the build
set BUILD_DIR=%TEMP%\translingoo_build
if exist %BUILD_DIR% rmdir /s /q %BUILD_DIR%
mkdir %BUILD_DIR%
echo Created temporary build directory: %BUILD_DIR%

REM Copy flask_app contents
xcopy /E /I /Y .\* %BUILD_DIR%\

REM Copy src directory from parent
mkdir %BUILD_DIR%\src
xcopy /E /I /Y ..\src\* %BUILD_DIR%\src\

echo Step 2: Building the Docker image...
cd %BUILD_DIR%
docker build -t translingoo .

if %ERRORLEVEL% neq 0 (
    echo Failed to build Docker image.
    rmdir /s /q %BUILD_DIR%
    exit /b 1
)

echo Step 3: Cleaning up build files...
cd %~dp0
rmdir /s /q %BUILD_DIR%

echo Step 4: Running the Docker container...
echo The application will be available at http://localhost:5001
echo Press Ctrl+C to stop the container.

REM Run the container
docker run --rm -p 5001:5000 translingoo

echo Deployment complete! 