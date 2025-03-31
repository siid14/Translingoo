#!/bin/bash

# Colors for better output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${YELLOW}Translingoo Deployment Script (No Docker)${NC}"
echo "==============================="

# Check for Python
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}Python 3 is not installed. Please install Python 3 first.${NC}"
    exit 1
fi

# Make sure we're in the right directory
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
cd "$SCRIPT_DIR"

echo -e "${YELLOW}Step 1: Installing required packages...${NC}"
pip3 install -r requirements.txt

# Create symbolic link to the src directory if it doesn't exist
if [ ! -d "./src" ]; then
    echo -e "${YELLOW}Creating symbolic link to src directory...${NC}"
    ln -s ../src ./src
fi

# Ensure the upload and download directories exist
mkdir -p uploads downloads

echo -e "${YELLOW}Step 2: Starting the Flask application...${NC}"
echo "The application will be available at http://localhost:5000"
echo -e "${GREEN}Press Ctrl+C to stop the server.${NC}"

# Run Flask with Gunicorn if available, otherwise use development server
if command -v gunicorn &> /dev/null; then
    echo "Using Gunicorn server (recommended for production)"
    gunicorn -w 4 -b 0.0.0.0:5000 app:app
else
    echo "Using Flask development server (not recommended for production)"
    export FLASK_APP=app.py
    export FLASK_DEBUG=1
    python3 -m flask run --host=0.0.0.0
fi 