#!/bin/bash

# Colors for better output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${YELLOW}Translingoo Deployment Script${NC}"
echo "==============================="

# Check for Docker
if ! command -v docker &> /dev/null; then
    echo -e "${RED}Docker is not installed. Please install Docker first.${NC}"
    echo "Visit https://docs.docker.com/get-docker/ for installation instructions."
    exit 1
fi

# Make sure we're in the right directory
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
cd "$SCRIPT_DIR"

echo -e "${YELLOW}Step 1: Preparing to build the Docker image...${NC}"

# Create a temporary directory for the build
BUILD_DIR="$(mktemp -d)"
echo "Created temporary build directory: $BUILD_DIR"

# Copy flask_app contents
cp -r ./* "$BUILD_DIR/"

# Copy src directory from parent
mkdir -p "$BUILD_DIR/src"
cp -r ../src/* "$BUILD_DIR/src/"

echo -e "${YELLOW}Step 2: Building the Docker image...${NC}"
cd "$BUILD_DIR"
docker build -t translingoo .

if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to build Docker image.${NC}"
    rm -rf "$BUILD_DIR"
    exit 1
fi

echo -e "${YELLOW}Step 3: Cleaning up build files...${NC}"
cd "$SCRIPT_DIR"
rm -rf "$BUILD_DIR"

echo -e "${YELLOW}Step 4: Running the Docker container...${NC}"
echo "The application will be available at http://localhost:5001"
echo -e "${GREEN}Press Ctrl+C to stop the container.${NC}"

# Run the container
docker run --rm -p 5001:5000 translingoo

echo -e "${GREEN}Deployment complete!${NC}" 