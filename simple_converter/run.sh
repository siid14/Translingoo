#!/bin/bash

# Colors for better output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${YELLOW}Excel Translator - Simple Version${NC}"
echo "==============================="

# Check for Docker
if ! command -v docker &> /dev/null; then
    echo -e "${RED}Docker is not installed. Please install Docker first.${NC}"
    echo "Visit https://docs.docker.com/get-docker/ for installation instructions."
    exit 1
fi

# Check for input file
if [ -z "$1" ]; then
    echo -e "${RED}Please provide the path to the Excel file:${NC}"
    echo "Usage: ./run.sh /path/to/excelfile.xls [output_file.xlsx]"
    exit 1
fi

INPUT_FILE="$1"
if [ ! -f "$INPUT_FILE" ]; then
    echo -e "${RED}File not found: $INPUT_FILE${NC}"
    exit 1
fi

# Get absolute path of input file
INPUT_PATH=$(cd "$(dirname "$INPUT_FILE")" && pwd)/$(basename "$INPUT_FILE")

# Determine output file
if [ -z "$2" ]; then
    # Create default output filename
    BASENAME=$(basename "$INPUT_FILE")
    NAME_WITHOUT_EXT="${BASENAME%.*}"
    OUTPUT_FILE="$(dirname "$INPUT_PATH")/${NAME_WITHOUT_EXT}_translated.xlsx"
else
    OUTPUT_FILE="$2"
    # If not an absolute path, make it absolute
    if [[ "$OUTPUT_FILE" != /* ]]; then
        OUTPUT_FILE="$(pwd)/$OUTPUT_FILE"
    fi
fi

echo -e "${YELLOW}Input file:${NC} $INPUT_PATH"
echo -e "${YELLOW}Output will be saved to:${NC} $OUTPUT_FILE"

# Build the Docker image
echo -e "\n${YELLOW}Building the Docker image...${NC}"
docker build -t excel-translator .

if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to build Docker image.${NC}"
    exit 1
fi

# Run the container
echo -e "\n${YELLOW}Processing the Excel file...${NC}"
docker run --rm -v "$(dirname "$INPUT_PATH"):/data" excel-translator "/data/$(basename "$INPUT_PATH")" -o "/data/$(basename "$OUTPUT_FILE")"

if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to process the Excel file.${NC}"
    exit 1
else
    echo -e "\n${GREEN}Success!${NC} File has been translated and saved to:"
    echo "$OUTPUT_FILE"
fi 