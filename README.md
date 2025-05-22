# Excel Translator - Simple Version

A simple, standalone tool for translating Excel files with technical terms from English to French. This tool is specifically designed to handle Excel files containing Description and Message columns with technical terms.

## Features

- Works with both .xls and .xlsx files
- Detects header rows automatically
- Handles problematic Excel file formats
- Translates Description and Message columns
- Docker-based for consistent operation across all systems
- No installation of Python or dependencies needed

## Prerequisites

The only requirement is Docker. You can download it from [https://docs.docker.com/get-docker/](https://docs.docker.com/get-docker/)

## How to Use

### On macOS/Linux:

1. Download this folder
2. Open Terminal and navigate to this folder
3. Make the script executable: `chmod +x run.sh`
4. Run the translator: `./run.sh /path/to/your/excelfile.xls`
5. The translated file will be saved next to the original with "\_translated" added to the name

### On Windows:

1. Download this folder
2. Open Command Prompt and navigate to this folder
3. Run the translator: `run.bat C:\path\to\your\excelfile.xls`
4. The translated file will be saved next to the original with "\_translated" added to the name

## Advanced Usage

You can specify the output file name:

```bash
# On macOS/Linux
./run.sh /path/to/input.xls /path/to/output.xlsx

# On Windows
run.bat C:\path\to\input.xls C:\path\to\output.xlsx
```

## How It Works

1. The tool builds a Docker image containing the conversion script
2. Your Excel file is mounted into the Docker container
3. The script:
   - Determines the best way to read your specific Excel file
   - Identifies the key columns for translation
   - Applies translations to each column
   - Saves the result as a new Excel file

## Troubleshooting

If you encounter issues:

1. **Make sure Docker is running** - The whale icon should be visible in your taskbar/menu bar
2. **Check file permissions** - Ensure you have permission to read the input file and write to the output location
3. **Verify file format** - The file should be an Excel file (.xls or .xlsx) with Description or Message columns

## Support

For any issues, please contact technical support.

# Translingoo

Excel to Excel translation tool. Specifically designed for translating technical documentation containing "Description" and "Message" columns from English to French.

## Features

- Simple desktop GUI and web interface options
- Translates Excel files with "Description" and "Message" columns
- Outputs a translated Excel file with additional French translation columns
- Built-in dictionary of technical terms

## Handling Different Document Types

### Current Functionality

The application currently supports:

- Excel files (.xls and .xlsx) with clearly defined "Description" and "Message" columns
- Translation of specific columns from English to French
- Outputs a new Excel file with the original columns plus new translation columns

### Options for PDF Technical Documents

#### Option 1: Modify Current Project

This approach involves extending the current application:

1. Add PDF parsing capability using libraries like PyPDF2 or pdfplumber
2. Implement layout detection to identify tables and text content
3. Extract content for translation while preserving structure
4. Add technical vocabulary for engineering/electrical terms
5. Modify the UI to support PDF input and output options

**Pros:**

- Single application for all translation needs
- Leverage existing translation dictionary

**Cons:**

- Complex implementation to handle different document structures
- Challenging to maintain PDF layout in output
- May require significant UI changes

#### Option 2: Create a Separate Project

This approach involves creating a dedicated solution for PDF documents:

1. Build a specialized PDF processing pipeline
2. Focus on technical document translation with appropriate domain vocabulary
3. Implement structure-preserving output generation
4. Design a UI specifically for the PDF workflow

**Pros:**

- Purpose-built solution for complex documents
- Better handling of document structure
- Simpler user experience for each document type

**Cons:**

- Requires maintaining two separate applications
- Duplicate code for translation functionality

## Recommendation

For complex technical documents like the 63kV Buscoupler Protection specification, **Option 2 (separate project)** is likely the better approach. The structure and content of technical PDFs are significantly different from simple Excel files with defined columns, requiring specialized processing and a broader technical vocabulary.
