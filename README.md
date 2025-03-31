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
