# Excel Technical Translator

## Core Components

### 1. GUI Interface (`main.py`)

- Simple tkinter-based interface
- File selection for input/output
- Progress feedback
- Error handling

### 2. Excel Processing (`excel_processor.py`)

- Dictionary-based translation system
- Custom Excel file parsing
- Language detection (English/French)
- Preservation of Excel structure

### 3. Translation System

- Static dictionary of technical terms
- Bidirectional translation (English ↔ French)
- Specialized for industrial/electrical terminology

## Requirements

### Technical Dependencies

- Python 3.12+
- Dependencies:
  - pandas
  - openpyxl
  - tkinter (built-in)
  - pathlib (built-in)

### Functional Requirements

- Excel file (.xlsx) support
- Specific column ("Description") processing
- Maximum 1000 rows per file
- Offline operation
- Preservation of Excel structure

## Edge Cases & Limitations

### Text Processing

- Mixed language content
- Partial matches in translations
- Case sensitivity
- Extra spaces/formatting
- Special characters
- Abbreviations

### File Handling

- Corrupted Excel files
- Protected/locked files
- Large files (>1000 rows)
- Different Excel formats
- Missing columns
- Empty cells

### Translation

- Unknown technical terms
- Context-dependent translations
- Compound terms
- Ambiguous terms
- New terminology

## Installation

```bash
# Create a virtual environment
python3 -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

```bash
# Run the application from the project root directory (Translingoo)
python3 src/main.py
```

**Note:** Make sure to run this command from the project root directory, not from inside the src directory.

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request

# PDF Technical Document Translator

A specialized tool for translating technical PDF documents from English to French with multiple output formats.

## Features

- Web-based interface for easy document uploading
- Multiple output formats (Word, Excel, PDF)
- Preserves document formatting where possible
- Rich technical vocabulary for engineering/electrical terms
- Automatic translation of tables and text content

## How to Run

### On macOS/Linux:

within src directory

1. Make sure Python 3.8+ is installed: `python --version`
2. Make the run script executable: `chmod +x run_pdf_translator.sh`
3. Run the application: `./run_pdf_translator.sh`
4. Open your browser and go to `http://localhost:5001`

### On Windows:

1. Make sure Python 3.8+ is installed: `python --version`
2. Run the application: `run_pdf_translator.bat`
3. Open your browser and go to `http://localhost:5001`

## Output Formats

### Word Document (.docx)

- **Best for**: Preserving formatting, tables, and structure
- **Process**: PDF is converted to Word, then translations are applied while maintaining formatting
- **Advantages**: Most visually similar to the original, editable afterwards

### Excel Spreadsheet (.xlsx)

- **Best for**: Manual review and editing of translations
- **Process**: Extracts text content with page references and type classification
- **Advantages**: Easy to edit and review all translations in one place

### PDF Document (.pdf)

- **Best for**: Quick translations where formatting is less critical
- **Process**: Creates a new PDF with translated content based on extracted text
- **Advantages**: Same file format as the source, albeit with simplified layout

## Technical Dictionary

The translator includes a specialized dictionary for engineering and electrical domain terms, with a focus on protection relay systems and electrical specifications.

Custom dictionaries can be added by creating a CSV file with columns `english` and `french`.

## Troubleshooting

- **Installation issues**: Check if your Python environment is compatible (3.8+)
- **Web UI not loading**: Make sure port 5001 is available
- **PDF conversion errors**: Try the Excel output format as a fallback
- **Word document formatting issues**: For complex documents, the Excel format may provide better results

## Command Line Usage

You can also use the translator from the command line:

```bash
# For Word document output (recommended)
python pdf_translator.py /path/to/document.pdf -w

# For Excel output
python pdf_translator.py /path/to/document.pdf -e

# For PDF output
python pdf_translator.py /path/to/document.pdf
```

## Dependencies

All dependencies are automatically installed by the run scripts. If needed, you can install them manually:

```bash
pip install -r requirements_pdf.txt
```
