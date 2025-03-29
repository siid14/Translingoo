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
# Run the application
python3 src/main.py
```

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request
