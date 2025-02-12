# Translingoo

A specialized Excel translation tool designed for GE industrial system exports, focusing on translating technical terms between English and French in electrical system logs.

## Overview

Translingoo is a Python-based utility that processes Excel files containing industrial system logs and translates technical descriptions between English and French. It is specifically designed to handle GE industrial system exports while preserving the accuracy of technical terminology.

## Key Features

- Supports complex Excel (.xlsx) files containing system logs
- Translates technical descriptions between English and French
- Preserves the original Excel structure and formatting
- Works entirely offline
- Ensures technical accuracy through predefined translations
- Automatically detects the language and performs bidirectional translation
- Skips text that is already in the target language

## Technical Requirements

- Python 3.12+
- Required Python packages:
  - `pandas`
  - `openpyxl`
  - `pathlib`

## Installation

1. Clone the repository:

   ```sh
   git clone https://github.com/your-repo/translingoo.git
   cd translingoo
   ```

2. Create and activate a virtual environment:

   - On macOS/Linux:
     ```sh
     python -m venv venv
     source venv/bin/activate
     ```
   - On Windows:

     ```sh
     python -m venv venv

     or

     python3 -m venv venv
     venv\Scripts\activate
     ```

3. Install the required packages:
   ```sh
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage

1. Place your Excel file in an accessible location.
2. Run the program:

   ```sh
   python src/main.py

   or

   python3 src/main.py
   ```

3. Follow the prompts to select your input file.
4. The translated file will be saved with a "\_translated" suffix.

### Using as a Module

You can also use Translingoo within a Python script:

```python
from src.excel_processor import ExcelProcessor

# Initialize processor
processor = ExcelProcessor()

# Process file
processor.load_excel("input.xlsx")
processor.process_file()
processor.save_excel("output_translated.xlsx")
```

## How It Works

Translingoo employs a dictionary-based translation system tailored for industrial and electrical terminology. This approach ensures:

- Consistent translations
- High technical accuracy
- Full offline functionality
- Preservation of domain-specific terminology

### Example Translations

| English                    | French                             |
| -------------------------- | ---------------------------------- |
| DEAD INCOMING DEAD RUNNING | ENTRÉE INACTIVE EXÉCUTION INACTIVE |
| CLOSE PERMISSIVE           | PERMISSIF DE FERMETURE             |
| PRESENCE OF VOLTAGE        | PRÉSENCE DE TENSION                |

## Project Structure

```
Translingoo/
├── src/
│   ├── __init__.py
│   ├── excel_processor.py  # Core translation logic
│   └── main.py  # Entry point
├── requirements.txt  # Project dependencies
└── README.md  # Documentation
```

## Limitations

- Supports only predefined technical terms
- Focuses on English-French translation pairs
- Designed for GE industrial system export formats
- Uses a dictionary-based approach (not AI/ML-based)
- Processes a maximum of 1000 rows per file

## Adding New Translations

To expand the translation dictionary:

1. Open `src/excel_processor.py`
2. Locate the `translations` dictionary
3. Add new translation pairs:

```python
translations = {
    'ENGLISH_TERM': 'FRENCH_TERM',
    # Add your new translations here
}
```

## Error Handling

Translingoo provides detailed debugging output, including:

- Verification of file existence and format
- Status updates on the translation process
- Detailed error messages
- Validation of Excel structure

## Troubleshooting

### Common Issues and Solutions

1. **File not found**

   - Ensure the file path is correct
   - Check file permissions

2. **Translation not working**

   - Verify the term exists in the translations dictionary
   - Check for exact spelling and capitalization

3. **Excel format issues**
   - Ensure the file is in .xlsx format
   - Check for file corruption

## License

MIT License

## Support

For issues or questions:

1. Open an issue on GitHub
2. Provide the following details:
   - Error messages
   - Sample data (if possible)
   - Expected behavior
   - Your environment details

## Version History

- **1.0.0**: Initial release
  - Basic translation functionality
  - Excel file handling
  - Dictionary-based translations

## Acknowledgments

- Developed for use with GE industrial system exports
- Designed for electrical system terminology
- Built for reliability and accuracy
