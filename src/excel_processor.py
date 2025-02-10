import pandas as pd
from pathlib import Path

class ExcelProcessor:
    def __init__(self):
        self.input_df = None
        self.output_df = None

    def load_excel(self, file_path):
        """Load the Excel file into a pandas DataFrame."""
        try:
            print(f"\nDEBUG: Attempting to load file: {file_path}")
            print(f"DEBUG: File exists: {Path(file_path).exists()}")
            
            import zipfile
            import xml.etree.ElementTree as ET
            
            with zipfile.ZipFile(file_path) as zf:
                # Define namespaces
                namespaces = {
                    'main': 'http://purl.oclc.org/ooxml/spreadsheetml/main',
                    'r': 'http://purl.oclc.org/ooxml/officeDocument/relationships'
                }
                
                # Read shared strings if they exist
                shared_strings = []
                if 'xl/sharedStrings.xml' in zf.namelist():
                    with zf.open('xl/sharedStrings.xml') as f:
                        strings_tree = ET.parse(f)
                        strings_root = strings_tree.getroot()
                        for si in strings_root.findall('.//{%s}t' % namespaces['main']):
                            shared_strings.append(si.text)
                    print(f"DEBUG: Loaded {len(shared_strings)} shared strings")
                
                # Read the worksheet
                print("\nDEBUG: Reading sheet1.xml directly")
                with zf.open('xl/worksheets/sheet1.xml') as f:
                    sheet_tree = ET.parse(f)
                    sheet_root = sheet_tree.getroot()
                    
                    # Get all rows
                    rows = sheet_root.findall('.//{%s}row' % namespaces['main'])
                    print(f"Found {len(rows)} rows")
                    
                    # Process rows into data
                    data = []
                    for row in rows:
                        row_data = []
                        cells = row.findall('{%s}c' % namespaces['main'])
                        
                        for cell in cells:
                            # Get cell value
                            v = cell.find('{%s}v' % namespaces['main'])
                            t = cell.get('t')  # cell type
                            
                            if v is not None:
                                value = v.text
                                # If cell type is 's', it's a shared string
                                if t == 's' and shared_strings:
                                    try:
                                        value = shared_strings[int(value)]
                                    except (ValueError, IndexError):
                                        pass
                                row_data.append(value)
                            else:
                                row_data.append(None)
                        
                        data.append(row_data)
                        
                        # Print first few rows for debugging
                        if len(data) <= 5:
                            print(f"Row {len(data)}: {row_data}")
                    
                    # Create DataFrame
                    self.input_df = pd.DataFrame(data[13:])  # Skip first 13 rows as before
                    
                    # Set column names based on the first row
                    self.input_df.columns = self.input_df.iloc[0]
                    self.input_df = self.input_df[1:]  # Remove the header row
                    
                    print("\nDEBUG: Successfully created DataFrame")
                    print("\nDEBUG: File Content Preview:")
                    print(self.input_df.head())
                    print(f"\nDEBUG: DataFrame shape: {self.input_df.shape}")
                    print(f"DEBUG: Columns: {self.input_df.columns.tolist()}")
                    
                    return True
                
        except Exception as e:
            print(f"\nDEBUG: Error details:")
            print(f"- Error type: {type(e).__name__}")
            print(f"- Error message: {str(e)}")
            print(f"- File path: {file_path}")
            return False

    def process_file(self):
        """Process the loaded Excel file."""
        if self.input_df is None:
            print("DEBUG: input_df is None")
            return False
        
        print("\nDEBUG: Starting file processing")
        
        # Create a copy of the input DataFrame
        self.output_df = self.input_df.copy()
        print(f"DEBUG: Created output DataFrame with {len(self.output_df)} rows")
        
        # Ensure we have the Description column
        if 'Description' not in self.output_df.columns:
            print("DEBUG: 'Description' column not found")
            print(f"DEBUG: Available columns: {self.output_df.columns.tolist()}")
            return False
            
        # Limit to 1000 rows if necessary
        if len(self.output_df) > 1000:
            self.output_df = self.output_df.head(1000)
            print("DEBUG: Truncated to 1000 rows")
            
        return True

    def save_excel(self, output_path):
        """Save the processed DataFrame to a new Excel file."""
        if self.output_df is None:
            print("DEBUG: output_df is None")
            return False
            
        try:
            print(f"\nDEBUG: Attempting to save file to: {output_path}")
            print(f"DEBUG: DataFrame shape: {self.output_df.shape}")
            self.output_df.to_excel(output_path, index=False, engine='openpyxl')
            print("DEBUG: File saved successfully")
            return True
        except Exception as e:
            print(f"\nDEBUG: Error while saving:")
            print(f"- Error type: {type(e).__name__}")
            print(f"- Error message: {str(e)}")
            return False 