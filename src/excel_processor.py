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
            
            # Try reading with default settings first
            print("DEBUG: Reading Excel with following parameters:")
            print("- sheet_name=0")
            print("- skiprows=13")
            print("- engine=openpyxl")
            
            # First try to get sheet names
            xls = pd.ExcelFile(file_path, engine='openpyxl')
            print(f"DEBUG: Available sheets: {xls.sheet_names}")
            
            self.input_df = pd.read_excel(
                file_path,
                sheet_name=0,  # Explicitly specify first sheet
                engine='openpyxl',
                skiprows=13,  # Skip the header rows until the actual data starts
                usecols=[
                    '#',
                    'Date & Time',
                    'Origin',
                    'Description',
                    'Type',
                    'Message',
                    'Category',
                    'Device',
                    'Source'
                ]
            )
            
            print(f"DEBUG: DataFrame loaded successfully")
            print(f"DEBUG: Columns found: {self.input_df.columns.tolist()}")
            print(f"DEBUG: Number of rows: {len(self.input_df)}")
            
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