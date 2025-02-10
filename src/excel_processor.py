import pandas as pd
from pathlib import Path

class ExcelProcessor:
    def __init__(self):
        self.input_df = None
        self.output_df = None

    def load_excel(self, file_path):
        """Load the Excel file into a pandas DataFrame."""
        try:
            self.input_df = pd.read_excel(file_path)
            return True
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return False

    def process_file(self):
        """Process the loaded Excel file."""
        if self.input_df is None:
            return False
        
        # Create a copy of the input DataFrame
        self.output_df = self.input_df.copy()
        
        # Ensure we have the Description column
        if 'Description' not in self.output_df.columns:
            print("Error: 'Description' column not found in Excel file")
            return False
            
        # Limit to 1000 rows if necessary
        if len(self.output_df) > 1000:
            self.output_df = self.output_df.head(1000)
            print("Warning: File truncated to 1000 rows")
            
        return True

    def save_excel(self, output_path):
        """Save the processed DataFrame to a new Excel file."""
        if self.output_df is None:
            return False
            
        try:
            self.output_df.to_excel(output_path, index=False)
            return True
        except Exception as e:
            print(f"Error saving Excel file: {e}")
            return False 