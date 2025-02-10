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
            
            # Try multiple approaches
            approaches = [
                # Approach 1: Basic read with no parameters
                lambda: pd.read_excel(file_path),
                
                # Approach 2: Read with openpyxl and different parameters
                lambda: pd.read_excel(file_path, engine='openpyxl', sheet_name=None),
                
                # Approach 3: Try reading with a specific sheet name
                lambda: pd.read_excel(file_path, sheet_name='Sheet1'),
                
                # Approach 4: Try reading with no header
                lambda: pd.read_excel(file_path, header=None),
                
                # Approach 5: Try reading with different header position
                lambda: pd.read_excel(file_path, header=13),
            ]
            
            last_error = None
            for i, approach in enumerate(approaches, 1):
                try:
                    print(f"\nDEBUG: Trying approach {i}")
                    result = approach()
                    
                    # If result is a dict (multiple sheets), take the first sheet
                    if isinstance(result, dict):
                        if result:
                            self.input_df = next(iter(result.values()))
                            print(f"DEBUG: Successfully read first sheet from dictionary")
                        else:
                            raise ValueError("No sheets found in the workbook")
                    else:
                        self.input_df = result
                    
                    print("DEBUG: Successfully read the file!")
                    print("\nDEBUG: File Content Preview:")
                    print(self.input_df.head())
                    print(f"\nDEBUG: DataFrame shape: {self.input_df.shape}")
                    print(f"DEBUG: Columns: {self.input_df.columns.tolist()}")
                    return True
                    
                except Exception as e:
                    print(f"DEBUG: Approach {i} failed: {str(e)}")
                    last_error = e
                    continue
            
            # If we get here, all approaches failed
            print("\nDEBUG: All approaches failed to read the file")
            print(f"Last error: {str(last_error)}")
            
            # One final attempt - try to read raw bytes to see if file is corrupted
            with open(file_path, 'rb') as f:
                header = f.read(8)  # Read first 8 bytes
                print(f"\nDEBUG: File header bytes: {header.hex()}")
            
            return False
            
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