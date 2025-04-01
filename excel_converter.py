#!/usr/bin/env python3
"""
Excel Converter - A simple script to convert Excel files to CSV
"""

import pandas as pd
import os
import sys
import argparse

def convert_excel_to_csv(input_file, output_file=None, sheet_name=0):
    """Convert an Excel file to CSV format."""
    
    if not os.path.exists(input_file):
        print(f"Error: File {input_file} does not exist.")
        return False
    
    if output_file is None:
        # Create output filename by replacing extension with .csv
        basename = os.path.basename(input_file)
        name_without_ext = os.path.splitext(basename)[0]
        output_file = f"{name_without_ext}.csv"
    
    print(f"Converting {input_file} to {output_file}...")
    
    # Try different engines to read the Excel file
    engines = ['openpyxl', 'xlrd', 'odf']
    
    for engine in engines:
        try:
            print(f"Trying with engine: {engine}")
            
            # First attempt: Just read the file directly
            df = pd.read_excel(input_file, engine=engine, sheet_name=sheet_name)
            
            # Save to CSV
            df.to_csv(output_file, index=False)
            print(f"Successfully converted to CSV format: {output_file}")
            print(f"CSV file shape: {df.shape}")
            return True
            
        except Exception as e:
            print(f"Failed with engine {engine}: {str(e)}")
            
            # Try with additional options
            try:
                print(f"Trying with {engine} and header=None...")
                df = pd.read_excel(input_file, engine=engine, sheet_name=sheet_name, header=None)
                
                # Check if we found content
                if len(df) > 0:
                    # Find the header row (usually within first 20 rows)
                    header_row = None
                    for i in range(min(20, len(df))):
                        row_str = ' '.join(df.iloc[i].astype(str).values)
                        if 'Description' in row_str and ('Message' in row_str or 'Type' in row_str):
                            header_row = i
                            break
                    
                    if header_row is not None:
                        # Use this row as header
                        new_df = pd.DataFrame(df.values[header_row+1:], columns=df.iloc[header_row])
                        new_df.to_csv(output_file, index=False)
                        print(f"Successfully converted to CSV format with header detection: {output_file}")
                        print(f"CSV file shape: {new_df.shape}")
                        return True
                    else:
                        # Use default headers
                        df.to_csv(output_file, index=False)
                        print(f"Successfully converted to CSV format with default headers: {output_file}")
                        print(f"CSV file shape: {df.shape}")
                        return True
            except Exception as sub_e:
                print(f"Failed with {engine} and header=None: {str(sub_e)}")
    
    print("All conversion methods failed.")
    return False

def main():
    """Main function to parse arguments and convert Excel files."""
    parser = argparse.ArgumentParser(description='Convert Excel files to CSV format.')
    parser.add_argument('input_file', help='Path to the input Excel file (.xls or .xlsx)')
    parser.add_argument('-o', '--output', help='Path to the output CSV file (default: same name with .csv extension)')
    parser.add_argument('-s', '--sheet', type=int, default=0, help='Sheet index to convert (default: 0)')
    
    args = parser.parse_args()
    
    success = convert_excel_to_csv(args.input_file, args.output, args.sheet)
    
    if success:
        print("Conversion completed successfully.")
        sys.exit(0)
    else:
        print("Conversion failed.")
        sys.exit(1)

if __name__ == "__main__":
    main() 