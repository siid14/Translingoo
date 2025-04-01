#!/usr/bin/env python3
"""
Excel Converter - A simple script to convert Excel files and process translations
"""

import pandas as pd
import os
import sys
import argparse
import re
import warnings

# Suppress warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=UserWarning, module='pandas')

def convert_and_process(input_file, output_file=None, columns_to_translate=None):
    """Convert an Excel file and apply translations."""
    
    if not os.path.exists(input_file):
        print(f"Error: File {input_file} does not exist.")
        return False
    
    if output_file is None:
        # Create output filename by replacing extension with _translated.xlsx
        basename = os.path.basename(input_file)
        name_without_ext = os.path.splitext(basename)[0]
        output_file = f"{name_without_ext}_translated.xlsx"
    
    print(f"Processing {input_file} to {output_file}...")
    
    # Default columns to translate if none specified
    if columns_to_translate is None:
        columns_to_translate = ["Description", "Message"]
    
    # Try different engines to read the Excel file
    engines = ['xlrd', 'openpyxl', 'odf']
    
    for engine in engines:
        try:
            print(f"Trying with engine: {engine}")
            
            # First attempt: Just read the file directly
            df = pd.read_excel(input_file, engine=engine)
            
            # Check if we have data and the required columns
            if len(df) > 0:
                print(f"Successfully loaded with {engine}")
                print(f"Columns: {df.columns.tolist()}")
                print(f"Shape: {df.shape}")
                
                # Find the actual columns to translate (case insensitive)
                columns_map = {}
                for col in columns_to_translate:
                    found = False
                    for df_col in df.columns:
                        if col.lower() == str(df_col).lower():
                            columns_map[col] = df_col
                            found = True
                            break
                    if not found:
                        print(f"Warning: Column '{col}' not found in the Excel file")
                
                if columns_map:
                    # Apply translations to each column
                    for original_col, actual_col in columns_map.items():
                        print(f"Translating column: {actual_col}")
                        
                        # Create a new column for translations
                        new_column_name = f"{actual_col} Français"
                        
                        # Apply translation to each cell
                        df[new_column_name] = df[actual_col].apply(lambda x: translate_text(x))
                        
                    # Save the result
                    df.to_excel(output_file, index=False)
                    print(f"Successfully saved translated file: {output_file}")
                    return True
                else:
                    print("No columns to translate were found in the file")
                    return False
            
        except Exception as e:
            print(f"Failed with engine {engine}: {str(e)}")
            
            # Try with additional options
            try:
                print(f"Trying with {engine} and header detection...")
                # Read without headers first
                raw_df = pd.read_excel(input_file, engine=engine, header=None)
                
                if len(raw_df) > 0:
                    # Find the header row
                    header_row = None
                    for i in range(min(20, len(raw_df))):
                        row_str = ' '.join(raw_df.iloc[i].astype(str).values)
                        if 'Description' in row_str and ('Message' in row_str or 'Type' in row_str):
                            header_row = i
                            break
                    
                    if header_row is not None:
                        print(f"Found header at row {header_row}")
                        df = pd.read_excel(input_file, engine=engine, header=header_row)
                        
                        # Check if we have the columns to translate
                        columns_map = {}
                        for col in columns_to_translate:
                            found = False
                            for df_col in df.columns:
                                if col.lower() == str(df_col).lower():
                                    columns_map[col] = df_col
                                    found = True
                                    break
                            if not found:
                                print(f"Warning: Column '{col}' not found in the Excel file")
                        
                        if columns_map:
                            # Apply translations to each column
                            for original_col, actual_col in columns_map.items():
                                print(f"Translating column: {actual_col}")
                                
                                # Create a new column for translations
                                new_column_name = f"{actual_col} Français"
                                
                                # Apply translation to each cell
                                df[new_column_name] = df[actual_col].apply(lambda x: translate_text(x))
                            
                            # Save the result
                            df.to_excel(output_file, index=False)
                            print(f"Successfully saved translated file: {output_file}")
                            return True
                        else:
                            print("No columns to translate were found in the file")
                            return False
            except Exception as sub_e:
                print(f"Failed with {engine} and header detection: {str(sub_e)}")
    
    print("All processing methods failed.")
    return False

def translate_text(text):
    """Translate text from English to French."""
    if pd.isna(text) or text is None or str(text).strip() == '':
        return text
        
    # Dictionary of translations
    translations = {
        # Common values in Description and Message columns
        'ABSENCE OF REFERENCE VOLTAGE': 'ABSENCE DE TENSION DE REFERENCE',
        'ABSENCE OF VOLTAGE': 'ABSENCE DE TENSION',
        'DEAD INCOMING DEAD RUNNING': 'ENTRÉE HORS TENSION, FONCTIONNEMENT HORS TENSION',  
        'DEAD INCOMING  DEAD RUNNING': 'ENTRÉE HORS TENSION, FONCTIONNEMENT HORS TENSION',  
        'LIVE INCOMING LIVE RUNNING': 'ENTRÉE SOUS TENSION, FONCTIONNEMENT SOUS TENSION',   
        'LIVE INCOMING  LIVE RUNNING': 'ENTRÉE SOUS TENSION, FONCTIONNEMENT SOUS TENSION',  
        'CLOSE PERMISSIVE': 'AUTORISATION DE FERMETURE',
        'PRESENE OF REFERENCE VOLTAGE': 'PRÉSENCE DE TENSION DE RÉFÉRENCE',
        'PRASENCE OF VOLTAGE': 'PRÉSENCE DE TENSION',
        'LIVE INCOMING  DEAD RUNNING': 'ENTRÉE SOUS TENSION, FONCTIONNEMENT HORS TENSION',
        'POSSIBLE CLOSING': 'FERMETURE AUTORISEE',
        'BUS-1 SELECT': 'SÉLECTION DU BUS-1',
        'BUS-2 DESELECT': 'DÉSÉLECTION DU BUS-2',
        'BAY L/R MODE': 'MODE LOCAL/DISTANT DE LA TRAVÉE',
        'BAY L/R  MODE': 'MODE L/R TRAVEE',
        'IINTERLOCK PERMISSIVE': 'VERROUILLAGE AUTORISÉ',
        'CARRIER IN': 'PORTEUSE ENTRANTE',
        'CARRIER OUT': 'PORTEUSE SORTANTE',
        'EQUIPMENT BCU': 'EQUIPEMENT BCU',
        'COMMUNICATION': 'COMMUNICATION',
        'UNLOCKING RELAY ACTIVATED': 'RELAIS DEVERROUILLAGE ACTIVE',
        'PROTECTION': 'PROTECTION',
        'STAGE START': 'DEMARRAGE ETAPE',
        'STAGE': 'ETAPE',
        'TRIP CIRCUIT FAULT': 'DEFAUT CIRCUIT DE DECLENCHEMENT',
        'I/L PERMISSIVE': 'PERMISSIF V/F',
        'CLOSE I/L PERMISSIVE': 'PERMISSIF V/F FERMETURE',
        'OPEN I/L PERMISSIVE': 'PERMISSIF V/F OUVERTURE',
        
        # Alarm/status values typically in Message column
        'Set': 'Réglé',
        'Set - App Ack': 'Réglé - App Ack',
        'SET': 'RÉGLÉ',
        'SET - APP ACK': 'RÉGLÉ - APP ACK',
        'Reset': 'Réinitialiser',
        'Reset - App Ack': 'Réinitialiser - App Ack',
        'RESET': 'RÉINITIALISER',
        'RESET - APP ACK': 'RÉINITIALISER - APP ACK',
        'Operational': 'Opérationnel',
        'OPERATIONAL': 'OPÉRATIONNEL',
        'Alarm': 'Alarme',
        'ALARM': 'ALARME',
        'Normal': 'Normal',
        'NORMAL': 'NORMAL',
        'Operated': 'Opéré',
        'OPERATED': 'OPÉRÉ',
        'TRIP - App Ack': 'DÉCLENCHEMENT - App Ack',
        'OPEN - App Ack': 'OUVERT - App Ack',
        'OPEN - Clearing': 'OUVERT - Effacement',
        'Fail - App Ack': 'Défaillance - App Ack',
        'Fail - Clearing': 'Défaillance - Effacement',
        'Faulty - Clearing': 'Défectueux - Effacement',
        'Operated - Clearing': 'Opéré - Effacement',
        'Operated - App Ack': 'Opéré - App Ack',
        'OFF - App Ack': 'DÉSACTIVÉ - App Ack',
        'Off - App Ack': 'Désactivé - App Ack',
        'Trip': 'Déclenchement',
        'Closed': 'Fermé',
        'On Sync': 'En Synchronisation',
        'Healthy': 'En Bon État',
        'Operational Mode': 'Mode Opérationnel',
        'Fail': 'Défaillance',
        'Faulty': 'Défectueux',
        'Off': 'Désactivé'
    }
    
    # Try to find a direct match
    text_str = str(text).strip().upper()
    
    if text_str in translations:
        return translations[text_str]
    
    # Try with normalized spacing
    normalized_text = ' '.join(text_str.split())
    for key, value in translations.items():
        if normalized_text == ' '.join(key.upper().split()):
            return value
    
    # If no translation found, return original
    return text

def main():
    """Main function to parse arguments and process Excel files."""
    parser = argparse.ArgumentParser(description='Process Excel files and apply translations.')
    parser.add_argument('input_file', help='Path to the input Excel file (.xls or .xlsx)')
    parser.add_argument('-o', '--output', help='Path to the output Excel file (default: input_name_translated.xlsx)')
    parser.add_argument('-c', '--columns', nargs='+', help='Columns to translate (default: Description Message)')
    parser.add_argument('--skip-description', action='store_true', help='Skip translating the Description column')
    parser.add_argument('--skip-message', action='store_true', help='Skip translating the Message column')
    
    args = parser.parse_args()
    
    # Handle column selection based on skip arguments
    columns_to_translate = args.columns
    
    if columns_to_translate is None:
        columns_to_translate = []
        if not args.skip_description:
            columns_to_translate.append("Description")
        if not args.skip_message:
            columns_to_translate.append("Message")
        
        # If both columns are skipped, use an empty list which will result in an error
        if not columns_to_translate:
            print("Error: Cannot skip all translation columns. At least one column must be translated.")
            sys.exit(1)
    
    print(f"Columns to translate: {', '.join(columns_to_translate)}")
    
    success = convert_and_process(args.input_file, args.output, columns_to_translate)
    
    if success:
        print("Processing completed successfully.")
        sys.exit(0)
    else:
        print("Processing failed.")
        sys.exit(1)

if __name__ == "__main__":
    main() 