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
            
        try:
            # Technical terms translation dictionary (English to French)
            translations = {
                # Existing translations
                'ABSENCE OF REFERENCE VOLTAGE': 'ABSENCE DE TENSION DE REFERENCE',
                'ABSENCE OF VOLTAGE': 'ABSENCE DE TENSION',
                'DEAD INCOMING DEAD RUNNING': 'ENTRÉE HORS TENSION, FONCTIONNEMENT HORS TENSION',  # Fixed spacing
                'DEAD INCOMING  DEAD RUNNING': 'ENTRÉE HORS TENSION, FONCTIONNEMENT HORS TENSION',  # Extra space variant
                
                'LIVE INCOMING LIVE RUNNING': 'ENTRÉE SOUS TENSION, FONCTIONNEMENT SOUS TENSION',   # Fixed spacing
                'LIVE INCOMING  LIVE RUNNING': 'ENTRÉE SOUS TENSION, FONCTIONNEMENT SOUS TENSION',  # Extra space variant
                'CLOSE PERMISSIVE': 'AUTORISATION DE FERMETURE',
                
                # New translations
                'PRESENE OF REFERENCE VOLTAGE': 'PRÉSENCE DE TENSION DE RÉFÉRENCE',
                'PRASENCE OF VOLTAGE': 'PRÉSENCE DE TENSION',
                'LIVE INCOMING  DEAD RUNNING': 'ENTRÉE SOUS TENSION, FONCTIONNEMENT HORS TENSION',
                'POSSIBLE CLOSING': 'FERMETURE AUTORISEE',
                'BUS-1 SELECT': 'SÉLECTION DU BUS-1',
                'BUS-2 DESELECT': 'DÉSÉLECTION DU BUS-2',  # New translation
                'BAY L/R MODE': 'MODE L/R TRAVEE',  # Moving to top section for visibility
                'BAY L/R  MODE': 'MODE L/R TRAVEE',  # Extra space variant
                'IINTERLOCK PERMISSIVE': 'AUTORISATION DE VERROUILLAGE',
                
                # Additional new translations
                'CARRIER IN': 'PORTEUSE ENTRANTE',
                'CARRIER OUT': 'PORTEUSE SORTANTE',
                'EQUIPMENT BCU': 'EQUIPEMENT BCU',  # Bloc de Contrôle Unité
                'COMMUNICATION': 'COMMUNICATION',
                'UNLOCKING RELAY ACTIVATED': 'RELAIS DEVERROUILLAGE ACTIVE',
                'PROTECTION': 'PROTECTION',
                'STAGE START': 'DEMARRAGE ETAPE',
                'STAGE': 'ETAPE',
                'SEND CH-1': 'ENVOI CANAL 1',
                'SEND CH-2': 'ENVOI CANAL 2',
                'ZONE PROTECTION': 'PROTECTION DE ZONE',
                'BPH PROTECTION': 'PROTECTION BPH',  # Bus Phase Protection Haute
                'RPH PROTECTION': 'PROTECTION RPH',  # Remote Phase Protection Haute
                'YPH PROTECTION': 'PROTECTION YPH',  # Yard Phase Protection Haute
                
                # Overcurrent protection stages
                '50/51 STAGE-1': 'ETAPE-1 50/51',  # Overcurrent protection stage 1
                '50/51 STAGE-1 START': 'DEMARRAGE ETAPE-1 50/51',
                '50/51 STAGE-2': 'ETAPE-2 50/51',
                '50/51 STAGE-2 START': 'DEMARRAGE ETAPE-2 50/51',
                
                # Communication
                'DT SEND CH-1': 'ENVOI CANAL-1 DT',
                'DT SEND CH-2': 'ENVOI CANAL-2 DT',
                
                # Directional Protection
                '67N STAGE-1': 'ETAPE-1 67N',
                '67N STAGE-1 START': 'DEMARRAGE ETAPE-1 67N',
                
                # Zone Protection
                '21 ZONE-1 PROTECTION': 'PROTECTION ZONE-1 21',
                '21 ZONE-2 PROTECTION': 'PROTECTION ZONE-2 21',
                '21 ZONE-3 PROTECTION': 'PROTECTION ZONE-3 21',
                '21 ZONE-4 PROTECTION': 'PROTECTION ZONE-4 21',
                '21 ZONE-1 PROTECTION START': 'DEMARRAGE PROTECTION ZONE-1 21',
                '21 ZONE-2 PROTECTION START': 'DEMARRAGE PROTECTION ZONE-2 21',
                '21 ZONE-3 PROTECTION START': 'DEMARRAGE PROTECTION ZONE-3 21',
                '21 ZONE-4 PROTECTION START': 'DEMARRAGE PROTECTION ZONE-4 21',
                '21 ZONE-1 YPH PROTECTION': 'PROTECTION YPH ZONE-1 21',
                '21 ZONE-1 BPH PROTECTION': 'PROTECTION BPH ZONE-1 21',
                
                # Differential Protection
                '87L PROTECTION': 'PROTECTION 87L',
                '87L PROTECTION START': 'DEMARRAGE PROTECTION 87L',
                '87L PROTECTION A-PH': 'PROTECTION 87L PHASE-A',
                
                # Rest of existing translations
                'ON/OFF SECONDARY SPS': 'MARCHE/ARRET SPS SECONDAIRE',
                'ON/OFF MAIN FOR CB1': 'MARCHE/ARRET PRINCIPAL POUR CB1',
                'DUMMY': 'FACTICE/RESERVE',
                'COMP. POSITION': 'POSITION COMP.',
                'COMP_POSITION': 'POSITION_COMP',
                'DISCONNECTOR G1 POSITION': 'POSITION SECTIONNEUR G1',
                'DISCONNECTOR G2 POSITION': 'POSITION SECTIONNEUR G2',
                'DISCONNECTOR G3 POSITION': 'POSITION SECTIONNEUR G3',
                'TRIP CIRCUIT FAULT': 'DEFAUT CIRCUIT DE DECLENCHEMENT',
                'I/L PERMISSIVE': 'PERMISSIF V/F',
                'CLOSE I/L PERMISSIVE': 'PERMISSIF V/F FERMETURE',
                'OPEN I/L PERMISSIVE': 'PERMISSIF V/F OUVERTURE',
                'CIRCUIT BREAKER GCB1 POSITION': 'DISJONCTEUR GCB1 POSITION',
                'CIRCUIT BREAKER-GCB1 POS': 'DISJONCTEUR-GCB1 POS'
            }
            
            print("\nDEBUG: Starting translation of Description column")
            
            def is_french(text):
                """Check if text contains French-specific words/patterns"""
                french_indicators = [
                    'DE', 'DES', 'PERMISSIF', 'SECTIONNEUR', 'TERRE', 'DISJONCTEUR', 
                    'MARCHE', 'ARRET', 'ENTREE', 'INACTIVE', 'EXECUTION', 'COMMANDE',
                    'MANUELLE', 'PORTEUSE', 'ENTRANTE', 'SORTANTE', 'EQUIPEMENT',
                    'RELAIS', 'DEVERROUILLAGE', 'DEMARRAGE', 'ETAPE', 'ENVOI', 'CANAL',
                    'PROTECTION', 'PHASE', 'ZONE'
                ]
                return any(indicator in text.upper() for indicator in french_indicators)
            
            def translate_text(text):
                if pd.isna(text) or text is None or str(text).strip() == '':
                    return text
                    
                # Normalize the text by removing extra spaces
                text_str = ' '.join(str(text).strip().upper().split())
                
                # First check if the text is French
                if is_french(text_str):
                    print(f"DEBUG: Keeping French text: '{text}'")
                    return text
                
                # If not French, try to translate from English to French
                if text_str in translations:
                    translated = translations[text_str]
                    print(f"DEBUG: Translated '{text}' → '{translated}'")
                    return translated
                
                # Try with original spacing if normalized version not found
                original_upper = str(text).strip().upper()
                if original_upper in translations:
                    translated = translations[original_upper]
                    print(f"DEBUG: Translated '{text}' → '{translated}'")
                    return translated
                
                # If no translation found, return original
                print(f"DEBUG: No translation found for '{text}', keeping original")
                return text
            
            # Apply translation
            self.output_df['Description'] = self.output_df['Description'].apply(translate_text)
            print("DEBUG: Translation completed")
            
            # Limit to 1000 rows if necessary
            if len(self.output_df) > 1000:
                self.output_df = self.output_df.head(1000)
                print("DEBUG: Truncated to 1000 rows")
            
            return True
            
        except Exception as e:
            print(f"\nDEBUG: Error during translation:")
            print(f"- Error type: {type(e).__name__}")
            print(f"- Error message: {str(e)}")
            return False

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