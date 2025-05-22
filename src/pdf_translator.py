#!/usr/bin/env python3
"""
PDF Translator - A script to translate technical PDF documents
"""

import os
import re
import sys
import fitz  # PyMuPDF
import pandas as pd
from pathlib import Path
from tqdm import tqdm
from pdf2docx import Converter
import docx
from docx.shared import Pt, RGBColor

class PDFTranslator:
    def __init__(self, dictionary_path=None):
        """Initialize the PDF translator with optional custom dictionary path."""
        self.technical_terms = {
            # General electrical terms
            "Protection": "Protection",
            "Relay": "Relais",
            "Voltage": "Tension",
            "Current": "Courant",
            "Phase": "Phase",
            "Earth": "Terre",
            "Fault": "Défaut",
            "Over current": "Surintensité",
            "Earth fault": "Défaut à la terre",
            "Trip": "Déclenchement",
            "Rating": "Calibre",
            "Setting": "Réglage",
            "Curve": "Courbe",
            "Directional": "Directionnel",
            "Non Directional": "Non Directionnel",
            "Operating time": "Temps de fonctionnement",
            "Grading Margin": "Marge d'échelonnement",
            "CT Ratio": "Rapport TI",
            "VT Ratio": "Rapport TP",
            "Primary": "Primaire",
            "Secondary": "Secondaire",
            "Pickup": "Pickup",
            "Function": "Fonction",
            "Input": "Entrée",
            "Direction": "Direction",
            "Stage": "Étape",
            "Reset": "Réinitialisation",
            "Tripping characteristic": "Caractéristique de déclenchement",
            "Curve type": "Type de courbe",
            "Down stream": "En aval",
            "Grading": "Échelonnement",
            "Required": "Requis",
            "Maximum fault current": "Courant de défaut maximum",
            "Operate time": "Temps de fonctionnement",
            "Fault current": "Courant de défaut",
            "Calculation": "Calcul",
            "Considered": "Considéré",
            "Full load current": "Courant de pleine charge",
            "Date": "Date",
            "Revision": "Révision",
            "Prepared By": "Préparé Par",
            "Checked By": "Vérifié Par",
            "Approved By": "Approuvé Par",
            
            # Technical terms specific to the buscoupler document
            "SETTING CALCULATION DOCUMENT": "DOCUMENT DE CALCUL DE RÉGLAGE",
            "BUSCOUPLER PROTECTION": "PROTECTION DE COUPLEUR DE BARRES",
            "SETTING RECOMMENDATION": "RECOMMANDATION DE RÉGLAGE",
            "CONFIDENTIAL": "CONFIDENTIEL",
            "The information contained in this document is not to be communicated either directly or indirectly to any person not authorised to receive it": 
            "Les informations contenues dans ce document ne doivent être communiquées ni directement ni indirectement à toute personne non autorisée à les recevoir",
            
            "Customer": "Client",
            "Project": "Projet",
            "ALIMENTATION EN ENERGIE ELECTRIQUE EN UN POINT UNIQUE DE 2x30MVA PAR LIGNES PERSONNALISEES ET SECURISEES":
            "POWER SUPPLY AT A SINGLE POINT OF 2x30MVA BY CUSTOMIZED AND SECURED LINES",
            
            "Relay used": "Relais utilisé",
            "Rated voltage": "Tension nominale",
            "CT Ratio": "Rapport TC",
            "CTR-Pri": "CTR-Pri",
            "CTR-Sec": "CTR-Sec",
            "Full load current in Primary": "Courant de pleine charge au primaire",
            "Maximum fault current in Primary (Given)": "Courant de défaut maximal au primaire (donné)",
            
            "Over current protection": "Protection contre les surintensités",
            "Phase TOC-1 (I>1) Non Directional Over current: 51": "Phase TOC-1 (I>1) Surintensité non directionnelle: 51",
            "I>1 Function": "Fonction I>1",
            "I>1 Input": "Entrée I>1",
            "I>1 Direction": "Direction I>1",
            "Stage 1 Over current pickup in primary": "Pickup de surintensité étape 1 au primaire",
            "Stage 1 Over current pickup in secondary (I>1 Current Set)": "Pickup de surintensité étape 1 au secondaire (I>1 Réglage de courant)",
            "Tripping characteristic for I>1": "Caractéristique de déclenchement pour I>1",
            "Curve type (I>1 Curve)": "Type de courbe (I>1 Courbe)",
            "Down stream Operate time (Trafo Downstream)": "Temps de fonctionnement en aval (Trafo en aval)",
            "Grading Margin": "Marge d'échelonnement",
            "Required operating time": "Temps de fonctionnement requis",
            
            "Note: The IDMT curves shall saturate if the fault current is more than 20 times of pickup current. As in this case the actual fault current is 37800 A which is more than 20 times of pickup current hence for calculation purpose we shall consider": 
            "Remarque: Les courbes IDMT saturent si le courant de défaut est supérieur à 20 fois le courant de pickup. Dans ce cas, le courant de défaut réel est de 37800 A, ce qui est supérieur à 20 fois le courant de pickup, donc pour les besoins du calcul, nous considérerons",
            
            "Fault current considered for calculation": "Courant de défaut considéré pour le calcul",
            "Operating time @ TMS = 1": "Temps de fonctionnement @ TMS = 1",
            "TMS (I>1 TMS)": "TMS (I>1 TMS)",
            "I>1 Reset Char": "Caractéristique de réinitialisation I>1",
            "I>1 tReset": "I>1 tRéinitialisation",
            
            "Earth fault protection": "Protection contre les défauts à la terre",
            "EF1- TOC (IN>1) Non Directional Earth fault: 51N": "EF1- TOC (IN>1) Défaut à la terre non directionnel: 51N",
            "IN>1 Function": "Fonction IN>1",
            "IN>1 Direction": "Direction IN>1",
            "Stage 1 Earth fault pickup in primary": "Pickup de défaut à la terre étape 1 au primaire",
            "Stage 1 Earth fault pickup in secondary (IN1>1 Current)": "Pickup de défaut à la terre étape 1 au secondaire (Courant IN1>1)",
            "Tripping characteristic for IN>1": "Caractéristique de déclenchement pour IN>1",
            "Curve type (IN1>1 Curve)": "Type de courbe (Courbe IN1>1)",
            "LV (Down stream) Operate time (Trafo Downstream)": "Temps de fonctionnement BT (en aval) (Trafo en aval)",
            "TMS (IN1>1 TMS)": "TMS (TMS IN1>1)",
            "IN1>1 Reset Char": "Caractéristique de réinitialisation IN1>1",
            "IN1>1 tReset": "Réinitialisation IN1>1",
            
            # Units
            "kV": "kV",
            "A": "A",
            "s": "s",
        }

        # Load custom dictionary if provided
        if dictionary_path and os.path.exists(dictionary_path):
            try:
                custom_dict = pd.read_csv(dictionary_path)
                for _, row in custom_dict.iterrows():
                    self.technical_terms[row['english']] = row['french']
                print(f"Loaded {len(custom_dict)} custom translations")
            except Exception as e:
                print(f"Error loading custom dictionary: {str(e)}")

    def process_pdf(self, input_pdf, output_pdf):
        """Process a PDF file and create a translated version."""
        if not os.path.exists(input_pdf):
            print(f"Error: File {input_pdf} does not exist.")
            return False

        try:
            # Use the extract_to_excel and then convert back approach as it's more reliable
            # First, create a temporary Excel file
            temp_excel = f"{output_pdf}_temp.xlsx"
            success = self.extract_to_excel(input_pdf, temp_excel)
            
            if not success:
                print("Failed to extract PDF content")
                return False
                
            # Open the Excel file
            try:
                df = pd.read_excel(temp_excel)
                
                # Create a new PDF document
                doc = fitz.open()
                
                # Group by page number
                page_groups = df.groupby("Page")
                
                # Process each page
                for page_num, group in tqdm(page_groups, desc="Creating PDF pages"):
                    # Create a new page in the output document
                    page = doc.new_page(width=595, height=842)  # A4 size
                    
                    # Sort by position (if available) or just by index
                    sorted_group = group.sort_index()
                    
                    # Initialize y position for text
                    y_pos = 50
                    
                    # Add a header
                    page.insert_text((50, y_pos), f"Page {page_num} - Translated Content", fontsize=16)
                    y_pos += 30
                    
                    # Add content
                    for _, row in sorted_group.iterrows():
                        text_type = row["Type"]
                        french_text = row.get("French", "")
                        
                        if pd.notna(french_text) and french_text.strip():
                            # Add type as a subheader
                            page.insert_text((50, y_pos), f"{text_type}:", fontsize=12, color=(0, 0, 0.8))
                            y_pos += 20
                            
                            # Add text with word wrapping (simple approach)
                            words = french_text.split()
                            line = ""
                            for word in words:
                                test_line = line + " " + word if line else word
                                if len(test_line) > 70:  # Approximate line length
                                    page.insert_text((60, y_pos), line)
                                    y_pos += 15
                                    line = word
                                else:
                                    line = test_line
                            
                            if line:  # Add the last line
                                page.insert_text((60, y_pos), line)
                                y_pos += 25
                            
                            # Add some extra space after each text block
                            y_pos += 10
                            
                            # If near the bottom of the page, create a new page
                            if y_pos > 800:
                                page = doc.new_page(width=595, height=842)
                                y_pos = 50
                
                # Save the new PDF
                doc.save(output_pdf)
                doc.close()
                
                # Clean up the temp file
                try:
                    os.remove(temp_excel)
                except:
                    pass
                
                print(f"Successfully created translated PDF: {output_pdf}")
                return True
                
            except Exception as excel_error:
                print(f"Error processing Excel data: {str(excel_error)}")
                return False
            
        except Exception as e:
            print(f"Error processing PDF: {str(e)}")
            return False
    
    def _process_blocks(self, input_page, output_page, blocks):
        """Process text blocks and add translated content to output page."""
        # Create a drawing context for the output page
        shape = output_page.new_shape()
        
        for block in blocks:
            # Process based on block type
            if block["type"] == 0:  # Text block
                for line in block["lines"]:
                    line_text = ""
                    for span in line["spans"]:
                        # Get the text and its position
                        text = span["text"]
                        bbox = fitz.Rect(span["bbox"])
                        font_size = span["size"]
                        
                        # Translate the text
                        translated_text = self._translate_text(text)
                        
                        # Add translated text to output page - use a safe default font
                        try:
                            output_page.insert_text(
                                bbox.tl,  # top-left point
                                translated_text,
                                fontname="helv",  # Use built-in Helvetica font
                                fontsize=font_size,
                                color=(0, 0, 0)  # black
                            )
                        except Exception as e:
                            print(f"Warning: Could not add text '{translated_text}': {e}")
            
            elif block["type"] == 1:  # Image block
                # Images are handled separately
                pass
                
            # Copy table structures and borders
            if "lines" in block:
                for line in block.get("lines", []):
                    try:
                        p1 = line["p1"]
                        p2 = line["p2"]
                        shape.draw_line((p1[0], p1[1]), (p2[0], p2[1]))
                    except Exception as e:
                        print(f"Warning: Could not draw line: {e}")
        
        # Commit the shapes to the page
        shape.commit()
    
    def _translate_text(self, text):
        """Translate a piece of text from English to French."""
        if not text or text.strip() == "":
            return text
        
        # Check for direct matches in our technical terms dictionary
        if text in self.technical_terms:
            return self.technical_terms[text]
        
        # Normalize text for case-insensitive matching
        normalized_text = text.strip()
        upper_text = normalized_text.upper()
        
        # Check for case-insensitive match
        for eng, fr in self.technical_terms.items():
            if eng.upper() == upper_text:
                return fr
        
        # Try to translate individual terms within the text
        translated = text
        for eng, fr in sorted(self.technical_terms.items(), key=lambda x: len(x[0]), reverse=True):
            # Skip short terms to avoid incorrect replacements
            if len(eng) < 4:
                continue
                
            pattern = r'\b' + re.escape(eng) + r'\b'
            translated = re.sub(pattern, fr, translated, flags=re.IGNORECASE)
        
        return translated

    def extract_to_excel(self, input_pdf, output_excel):
        """Extract text content from PDF to Excel for translation."""
        if not os.path.exists(input_pdf):
            print(f"Error: File {input_pdf} does not exist.")
            return False
            
        try:
            # Open the PDF
            doc = fitz.open(input_pdf)
            
            # Create dataframes to store extracted text
            text_data = []
            table_data = []
            
            print(f"Extracting content from PDF: {input_pdf}")
            
            # Extract text from each page
            for page_num, page in enumerate(tqdm(doc, desc="Extracting pages")):
                # Extract regular text first (more reliable)
                try:
                    # Get text blocks
                    page_text = page.get_text("text")
                    if page_text.strip():
                        # Split by lines
                        lines = page_text.split('\n')
                        for line in lines:
                            if line.strip():
                                french_text = self._translate_text(line.strip())
                                text_data.append({
                                    "Page": page_num + 1,
                                    "Type": "Text",
                                    "English": line.strip(),
                                    "French": french_text
                                })
                except Exception as e:
                    print(f"Warning: Error extracting text from page {page_num + 1}: {e}")
                
                # Try to extract tables as a bonus
                try:
                    tables = self._extract_tables(page)
                    for table in tables:
                        for row in table:
                            table_data.append({
                                "Page": page_num + 1,
                                "Type": "Table",
                                "English": " | ".join(row),
                                "French": " | ".join([self._translate_text(cell) for cell in row])
                            })
                except Exception as e:
                    print(f"Warning: Error extracting tables from page {page_num + 1}: {e}")
            
            # Create DataFrame and save to Excel
            df_text = pd.DataFrame(text_data)
            df_table = pd.DataFrame(table_data)
            
            # Combine into one DataFrame
            if not df_text.empty or not df_table.empty:
                if df_text.empty:
                    df_combined = df_table
                elif df_table.empty:
                    df_combined = df_text
                else:
                    df_combined = pd.concat([df_text, df_table])
                
                # Sort by page number
                df_combined = df_combined.sort_values(by=["Page", "Type"])
                
                # Clean up data to avoid Excel corruption
                for col in df_combined.columns:
                    # Replace any characters that might cause Excel issues
                    if df_combined[col].dtype == 'object':  # Only process string columns
                        df_combined[col] = df_combined[col].apply(
                            lambda x: str(x).replace('\0', '').replace('\r', ' ').replace('\x00', '')
                            if pd.notna(x) else x
                        )
                
                # Save to Excel with more compatible options
                try:
                    # First, try saving with engine='openpyxl' and explicit options
                    with pd.ExcelWriter(output_excel, engine='openpyxl', mode='w') as writer:
                        df_combined.to_excel(writer, sheet_name="Translation", index=False)
                    print(f"Successfully saved extracted content to Excel: {output_excel}")
                    return True
                except Exception as e:
                    print(f"Warning: First Excel export attempt failed: {e}")
                    
                    # Try an alternative approach with fewer formatting options
                    try:
                        # Simplify column names
                        df_combined.columns = [str(col).replace(' ', '_') for col in df_combined.columns]
                        
                        # Save with xlsxwriter which can be more reliable
                        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                            df_combined.to_excel(writer, sheet_name="Translation", index=False)
                        print(f"Successfully saved extracted content to Excel (second attempt): {output_excel}")
                        return True
                    except Exception as e2:
                        print(f"Error saving Excel file (second attempt): {e2}")
                        
                        # If both export methods fail, try CSV as last resort
                        try:
                            csv_path = output_excel.replace('.xlsx', '.csv')
                            df_combined.to_csv(csv_path, index=False, encoding='utf-8-sig')
                            print(f"Exported as CSV instead: {csv_path}")
                            return True
                        except Exception as e3:
                            print(f"All export attempts failed: {e3}")
                            return False
            else:
                print("No content found in PDF")
                return False
            
        except Exception as e:
            print(f"Error extracting to Excel: {str(e)}")
            return False
    
    def _extract_tables(self, page):
        """
        Simple heuristic-based table extraction.
        Returns a list of tables, where each table is a list of rows.
        """
        # This is a simplified placeholder for table extraction
        # In a full implementation, you would need more sophisticated table detection
        tables = []
        
        try:
            # Look for tabular structures
            # For simplicity, we're looking for text aligned in columns
            text = page.get_text("dict")
            
            # Group text by y-position (rows)
            rows = {}
            for block in text["blocks"]:
                if block["type"] == 0:  # Text block
                    for line in block["lines"]:
                        y = int(line["bbox"][1])  # top y-coordinate
                        if y not in rows:
                            rows[y] = []
                        
                        for span in line["spans"]:
                            rows[y].append((span["bbox"][0], span["text"]))  # x-position and text
            
            # Sort rows by y-position
            sorted_rows = [rows[y] for y in sorted(rows.keys())]
            
            # Group consecutive rows that might form a table
            current_table = []
            
            for row in sorted_rows:
                # Sort spans by x-position
                sorted_spans = [span[1] for span in sorted(row, key=lambda x: x[0])]
                
                # Heuristic: If the row has multiple text elements, it might be a table row
                if len(sorted_spans) >= 2:
                    current_table.append(sorted_spans)
                elif current_table:
                    # End of table
                    if len(current_table) >= 2:  # At least 2 rows to form a table
                        tables.append(current_table)
                    current_table = []
            
            # Don't forget the last table
            if current_table and len(current_table) >= 2:
                tables.append(current_table)
        
        except Exception as e:
            print(f"Warning: Error in table extraction: {e}")
        
        return tables

    def process_pdf_to_word(self, input_pdf, output_docx):
        """Process a PDF file by converting to Word, translating, and saving as docx."""
        if not os.path.exists(input_pdf):
            print(f"Error: File {input_pdf} does not exist.")
            return False

        try:
            # Create temp files
            temp_docx = f"{output_docx}_original.docx"
            
            # Step 1: Convert PDF to Word
            print(f"Converting PDF to Word: {input_pdf}")
            cv = Converter(input_pdf)
            cv.convert(temp_docx)
            cv.close()
            
            if not os.path.exists(temp_docx):
                print("PDF to Word conversion failed")
                return False
                
            # Step 2: Translate the Word document
            print(f"Translating Word document")
            self._translate_word_document(temp_docx, output_docx)
            
            # Clean up temp file
            try:
                os.remove(temp_docx)
            except:
                pass
                
            print(f"Successfully created translated Word document: {output_docx}")
            return True
            
        except Exception as e:
            print(f"Error in PDF-to-Word processing: {str(e)}")
            return False
    
    def _translate_word_document(self, input_docx, output_docx):
        """Translate text in a Word document while preserving formatting."""
        try:
            # Load the document
            doc = docx.Document(input_docx)
            
            # Process paragraphs
            for para in tqdm(doc.paragraphs, desc="Translating paragraphs"):
                if para.text.strip():
                    # Store original formatting
                    runs_formatting = []
                    for run in para.runs:
                        runs_formatting.append({
                            'bold': run.bold,
                            'italic': run.italic,
                            'underline': run.underline,
                            'font_size': run.font.size,
                            'font_name': run.font.name,
                            'color': run.font.color.rgb if run.font.color and run.font.color.rgb else None
                        })
                    
                    # Translate the whole paragraph
                    translated_text = self._translate_text(para.text)
                    
                    # Clear paragraph and add translated text
                    para.clear()
                    para.add_run(translated_text)
                    
                    # If we had multiple runs with different formatting, try to preserve by splitting the text
                    if len(runs_formatting) > 1:
                        # This is a simplified approach - for better results, a more complex algorithm would be needed
                        para.clear()
                        words = translated_text.split()
                        runs_count = len(runs_formatting)
                        words_per_run = max(1, len(words) // runs_count)
                        
                        for i, fmt in enumerate(runs_formatting):
                            start_idx = i * words_per_run
                            end_idx = (i + 1) * words_per_run if i < runs_count - 1 else len(words)
                            if start_idx < len(words):
                                run_text = ' '.join(words[start_idx:end_idx])
                                run = para.add_run(run_text + ' ')
                                run.bold = fmt['bold']
                                run.italic = fmt['italic']
                                run.underline = fmt['underline']
                                if fmt['font_size']:
                                    run.font.size = fmt['font_size']
                                if fmt['font_name']:
                                    run.font.name = fmt['font_name']
                                if fmt['color']:
                                    run.font.color.rgb = fmt['color']
            
            # Process tables
            for table in tqdm(doc.tables, desc="Translating tables"):
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if para.text.strip():
                                # Translate the cell text
                                translated_text = self._translate_text(para.text)
                                para.clear()
                                para.add_run(translated_text)
            
            # Save the translated document
            doc.save(output_docx)
            return True
            
        except Exception as e:
            print(f"Error translating Word document: {str(e)}")
            return False

def main():
    """Main function to parse arguments and process PDF files."""
    import argparse
    
    parser = argparse.ArgumentParser(description='Process and translate technical PDF documents.')
    parser.add_argument('input_file', help='Path to the input PDF file')
    parser.add_argument('-o', '--output', help='Path to the output file (default: input_name_translated.pdf/xlsx/docx)')
    parser.add_argument('-e', '--excel', action='store_true', help='Extract content to Excel instead of creating a PDF')
    parser.add_argument('-w', '--word', action='store_true', help='Convert to Word document for better formatting')
    parser.add_argument('-d', '--dictionary', help='Path to a custom translation dictionary CSV file')
    
    args = parser.parse_args()
    
    # Create translator
    translator = PDFTranslator(dictionary_path=args.dictionary)
    
    if args.word:
        # Process with Word conversion
        if args.output:
            output_file = args.output
        else:
            basename = os.path.basename(args.input_file)
            name_without_ext = os.path.splitext(basename)[0]
            output_file = f"{name_without_ext}_translated.docx"
        
        success = translator.process_pdf_to_word(args.input_file, output_file)
    elif args.excel:
        # Extract to Excel
        if args.output:
            output_file = args.output
        else:
            basename = os.path.basename(args.input_file)
            name_without_ext = os.path.splitext(basename)[0]
            output_file = f"{name_without_ext}_translation.xlsx"
        
        success = translator.extract_to_excel(args.input_file, output_file)
    else:
        # Create translated PDF
        if args.output:
            output_file = args.output
        else:
            basename = os.path.basename(args.input_file)
            name_without_ext = os.path.splitext(basename)[0]
            output_file = f"{name_without_ext}_translated.pdf"
        
        success = translator.process_pdf(args.input_file, output_file)
    
    if success:
        print("Processing completed successfully.")
        sys.exit(0)
    else:
        print("Processing failed.")
        sys.exit(1)

if __name__ == "__main__":
    main() 