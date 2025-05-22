#!/usr/bin/env python3
"""
PDF Translator Web App - A Flask-based web interface for the PDF Translator
"""

import os
import uuid
import sys
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
from werkzeug.utils import secure_filename

# Import our PDF Translator
from pdf_translator import PDFTranslator

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'pdf-translator-translingoo')
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), '..', 'pdf_uploads')
app.config['DOWNLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), '..', 'pdf_downloads')
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB max file size

# Ensure the upload and download directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('pdf_index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # Check if the form has files
    if 'file' not in request.files:
        flash('No file part', 'error')
        return redirect(request.url)
    
    file = request.files['file']
    
    # If user doesn't select file, browser submits empty file
    if file.filename == '':
        flash('No selected file', 'error')
        return redirect(request.url)
    
    # Get translation options - default to Excel for better reliability
    output_type = request.form.get('output_type', 'excel')
    
    if file and allowed_file(file.filename):
        try:
            # Create unique filename
            original_filename = secure_filename(file.filename)
            file_extension = original_filename.rsplit('.', 1)[1].lower()
            unique_id = str(uuid.uuid4())
            unique_filename = f"{unique_id}.{file_extension}"
            
            # Save the uploaded file
            upload_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(upload_path)
            
            # Process the file
            translator = PDFTranslator()
            
            # Determine output path and format
            if output_type == 'word':
                output_filename = f"{unique_id}_translated.docx"
                output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
                success = translator.process_pdf_to_word(upload_path, output_path)
                original_download_name = original_filename.replace('.' + file_extension, '_translated.docx')
            elif output_type == 'excel' or output_type == 'auto':
                output_filename = f"{unique_id}_translation.xlsx"
                output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
                success = translator.extract_to_excel(upload_path, output_path)
                
                # Check if the Excel file exists, if not, look for CSV fallback
                if success and not os.path.exists(output_path):
                    csv_path = output_path.replace('.xlsx', '.csv')
                    if os.path.exists(csv_path):
                        # Update filename and path to use CSV instead
                        output_filename = output_filename.replace('.xlsx', '.csv')
                        output_path = csv_path
                        original_download_name = original_filename.replace('.' + file_extension, '_translation.csv')
                        flash('Excel export was converted to CSV format for compatibility.', 'warning')
                
                original_download_name = original_filename.replace('.' + file_extension, '_translation.xlsx')
            else:
                output_filename = f"{unique_id}_translated.pdf"
                output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
                success = translator.process_pdf(upload_path, output_path)
                original_download_name = original_filename.replace('.' + file_extension, '_translated.pdf')
            
            if not success:
                flash('Error processing PDF file. Trying alternative format...', 'warning')
                # Try Word format as a fallback for all other formats
                output_filename = f"{unique_id}_translated.docx"
                output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
                success = translator.process_pdf_to_word(upload_path, output_path)
                original_download_name = original_filename.replace('.' + file_extension, '_translated.docx')
                
                if success:
                    flash('Original format failed, but Word document generation succeeded.', 'warning')
                else:
                    # Try Excel as last resort
                    output_filename = f"{unique_id}_translation.xlsx"
                    output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
                    success = translator.extract_to_excel(upload_path, output_path)
                    original_download_name = original_filename.replace('.' + file_extension, '_translation.xlsx')
                    if success:
                        flash('Word generation failed, but Excel extraction succeeded.', 'warning')
                
            if not success:
                flash('Error processing PDF file. Please try a different file or format.', 'error')
                return redirect(url_for('index'))
            
            flash('File processed successfully!', 'success')
            return redirect(url_for('download_file', filename=output_filename, original_name=original_download_name))
        
        except Exception as e:
            flash(f'An error occurred: {str(e)}', 'error')
            return redirect(url_for('index'))
    
    flash('Invalid file type. Please upload a PDF file.', 'error')
    return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    original_name = request.args.get('original_name', filename)
    return render_template('pdf_download.html', filename=filename, original_name=original_name)

@app.route('/get_file/<filename>')
def get_file(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True, 
                              download_name=request.args.get('original_name', filename))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)  # Use different port from the Excel translator 