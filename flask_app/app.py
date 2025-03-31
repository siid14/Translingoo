import os
import uuid
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
from werkzeug.utils import secure_filename
import sys

# Add the src directory to the Python path so we can import the ExcelProcessor
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from src.excel_processor import ExcelProcessor

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-key-for-translingoo')
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['DOWNLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'downloads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max file size

# Ensure the upload and download directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

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
    
    # Get translation options
    columns_to_translate = []
    if request.form.get('translate_description'):
        columns_to_translate.append('Description')
    if request.form.get('translate_message'):
        columns_to_translate.append('Message')
    
    if not columns_to_translate:
        flash('Please select at least one column to translate', 'error')
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        # Create unique filename
        original_filename = secure_filename(file.filename)
        file_extension = original_filename.rsplit('.', 1)[1].lower()
        unique_id = str(uuid.uuid4())
        unique_filename = f"{unique_id}.{file_extension}"
        
        # Save the uploaded file
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(upload_path)
        
        # Process the file
        processor = ExcelProcessor()
        
        if not processor.load_excel(upload_path):
            flash('Error loading Excel file. Please check if the file is valid.', 'error')
            return redirect(url_for('index'))
        
        if not processor.process_file(columns_to_translate):
            flash('Error processing Excel file. Please check the console for details.', 'error')
            return redirect(url_for('index'))
        
        # Save the processed file
        output_filename = f"{unique_id}_translated.xlsx"
        output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
        
        if not processor.save_excel(output_path):
            flash('Error saving translated file', 'error')
            return redirect(url_for('index'))
        
        flash('File processed successfully!', 'success')
        return redirect(url_for('download_file', filename=output_filename, original_name=original_filename.replace('.' + file_extension, '_translated.xlsx')))
    
    flash('Invalid file type. Please upload an Excel file (.xls or .xlsx)', 'error')
    return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    original_name = request.args.get('original_name', filename)
    return render_template('download.html', filename=filename, original_name=original_name)

@app.route('/get_file/<filename>')
def get_file(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True, 
                              download_name=request.args.get('original_name', filename))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 