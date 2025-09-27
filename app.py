from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import os
import tempfile
import uuid
import io

app = Flask(__name__)

# Simple configuration
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB limit

@app.route("/")
def home():
    return render_template('index.html')

@app.route('/merge', methods=['POST'])
def merge():
    # Handle multiple files from the advanced merge form
    files = []
    file_index = 0
    
    # Collect all uploaded files
    while f'file_{file_index}' in request.files:
        file = request.files[f'file_{file_index}']
        if file and file.filename:
            files.append(file)
        file_index += 1
    
    if len(files) < 2:
        return "Please select at least 2 PDF files to merge.", 400
    
    merger = PdfMerger()
    temp_files = []
    
    try:
        for file in files:
            if file.filename.lower().endswith('.pdf'):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"{uuid.uuid4().hex}_{filename}")
                file.save(filepath)
                temp_files.append(filepath)
                merger.append(filepath)
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"merged_{uuid.uuid4().hex}.pdf")
        merger.write(output_path)
        merger.close()
        
        # Cleanup temp files
        for temp_file in temp_files:
            try:
                os.remove(temp_file)
            except:
                pass
        
        return send_file(output_path, as_attachment=True, download_name="merged.pdf")
    except Exception as e:
        return f"Error: {str(e)}", 500

@app.route('/split', methods=['POST'])
def split():
    file = request.files.get('file')
    pages_input = request.form.get('pages', '').strip()
    
    if not file:
        return "No file uploaded", 400
    
    try:
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{uuid.uuid4().hex}_{filename}")
        file.save(input_path)
        
        reader = PdfReader(input_path)
        writer = PdfWriter()
        
        # Parse page ranges
        if pages_input:
            page_numbers = parse_page_ranges(pages_input, len(reader.pages))
            for page_num in page_numbers:
                if 0 <= page_num < len(reader.pages):
                    writer.add_page(reader.pages[page_num])
        else:
            # If no pages specified, extract first page
            if len(reader.pages) > 0:
                writer.add_page(reader.pages[0])
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"split_{uuid.uuid4().hex}.pdf")
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        # Cleanup
        try:
            os.remove(input_path)
        except:
            pass
        
        return send_file(output_path, as_attachment=True, download_name="split.pdf")
    except Exception as e:
        return f"Error: {str(e)}", 500

def parse_page_ranges(pages_str, max_pages):
    """Parse page ranges like '1-3,5,7' into list of page numbers (0-indexed)"""
    page_numbers = []
    parts = pages_str.split(',')
    
    for part in parts:
        part = part.strip()
        if '-' in part:
            start, end = part.split('-', 1)
            start = int(start.strip()) - 1  # Convert to 0-indexed
            end = int(end.strip()) - 1      # Convert to 0-indexed
            for i in range(start, min(end + 1, max_pages)):
                if i >= 0:
                    page_numbers.append(i)
        else:
            page_num = int(part.strip()) - 1  # Convert to 0-indexed
            if 0 <= page_num < max_pages:
                page_numbers.append(page_num)
    
    return sorted(list(set(page_numbers)))  # Remove duplicates and sort

# Placeholder routes for other tools mentioned in the template
@app.route('/compress', methods=['POST'])
def compress():
    return "Compress feature not implemented yet", 501

@app.route('/pdf_to_word', methods=['POST'])
def pdf_to_word():
    return "PDF to Word feature not implemented yet", 501

@app.route('/pdf_to_images', methods=['POST'])
def pdf_to_images():
    return "PDF to Images feature not implemented yet", 501

@app.route('/images_to_pdf', methods=['POST'])
def images_to_pdf():
    return "Images to PDF feature not implemented yet", 501

if __name__ == "__main__":
    app.run()
