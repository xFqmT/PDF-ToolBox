from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import os
import tempfile
import uuid

app = Flask(__name__)

# Simple configuration
app.config['UPLOAD_FOLDER'] = '/tmp'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

@app.route("/")
def home():
    return render_template('index.html')

@app.route("/api")
def api():
    return "Hello, API!"

@app.route('/merge', methods=['POST'])
def merge():
    files = request.files.getlist('file')
    if not files:
        return "No files uploaded", 400
    
    merger = PdfMerger()
    temp_files = []
    
    try:
        for file in files:
            if file.filename.endswith('.pdf'):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
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
    pages = request.form.get('pages', '1')
    
    if not file:
        return "No file uploaded", 400
    
    try:
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)
        
        reader = PdfReader(input_path)
        writer = PdfWriter()
        
        # Simple page extraction (just first page for now)
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

if __name__ == "__main__":
    app.run()
