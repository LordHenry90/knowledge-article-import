import os
import base64
import hashlib
import logging
from flask import Flask, render_template, request, send_file, redirect, url_for
from werkzeug.utils import secure_filename
import mammoth
import csv
import zipfile
import shutil
from bs4 import BeautifulSoup

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DATA_FOLDER'] = 'data'
app.config['IMAGES_FOLDER'] = os.path.join(app.config['DATA_FOLDER'], 'images')
app.config['OUTPUT_ZIP'] = 'KnowledgeArticlesImport.zip'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB limit

# Configure logging
logging.basicConfig(filename='app.log', level=logging.DEBUG)

# ... [rest of the functions remain the same] ...

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        try:
            if 'file' not in request.files:
                return redirect(request.url)
            files = request.files.getlist('file')
            if not files or files[0].filename == '':
                return redirect(request.url)
            
            # Clear previous data
            clear_files()
            
            for folder in [app.config['UPLOAD_FOLDER'], app.config['DATA_FOLDER'], app.config['IMAGES_FOLDER']]:
                os.makedirs(folder, exist_ok=True)
            
            filenames = []
            messages = []
            for file in files:
                if file and file.filename.endswith('.docx'):
                    filename = secure_filename(file.filename)
                    file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                    filenames.append(filename)
            
            create_content_properties()
            
            html_files = []
            for filename in filenames:
                html_file, msg = convert_docx_to_html(filename)
                html_files.append(html_file)
                messages.extend(msg)
            
            csv_records = [create_csv_record(html_file) for html_file in html_files]
            csv_file = create_csv_file(csv_records)
            
            files_to_zip = filenames + [csv_file]
            create_zip_file(files_to_zip)
            
            return send_file(app.config['OUTPUT_ZIP'], as_attachment=True)
        except Exception as e:
            app.logger.error(f"An error occurred: {str(e)}", exc_info=True)
            return f"An error occurred: {str(e)}", 500
    
    return render_template('upload.html')

@app.errorhandler(413)
def request_entity_too_large(error):
    return "File too large", 413

if __name__ == '__main__':
    app.run(debug=True)