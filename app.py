import os
import sys
import base64
import hashlib
import traceback
from flask import Flask, render_template, request, send_file, redirect, url_for
from werkzeug.utils import secure_filename
import mammoth
import csv
import zipfile
import shutil
from bs4 import BeautifulSoup
import uuid


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.abspath('uploads')
app.config['DATA_FOLDER'] = os.path.abspath('data')
app.config['IMAGES_FOLDER'] = os.path.join(app.config['DATA_FOLDER'], 'images')
app.config['OUTPUT_ZIP'] = 'KnowledgeArticlesImport.zip'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB limit

def debug_print(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)

def create_content_properties():
    content = """CSVEncoding=UTF8
RTAEncoding=UTF8
CSVSeparator=,
#DateFormat=yyyy-MM-dd"""
    with open('content.properties', 'w') as f:
        f.write(content)

def save_image(image_data, extension):
    image_hash = hashlib.md5(image_data).hexdigest()
    image_filename = f"{image_hash}{extension}"
    image_path = os.path.join(app.config['IMAGES_FOLDER'], image_filename)
    
    if not os.path.exists(image_path):
        with open(image_path, "wb") as f:
            f.write(image_data)
    
    return f"images/{image_filename}"

def verify_docx(file_path):
    try:
        with open(file_path, "rb") as docx:
            mammoth.extract_raw_text(docx)
        return True
    except Exception as e:
        debug_print(f"Error verifying DOCX file {file_path}: {str(e)}")
        return False

def convert_docx_to_html(docx_file_path):
    html_file = os.path.splitext(os.path.basename(docx_file_path))[0] + '.html'
    
    try:
        debug_print(f"Starting conversion of {docx_file_path}")
        
        if not os.path.exists(docx_file_path):
            debug_print(f"Error: DOCX file not found: {docx_file_path}")
            return None, [f"Error: DOCX file not found: {docx_file_path}"]
        
        # Ensure the images directory exists
        images_dir = os.path.join(app.config['DATA_FOLDER'], 'images')
        os.makedirs(images_dir, exist_ok=True)
        
        def handle_image(image):
            extension = image.content_type.split("/")[1]
            image_filename = f"{uuid.uuid4()}.{extension}"
            image_path = os.path.join('images', image_filename)
            with image.open() as image_bytes, open(os.path.join(app.config['DATA_FOLDER'], image_path), "wb") as f:
                f.write(image_bytes.read())
            return {"src": image_path}
        
        # Options for mammoth 1.8.0
        options = mammoth.convert_to_html.options(
            style_map=[
                "p[style-name='Heading 1'] => h1:fresh",
                "p[style-name='Heading 2'] => h2:fresh",
                "p[style-name='Heading 3'] => h3:fresh",
                "p[style-name='Heading 4'] => h4:fresh",
                "p[style-name='Heading 5'] => h5:fresh",
                "p[style-name='Heading 6'] => h6:fresh",
                "r[style-name='Strong'] => strong",
                "r[style-name='Emphasis'] => em"
            ],
            ignore_empty_paragraphs=False,
            convert_image=mammoth.images.img_element(handle_image)
        )
        
        with open(docx_file_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file, options=options)
        
        html = result.value
        messages = result.messages
        
        debug_print(f"Mammoth conversion completed for {docx_file_path}")
        
        if not html:
            debug_print(f"Mammoth produced empty HTML for {docx_file_path}")
            return None, [f"Error: Empty HTML produced for {docx_file_path}"]
        
        # Post-processing with BeautifulSoup
        soup = BeautifulSoup(html, 'html.parser')
        
        # Add default styling
        style = soup.new_tag('style')
        style.string = """
            body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
            h1, h2, h3, h4, h5, h6 { margin-top: 1em; margin-bottom: 0.5em; }
            p { margin-bottom: 1em; }
            img { max-width: 100%; height: auto; }
        """
        
        # Ensure proper HTML structure
        if soup.html is None:
            new_html = soup.new_tag('html')
            new_html.append(soup)
            soup = BeautifulSoup(str(new_html), 'html.parser')
        
        if soup.head is None:
            head = soup.new_tag('head')
            soup.html.insert(0, head)
        
        soup.head.append(style)
        
        if soup.body is None:
            body = soup.new_tag('body')
            for tag in soup.html.contents:
                if tag.name not in ['head', 'body']:
                    body.append(tag.extract())
            soup.html.append(body)
        
        # Save the updated HTML
        output_path = os.path.join(app.config['DATA_FOLDER'], html_file)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(str(soup))
        
        debug_print(f"HTML file saved: {output_path}")
        
        return html_file, messages
    except Exception as e:
        debug_print(f"Error converting {docx_file_path} to HTML: {str(e)}")
        debug_print(f"Exception details: {traceback.format_exc()}")
        return None, [f"Error converting {docx_file_path}: {str(e)}"]

def create_csv_record(filename):
    title = os.path.splitext(filename)[0].replace('_', ' ')
    url_name = title.replace(' ', '-')
    return [title, title, url_name, 'application', f"data/{filename}"]

def create_csv_file(records):
    csv_file = 'KnowledgeArticlesImport.csv'
    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Title', 'Summary', 'URLName', 'channels', 'Content__c'])
        writer.writerows(records)
    return csv_file

def create_zip_file(files):
    with zipfile.ZipFile(app.config['OUTPUT_ZIP'], 'w') as zipf:
        for file in files:
            zipf.write(file)
        zipf.write('content.properties')
        for root, _, files in os.walk(app.config['DATA_FOLDER']):
            for file in files:
                zipf.write(os.path.join(root, file))

def clear_files():
    for folder in [app.config['UPLOAD_FOLDER'], app.config['DATA_FOLDER']]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
    if os.path.exists(app.config['OUTPUT_ZIP']):
        os.remove(app.config['OUTPUT_ZIP'])
    if os.path.exists('content.properties'):
        os.remove('content.properties')

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        try:
            debug_print("File upload initiated")
            if 'file' not in request.files:
                debug_print("No file part in the request")
                return redirect(request.url)
            files = request.files.getlist('file')
            if not files or files[0].filename == '':
                debug_print("No selected file")
                return redirect(request.url)
            
            # Clear previous data
            clear_files()
            debug_print("Previous files cleared")
            
            for folder in [app.config['UPLOAD_FOLDER'], app.config['DATA_FOLDER'], app.config['IMAGES_FOLDER']]:
                os.makedirs(folder, exist_ok=True)
                debug_print(f"Folder created/verified: {folder}")
            
            filenames = []
            messages = []
            for file in files:
                if file and file.filename.endswith('.docx'):
                    filename = secure_filename(file.filename)
                    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    file.save(file_path)
                    filenames.append(file_path)  # Store full path
                    debug_print(f"File saved: {file_path}")
                    
                    # Verify file exists after saving
                    if os.path.exists(file_path):
                        debug_print(f"File verified: {file_path}")
                    else:
                        debug_print(f"Error: File not found after saving: {file_path}")
            
            create_content_properties()
            debug_print("Content properties created")
            
            html_files = []
            for file_path in filenames:
                if not os.path.exists(file_path):
                    debug_print(f"Error: File not found before conversion: {file_path}")
                    continue
                
                html_file, msg = convert_docx_to_html(file_path)
                if html_file:
                    html_files.append(html_file)
                    debug_print(f"HTML file generated: {html_file}")
                else:
                    debug_print(f"Failed to generate HTML for {file_path}")
                messages.extend(msg)
            
            if not html_files:
                debug_print("No HTML files were generated")
                raise Exception("No HTML files were generated. Check the console for details.")
            
            csv_records = [create_csv_record(os.path.basename(html_file)) for html_file in html_files]
            csv_file = create_csv_file(csv_records)
            debug_print(f"CSV file created: {csv_file}")
            
            files_to_zip = filenames + [csv_file]
            create_zip_file(files_to_zip)
            debug_print(f"Zip file created: {app.config['OUTPUT_ZIP']}")
            
            return send_file(app.config['OUTPUT_ZIP'], as_attachment=True)
        except Exception as e:
            debug_print(f"An error occurred: {str(e)}")
            debug_print(f"Exception details: {traceback.format_exc()}")
            return f"An error occurred: {str(e)}", 500
    
    return render_template('upload.html')

@app.route('/clear', methods=['POST'])
def clear_files_route():
    clear_files()
    return redirect(url_for('upload_file'))

if __name__ == '__main__':
    app.run(debug=True)