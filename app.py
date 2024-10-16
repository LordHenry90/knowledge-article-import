from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import mammoth
from docx import Document
import shutil
import csv
import zipfile
import re
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.shared import Inches

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
DATA_FOLDER = 'data'
IMAGES_FOLDER = os.path.join(DATA_FOLDER, 'images')

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DATA_FOLDER'] = DATA_FOLDER
app.config['IMAGES_FOLDER'] = IMAGES_FOLDER

# Ensure folders exist
for folder in [UPLOAD_FOLDER, DATA_FOLDER, IMAGES_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/', methods=['POST'])
def upload_files():
    for folder in [UPLOAD_FOLDER, DATA_FOLDER, IMAGES_FOLDER]:
        if not os.path.exists(folder):
            os.makedirs(folder)
    if 'file' not in request.files:
        return 'No files part'
    files = request.files.getlist('file')
    for file in files:
        if file.filename.endswith('.docx'):
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))
    return redirect(url_for('process_files'))

@app.route('/process', methods=['GET'])
def process_files():
    for folder in [UPLOAD_FOLDER, DATA_FOLDER, IMAGES_FOLDER]:
        if not os.path.exists(folder):
            os.makedirs(folder)
    content_files = []
    for filename in os.listdir(app.config['UPLOAD_FOLDER']):
        if filename.endswith('.docx'):
            process_docx(filename)
            content_files.append(filename.replace('.docx', '.html'))

    create_content_properties()
    create_csv(content_files)
    create_zip_file()

    return redirect(url_for('download_zip'))

def process_docx(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    doc = Document(file_path)
    
    html_path = os.path.join(app.config['DATA_FOLDER'], filename.replace('.docx', '.html'))
    image_mapping = {}  # Dictionary to map base64 images to file paths
    html_content = ""
    image_index = 0

    # Iterate over paragraphs and tables to maintain the order of text and images
    for block in iter_block_items(doc):
        if isinstance(block, str):
            html_content += f"<p>{block}</p>"
        elif isinstance(block, tuple):  # Image block
            image_data = block[1]
            image_filename = f"{filename.replace('.docx', '')}_{image_index}.png"
            image_path = os.path.join(app.config['IMAGES_FOLDER'], image_filename)
            with open(image_path, "wb") as img_file:
                img_file.write(image_data)
            image_mapping[f"image_{image_index}"] = f"../data/images/{image_filename}"
            html_content += f'<p><img src="{image_mapping[f"image_{image_index}"]}" /></p>'
            image_index += 1

    with open(html_path, 'w', encoding='utf-8') as html_file:
        html_file.write(html_content)

def iter_block_items(doc):
    """
    Generate a sequential stream of paragraphs and images in the order they appear in the document.
    """
    for block in doc.element.body:
        if block.tag == qn('w:p'):
            yield ''.join([node.text for node in block.iter(qn('w:t')) if node.text])
        elif block.tag == qn('w:drawing'):
            for pic in block.iter(qn('a:blip')):
                r_id = pic.get(qn('r:embed'))
                if r_id:
                    rel = doc.part.rels[r_id]
                    yield ('image', rel.target_part.blob)

def create_content_properties():
    content_properties_path = os.path.join(app.config['DATA_FOLDER'], 'content.properties')
    with open(content_properties_path, 'w') as properties_file:
        properties_file.write("CSVEncoding=UTF8\n")
        properties_file.write("RTAEncoding=UTF8\n")
        properties_file.write("CSVSeparator=,\n")
        properties_file.write("#DateFormat=yyyy-MM-dd\n")

def create_csv(content_files):
    csv_path = os.path.join(app.config['DATA_FOLDER'], 'KnowledgeArticlesImport.csv')
    with open(csv_path, mode='w', newline='') as csv_file:
        writer = csv.writer(csv_file, delimiter=',')
        writer.writerow(['Title', 'Summary', 'URLName', 'Channels', 'Content__c'])
        for content_file in content_files:
            title = content_file.replace('.html', '').replace('_', ' ')
            summary = title
            urlname = title.replace(' ', '-')
            channels = 'application'
            content = f"data/{content_file}"
            writer.writerow([title, summary, urlname, channels, content])

def create_zip_file():
    zip_path = os.path.join(app.config['DATA_FOLDER'], 'KnowledgeArticlesImport.zip')
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        # Add content.properties and KnowledgeArticlesImport.csv to the ZIP root
        content_properties_path = os.path.join(app.config['DATA_FOLDER'], 'content.properties')
        csv_path = os.path.join(app.config['DATA_FOLDER'], 'KnowledgeArticlesImport.csv')
        zipf.write(content_properties_path, 'content.properties')
        zipf.write(csv_path, 'KnowledgeArticlesImport.csv')
        
        # Add HTML files to the data/ folder in the ZIP
        html_folder = os.path.join(app.config['DATA_FOLDER'])
        for filename in os.listdir(html_folder):
            if filename.endswith('.html'):
                file_path = os.path.join(html_folder, filename)
                zipf.write(file_path, os.path.join('data', filename))
        
        # Add images to the data/images/ folder in the ZIP
        images_folder = app.config['IMAGES_FOLDER']
        for filename in os.listdir(images_folder):
            if filename.endswith('.png'):
                file_path = os.path.join(images_folder, filename)
                zipf.write(file_path, os.path.join('data', 'images', filename))

@app.route('/download', methods=['GET'])
def download_zip():
    zip_path = os.path.join(app.config['DATA_FOLDER'], 'KnowledgeArticlesImport.zip')
    return send_file(zip_path, as_attachment=True)

@app.route('/clear', methods=['POST'])
def clear_files():
    for folder in [UPLOAD_FOLDER, DATA_FOLDER, IMAGES_FOLDER]:
        if not os.path.exists(folder):
            os.makedirs(folder)
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)