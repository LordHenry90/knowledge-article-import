from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import mammoth
from docx import Document
import shutil
import csv
import zipfile
import re

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

    with open(file_path, "rb") as f:
        style_map = """
            p[style-name='Heading 1'] => h1:fresh,
            p[style-name='Heading 2'] => h2:fresh,
            p[style-name='Heading 3'] => h3:fresh,
            p[style-name='Heading 4'] => h4:fresh,
            p[style-name='Normal'] => p:fresh,
            p[style-name='List Paragraph'] => ul > li:fresh,
            p[style-name='Quote'] => blockquote:fresh
        """
        result = mammoth.convert_to_html(f, style_map=style_map)
        html_content = result.value

    # Extract images from the document
    for i, rel in enumerate(doc.part.rels.values()):
        if "image" in rel.reltype:
            image_data = rel.target_part.blob
            image_filename = f"{filename.replace('.docx', '')}_{i}.png"
            image_path = os.path.join(app.config['IMAGES_FOLDER'], image_filename)
            with open(image_path, "wb") as img_file:
                img_file.write(image_data)
            image_mapping[f"image_{i}"] = f"data/images/{image_filename}"

    # Replace base64 images with actual image paths in the HTML content
    def replace_base64_images(html):
        img_tags = re.findall(r'<img [^>]*src="data:image/.*?;base64,.*?"[^>]*>', html)
        for i, img_tag in enumerate(img_tags):
            new_img_tag = img_tag
            if i in image_mapping:
                new_img_tag = re.sub(r'src="data:image/.*?;base64,.*?"', f'src="{image_mapping[f'image_{i}']}"', img_tag)
            html = html.replace(img_tag, new_img_tag)
        return html

    html_content = replace_base64_images(html_content)

    with open(html_path, 'w', encoding='utf-8') as html_file:
        html_file.write(html_content)

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