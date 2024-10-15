from flask import Flask, request, send_file, render_template, url_for, redirect
from werkzeug.utils import secure_filename
import os
import zipfile
import shutil
from PIL import Image
import csv
import re
import logging
from bs4 import BeautifulSoup
from docx import Document, InlineShape

# Initialize Flask app
app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.DEBUG)

# Upload folder
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Home route
@app.route('/')
def index():
    return render_template('index.html', design_system=True)

# Upload route
@app.route('/upload', methods=['POST'])
def upload():
    try:
        if 'files[]' not in request.files:
            return 'No file part', 400
        files = request.files.getlist('files[]')
        if not files:
            return 'No selected file', 400

        output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'Root')
        os.makedirs(output_path, exist_ok=True)

        for file in files:
            if file.filename == '':
                continue
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            process_docx(file_path, output_path)

        # Zip output
        zip_path = os.path.join(OUTPUT_FOLDER, 'KnowledgeArticlesImport.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            # Add content.properties and CSV file
            zipf.write(os.path.join(output_path, 'content.properties'), 'content.properties')
            zipf.write(os.path.join(output_path, 'KnowledgeArticlesImport.csv'), 'KnowledgeArticlesImport.csv')
            # Add files from data folder
            data_path = os.path.join(output_path, 'data')
            for root, _, files in os.walk(data_path):
                for file in files:
                    zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), output_path))

        return render_template('download.html', zip_url=url_for('download_zip'))
    except Exception as e:
        logging.error(f"Error occurred during upload: {e}")
        return f"An error occurred: {e}", 500

# Download route
@app.route('/download')
def download_zip():
    try:
        zip_path = os.path.join(OUTPUT_FOLDER, 'KnowledgeArticlesImport.zip')
        return send_file(zip_path, as_attachment=True)
    except Exception as e:
        logging.error(f"Error occurred during download: {e}")
        return f"An error occurred: {e}", 500

# Delete uploaded files and generated output
@app.route('/delete')
def delete_files():
    try:
        # Delete uploaded files
        for root, dirs, files in os.walk(UPLOAD_FOLDER):
            for file in files:
                os.remove(os.path.join(root, file))

        # Delete generated output
        shutil.rmtree(OUTPUT_FOLDER)
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        return redirect(url_for('index'))
    except Exception as e:
        logging.error(f"Error occurred during file deletion: {e}")
        return f"An error occurred: {e}", 500

# Process DOCX file
def process_docx(file_path, output_path):
    try:
        root_data_path = os.path.join(output_path, 'data')
        images_path = os.path.join(root_data_path, 'images')
        os.makedirs(root_data_path, exist_ok=True)
        os.makedirs(images_path, exist_ok=True)

        # Extract content from DOCX using python-docx without converting images to base64
        doc = Document(file_path)
        html_content = ""
        img_counter = 0

        for block in iter_block_items(doc):
            if isinstance(block, InlineShape) and block.type == 3:  # Inline image
                img_counter += 1
                img_data = block._inline.graphic.graphicData.pic.blipFill.blip.blob
                img_filename = f'image_{img_counter}.png'
                img_path = os.path.join(images_path, img_filename)
                with open(img_path, 'wb') as img_file:
                    img_file.write(img_data)
                html_content += f'<p><img src="images/{img_filename}" alt="Image {img_counter}" /></p>'
            elif hasattr(block, 'text'):
                html_content += f"<p>{block.text}</p>"

        # Create HTML file
        html_content = f'<html><head><meta charset="utf-8"></head><body>{html_content}</body></html>'
        html_filename = secure_filename(os.path.splitext(os.path.basename(file_path))[0]) + '.html'
        html_path = os.path.join(root_data_path, html_filename)
        with open(html_path, 'w', encoding='utf-8') as html_file:
            html_file.write(html_content)

        # Create content.properties
        properties_path = os.path.join(output_path, 'content.properties')
        with open(properties_path, 'w') as prop_file:
            prop_file.write("CSVEncoding=UTF8\n")
            prop_file.write("RTAEncoding=UTF8\n")
            prop_file.write("CSVSeparator=,\n")
            prop_file.write("#DateFormat=yyyy-MM-dd\n")

        # Create CSV file
        csv_path = os.path.join(output_path, 'KnowledgeArticlesImport.csv')
        csv_exists = os.path.exists(csv_path)
        with open(csv_path, 'a', newline='', encoding='utf-8') as csv_file:
            csv_writer = csv.writer(csv_file)
            if not csv_exists:
                csv_writer.writerow(['Title', 'Summary', 'URLName', 'channels', 'Content__c'])
            title = os.path.splitext(os.path.basename(file_path))[0].replace('_', ' ')
            url_name = re.sub(r'[\s_]+', '-', os.path.splitext(os.path.basename(file_path))[0])
            csv_writer.writerow([title, title, url_name, 'application', f'data/{html_filename}'])
    except Exception as e:
        logging.error(f"Error occurred while processing DOCX file: {e}")
        raise

# Helper function to iterate over paragraphs and images
def iter_block_items(parent):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph
    from docx.shape import InlineShape

    if isinstance(parent, Document):
        parent_elm = parent.element.body
    else:
        raise ValueError("Unsupported parent for iter_block_items")

    for child in parent_elm:
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            for row in child:
                for cell in row:
                    for sub_element in iter_block_items(cell):
                        yield sub_element
        elif child.tag == qn('w:drawing'):
            yield InlineShape(OxmlElement(child), parent)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))