from flask import Flask, request, send_file, render_template, url_for
from werkzeug.utils import secure_filename
import os
import zipfile
import mammoth
import shutil
from PIL import Image
import csv
import re
import logging
from bs4 import BeautifulSoup

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
        if 'file' not in request.files:
            return 'No file part', 400
        file = request.files['file']
        if file.filename == '':
            return 'No selected file', 400

        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        # Process .docx file
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'Root')
        os.makedirs(output_path, exist_ok=True)
        process_docx(file_path, output_path)

        # Zip output
        zip_path = os.path.join(OUTPUT_FOLDER, 'KnowledgeArticlesImport.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            # Add content.properties and CSV file
            zipf.write(os.path.join(output_path, 'content.properties'), 'content.properties')
            zipf.write(os.path.join(output_path, 'KnowledgeArticlesImport.csv'), 'KnowledgeArticlesImport.csv')
            # Add files from Data folder
            data_path = os.path.join(output_path, 'Data')
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

# Process DOCX file
def process_docx(file_path, output_path):
    try:
        root_data_path = os.path.join(output_path, 'Data')
        os.makedirs(root_data_path, exist_ok=True)

        # Extract content from DOCX using mammoth
        with open(file_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html_content = result.value  # The generated HTML

        # Post-process HTML to fix list numbering using BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        ordered_lists = soup.find_all('ol')
        for ol in ordered_lists:
            list_items = ol.find_all('li')
            for index, li in enumerate(list_items, start=1):
                li['value'] = str(index)

        # Convert the modified soup back to HTML
        html_content = str(soup)

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
        with open(csv_path, 'w', newline='', encoding='utf-8') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow(['Title', 'Summary', 'URLName', 'channels', 'QUESTION__c', 'ANSWER__c', 'RichTextDescription__c'])
            title = os.path.splitext(os.path.basename(file_path))[0]
            url_name = re.sub(r'[\s_]+', '-', title)
            csv_writer.writerow([title, title, url_name, 'application', '', '', f'data/{html_filename}'])
    except Exception as e:
        logging.error(f"Error occurred while processing DOCX file: {e}")
        raise

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
    