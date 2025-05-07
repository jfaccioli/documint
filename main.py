from flask import Flask, request, render_template, send_from_directory, abort, after_this_request
from werkzeug.utils import secure_filename
import pandas as pd
from docx import Document
import os
import re
import zipfile
import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB upload limit

# Create folders if they don't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_WORD_EXTENSIONS = {'.docx'}
ALLOWED_EXCEL_EXTENSIONS = {'.xlsx', '.xls'}

def allowed_file(filename, allowed_extensions):
    return '.' in filename and os.path.splitext(filename)[1].lower() in allowed_extensions

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    word_file = request.files.get('wordfile')
    excel_file = request.files.get('excelfile')

    if not word_file or not allowed_file(word_file.filename, ALLOWED_WORD_EXTENSIONS):
        abort(400, 'Invalid or missing Word (.docx) file.')

    if not excel_file or not allowed_file(excel_file.filename, ALLOWED_EXCEL_EXTENSIONS):
        abort(400, 'Invalid or missing Excel (.xlsx or .xls) file.')

    word_filename = secure_filename(word_file.filename)
    excel_filename = secure_filename(excel_file.filename)

    word_filepath = os.path.join(app.config['UPLOAD_FOLDER'], word_filename)
    excel_filepath = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)

    word_file.save(word_filepath)
    excel_file.save(excel_filepath)

    try:
        data = pd.read_excel(excel_filepath)
    except Exception:
        abort(400, 'Failed to read the Excel file. Ensure it is a valid format.')

    columns = data.columns.tolist()
    return render_template('choose_column.html', columns=columns, word_filepath=word_filepath, excel_filepath=excel_filepath)

@app.route('/process', methods=['POST'])
def process_files():
    word_filepath = request.form['word_filepath']
    excel_filepath = request.form['excel_filepath']
    chosen_column = request.form['chosen_column']

    try:
        data = pd.read_excel(excel_filepath, engine='openpyxl')
    except Exception:
        abort(400, 'Error reading the Excel file during processing.')

    # Log Excel columns for debugging
    print(f"Excel columns: {data.columns.tolist()}")

    filenames = []

    for index, row in data.iterrows():
        doc = Document(word_filepath)
        row_dict = row.to_dict()

        # Log placeholders found in document for debugging
        placeholders_found = set()
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                matches = re.findall(r"«[^»]+»", run.text)
                placeholders_found.update(matches)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            matches = re.findall(r"«[^»]+»", run.text)
                            placeholders_found.update(matches)
        print(f"Placeholders found in document: {placeholders_found}")

        # Replace placeholders in paragraphs
        for paragraph in doc.paragraphs:
            _replace_placeholders_in_paragraph(paragraph, row_dict)

        # Replace placeholders in tables
        for table in doc.tables:
            _replace_placeholders_in_table(table, row_dict)

        output_filename = f"{secure_filename(str(row[chosen_column]))}_{index}.docx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        doc.save(output_path)
        filenames.append(output_path)

    zip_path = os.path.join(app.config['OUTPUT_FOLDER'], "processed_documents.zip")

    with zipfile.ZipFile(zip_path, 'w') as doc_zip:
        for file in filenames:
            doc_zip.write(file, arcname=os.path.basename(file))

    @after_this_request
    def remove_files(response):
        for file in filenames:
            if os.path.exists(file):
                os.remove(file)
        for file in [word_filepath, excel_filepath]:
            if os.path.exists(file):
                os.remove(file)
        if os.path.exists(zip_path):
            os.remove(zip_path)
        return response

    return send_from_directory(app.config['OUTPUT_FOLDER'], "processed_documents.zip", as_attachment=True)

def _replace_placeholders_in_paragraph(paragraph, row_data):
    # Combine all runs' text into a single string
    full_text = ''.join(run.text for run in paragraph.runs)
    replaced = False

    # Perform replacements on the full text
    for placeholder, value in row_data.items():
        if isinstance(value, datetime.datetime) and pd.notna(value):
            value = value.strftime('%d/%m/%Y')
        elif pd.isna(value):
            value = ""
        formatted_placeholder = placeholder.replace(' ', '_')
        pattern = re.compile(r"«" + re.escape(formatted_placeholder) + r"»", re.IGNORECASE)
        if pattern.search(full_text):
            full_text = pattern.sub(str(value), full_text)
            replaced = True

    # If replacements were made, update the paragraph
    if replaced:
        # Preserve original formatting by clearing runs and adding new run
        for run in paragraph.runs:
            run.text = ''
        paragraph.add_run(full_text)

def _replace_placeholders_in_table(table, row_data):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                # Log table cell text for debugging
                cell_text = ''.join(run.text for run in paragraph.runs)
                print(f"Table cell text: {cell_text}")
                _replace_placeholders_in_paragraph(paragraph, row_data)
            # Handle nested tables
            for nested_table in cell.tables:
                _replace_placeholders_in_table(nested_table, row_data)

if __name__ == '__main__':
    app.run(debug=True)
