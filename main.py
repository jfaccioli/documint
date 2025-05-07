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
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB limit

ALLOWED_WORD_EXTENSIONS = {'.docx'}
ALLOWED_EXCEL_EXTENSIONS = {'.xlsx', '.xls'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

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

    data.columns = [col.replace(" ", "_") for col in data.columns]  # Normalise headers
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

    data.columns = [col.replace(" ", "_") for col in data.columns]
    filenames = []

    for index, row in data.iterrows():
        doc = Document(word_filepath)
        _replace_placeholders(doc, row.to_dict())

        safe_filename = secure_filename(str(row[chosen_column])) or f"document_{index}"
        output_filename = f"{safe_filename}_{index}.docx"
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

def _replace_placeholders(doc, row_data):
    for paragraph in doc.paragraphs:
        _replace_placeholders_in_runs(paragraph.runs, row_data)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_placeholders_in_runs(paragraph.runs, row_data)

def _replace_placeholders_in_runs(runs, row_data):
    for run in runs:
        for placeholder, value in row_data.items():
            if isinstance(value, datetime.datetime) and pd.notna(value):
                value = value.strftime('%d/%m/%Y')
            elif pd.isna(value):
                value = ""
            formatted_placeholder = placeholder.replace(" ", "_")
            tag = f"«{formatted_placeholder}»"
            if tag in run.text:
                run.text = run.text.replace(tag, str(value))

if __name__ == '__main__':
    app.run(debug=True)
