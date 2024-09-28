from flask import Flask, render_template, request, redirect, url_for, flash
import openpyxl
import os
from dotenv import load_dotenv
import random
from flask import send_from_directory
from fpdf import FPDF
from io import BytesIO

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY')
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def check_excel_format(file):
    try:
        workbook = openpyxl.load_workbook(file)
        sheet = workbook.active
        headers = [cell.value for cell in sheet[1]]

        expected_headers = ['Unit', 'Questions', 'Marks', 'Type of Question', 'Probability of the Question coming']
        return headers == expected_headers
    except Exception as e:
        print(f"Error: {e}")
        return False

def get_question_types(file):
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.active
    question_types = {}

    for row in sheet.iter_rows(min_row=2, values_only=True):
        q_type = row[3]
        question_types[q_type] = question_types.get(q_type, 0) + 1

    return question_types

def generate_question_paper(file, request_data):
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.active

    questions = {row[3]: [] for row in sheet.iter_rows(min_row=2, values_only=True)}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        questions[row[3]].append(row)

    selected_questions = []
    for q_type, num_questions in request_data.items():
        num_questions = int(num_questions)
        if num_questions > 0 and q_type in questions:
            selected_questions.extend(random.sample(questions[q_type], min(num_questions, len(questions[q_type]))))

    return selected_questions

@app.route('/')
def index():
    return render_template('index.html')

def create_pdf(questions, output_filename):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    for unit, question, marks, q_type, probability in questions:
        pdf.cell(200, 10, txt=f"Unit: {unit}", ln=True)
        pdf.cell(200, 10, txt=f"Question: {question}", ln=True)
        pdf.cell(200, 10, txt=f"Marks: {marks}", ln=True)
        pdf.cell(200, 10, txt=f"Type: {q_type}", ln=True)
        pdf.cell(200, 10, txt=f"Probability: {probability}", ln=True)
        pdf.cell(200, 10, txt="", ln=True)  # Add a blank line between questions

    pdf.output(output_filename)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)

    file = request.files['file']

    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)

    if file and allowed_file(file.filename):
        # Use BytesIO to read the file directly without saving
        file_stream = BytesIO(file.read())

        if check_excel_format(file_stream):
            question_types = get_question_types(file_stream)
            return render_template('question_selection.html', question_types=question_types)
        else:
            flash('File uploaded but does not match the expected format.')
            return redirect(url_for('index'))

    flash('Invalid file type. Please upload an .xlsx file.')
    return redirect(request.url)

@app.route('/generate', methods=['POST'])
def generate():
    file_path = request.form['file_path']
    selected_questions = []

    for q_type in request.form:
        if q_type.endswith('_count'):
            count = int(request.form[q_type])  # Extract the number of questions
            question_type = q_type[:-6]  # Remove '_count' from the end
            selected_questions.extend(generate_question_paper(file_path, {question_type: count}))

    # Create PDF with the selected questions
    output_filename = os.path.join(UPLOAD_FOLDER, 'question_paper.pdf')
    create_pdf(selected_questions, output_filename)

    # Return the PDF file as an attachment
    return send_from_directory(UPLOAD_FOLDER, 'question_paper.pdf', as_attachment=True)

@app.route('/uploads/<path:filename>')
def serve_uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

if __name__ == '__main__':
    app.run(debug=True)
