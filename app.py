from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, Request, Response
import openpyxl
import os
from dotenv import load_dotenv
import random
from fpdf import FPDF
from io import BytesIO
import tempfile

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY')

ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def check_excel_format(file):
    try:
        workbook = openpyxl.load_workbook(file)
        sheet = workbook.active

        # Filter out empty or None headers
        headers = [str(cell.value).strip().lower() for cell in sheet[1] if cell.value is not None and cell.value.strip()]

        expected_headers = ['unit', 'questions', 'marks', 'type of question', 'probability']
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

def create_pdf(questions):
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

    # Create a temporary file to hold the PDF data
    with tempfile.NamedTemporaryFile(delete=True) as temp_pdf:
        pdf.output(temp_pdf.name)  # Save to the temporary file
        temp_pdf.seek(0)  # Rewind the file pointer to the beginning
        pdf_output = temp_pdf.read()  # Read the content of the temporary file into memory

    return pdf_output


@app.route('/')
def index():
    return render_template('index.html')

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

    # Create PDF in memory with the selected questions
    pdf_output = create_pdf(selected_questions)

    # Return the PDF file as an attachment
    return Response(pdf_output, mimetype='application/pdf', headers={"Content-Disposition": "attachment;filename=question_paper.pdf"})


if __name__ == '__main__':
    app.run(debug=True)
