import openpyxl
import secrets
import argparse
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Function to create a question paper with unique questions
def create_question_paper(questions, num_questions):
    selected_questions = set()  # To keep track of selected questions
    question_paper = []

    while len(question_paper) < num_questions:
        question = secrets.choice(questions)
        if question not in selected_questions:
            selected_questions.add(question)
            question_paper.append(question)

    return question_paper

# Function to write the question paper to a PDF
def write_to_pdf(question_paper, filename='question_paper.pdf'):
    c = canvas.Canvas(filename, pagesize=letter)
    c.drawString(100, 750, "Question Paper")
    
    for idx, question in enumerate(question_paper, start=1):
        c.drawString(100, 750 - idx * 20, f"{idx}. {question[1]} (Marks: {question[2]}, Type: {question[3]})")

    c.save()

def main(excel_file, num_questions):
    # Load the workbook and select the active worksheet
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active

    # Extract data from the sheet into a list
    data = [row for row in sheet.iter_rows(values_only=True)]
    questions = data[1:]  # Get all questions, ignoring the header

    # Generate the question paper
    question_paper = create_question_paper(questions, num_questions)

    # Write the question paper to a PDF
    write_to_pdf(question_paper)

    print("Question paper created and saved as 'question_paper.pdf'.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate a question paper from an Excel file.")
    parser.add_argument("excel_file", type=str, help="Path to the Excel file (.xlsx)")
    parser.add_argument("--num_questions", type=int, default=5, help="Number of questions to include in the question paper")
    
    args = parser.parse_args()
    
    main(args.excel_file, args.num_questions)
