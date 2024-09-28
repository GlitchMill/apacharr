# Acharr

This Python application generates a question paper from an Excel file containing questions and their details. The questions are randomly selected based on a unique selection process, and the final question paper is saved as a PDF.

## Features

- Load questions from an Excel `.xlsx` file.
- Generate a specified number of unique questions for the question paper.
- Save the generated question paper as a PDF.

## Prerequisites

Before running the application, ensure you have Python 3 installed along with the necessary libraries. You can install the required libraries using pip:

pip install openpyxl reportlab

## Usage

To run the application, use the following command in your terminal:

python app.py path/to/your_file.xlsx --num_questions <number_of_questions>

- `path/to/your_file.xlsx`: The path to the Excel file containing the questions.
- `--num_questions <number_of_questions>`: (Optional) The number of questions to include in the question paper (default is 5).

### Example

python your_script.py questions.xlsx --num_questions 10

This command will generate a question paper with 10 questions from the `questions.xlsx` file.

## Excel File Format

The Excel file should contain the following columns:

| Unit        | Questions                               | Marks | Type of Question | Probability |
| ----------- | --------------------------------------- | ----- | ---------------- | ----------- |
| Mathematics | What is the derivative of x^2?          | 5     | Short Answer     | 0.2         |
| Physics     | Explain Newton's second law of motion.  | 10    | Essay            | 0.15        |
| Chemistry   | What is the chemical formula for water? | 3     | MCQ              | 0.25        |
| Biology     | Describe the process of photosynthesis. | 8     | Short Answer     | 0.1         |
| ...         | ...                                     | ...   | ...              | ...         |



## Output

The generated question paper will be saved as `question_paper.pdf` in the same directory as the script. You can customize the output filename by modifying the `write_to_pdf` function in the code.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- OpenPyXL for reading Excel files.
- ReportLab for generating PDF documents.

You can copy and paste this into a `README.md` file! Let me know if you need any adjustments.
