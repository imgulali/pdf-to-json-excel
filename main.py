import PyPDF2
import re
import json
from openpyxl import Workbook
import argparse
import os

def format_name(name):
    return " ".join(word.capitalize() for word in name.strip().split())

def validate_pdf_file(file_path):
    """Validates if the file exists and is a valid PDF."""
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"The file '{file_path}' does not exist.")
    if not file_path.lower().endswith('.pdf'):
        raise ValueError(f"The file '{file_path}' is not a PDF file.")

def main():
    parser = argparse.ArgumentParser(description="Extract student data from a PDF and save to JSON and Excel.")
    parser.add_argument("pdf_file", help="Path to the PDF file to process.")
    parser.add_argument("--json", default="students.json", help="Path to save the JSON file (default: students.json).")
    parser.add_argument("--excel", default="students.xlsx", help="Path to save the Excel file (default: students.xlsx).")
    
    args = parser.parse_args()

    # Validate the PDF file
    try:
        validate_pdf_file(args.pdf_file)
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")
        return

    # Read and extract text from the PDF
    try:
        with open(args.pdf_file, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ''.join(page.extract_text() for page in reader.pages)
    except Exception as e:
        print(f"Error reading the PDF file: {e}")
        return

    # Extract student data using regex
    pattern = r'(F\d{10})\s+([A-Z\s]+)'
    matches = re.findall(pattern, text)

    if not matches:
        print("No student data found in the PDF.")
        return

    students = [{"studentId": match[0], "name": format_name(match[1])} for match in matches]

    # Save data to JSON
    try:
        with open(args.json, 'w') as json_file:
            json.dump(students, json_file, indent=4)
        print(f"Student data saved to JSON: {args.json}")
    except Exception as e:
        print(f"Error saving to JSON file: {e}")
        return

    # Save data to Excel
    try:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Students"
        sheet.append(["Student ID", "Name"])
        for student in students:
            sheet.append([student["studentId"], student["name"]])
        workbook.save(args.excel)
        print(f"Student data saved to Excel: {args.excel}")
    except Exception as e:
        print(f"Error saving to Excel file: {e}")
        return

if __name__ == "__main__":
    main()