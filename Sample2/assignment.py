import os
import re
import pandas as pd
from PyPDF2 import PdfReader
from docx import Document

def extract_text_from_pdf(file_path):
    try:
        with open(file_path, 'rb') as f:
            reader = PdfReader(f)
            text = ''
            for page in reader.pages:
                text += page.extract_text()
        return text
    except Exception as e:
        print(f"Error extracting text from {file_path}: {e}")
        return None

def extract_text_from_docx(file_path):
    try:
        doc = Document(file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text
        return text
    except Exception as e:
        print(f"Error extracting text from {file_path}: {e}")
        return None

def extract_text(file_path):
    try:
        _, file_extension = os.path.splitext(file_path)
        if file_extension.lower() == '.pdf':
            return extract_text_from_pdf(file_path)
        elif file_extension.lower() == '.docx':
            return extract_text_from_docx(file_path)
        else:
            print(f"Unsupported file format: {file_extension}")
            return None
    except Exception as e:
        print(f"Error extracting text from {file_path}: {e}")
        return None

def extract_email(text):
    email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    emails = re.findall(email_regex, text)
    return emails

def extract_contact_number(text):
    contact_regex = r'\b(?:\d[ -]?){9,15}\b'
    contacts = re.findall(contact_regex, text)
    return contacts

def process_cv_files(folder_path):
    data = []
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        text = extract_text(file_path)
        if text:
            emails = extract_email(text)
            contacts = extract_contact_number(text)
            data.append({"File Name": filename, "Email": emails, "Contact Number": contacts, "Text": text})
    return data

def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False, engine='openpyxl')

if __name__ == "__main__":
    folder_path = input("Enter the folder path containing CVs: ").strip('"')
    output_file = input("Enter the output Excel file name (ending with .xls): ").strip('"')

    cv_data = process_cv_files(folder_path)
    save_to_excel(cv_data, output_file)
    print(f"Data saved to {output_file}")
