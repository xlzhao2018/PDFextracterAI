# -*- coding: utf-8 -*-
"""
Created on Fri Nov  1 16:49:58 2024

@author: xzhao
"""

import fitz  # PyMuPDF
import re
import openai
import pandas as pd
from PIL import Image
import pytesseract
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Configure Tesseract path if necessary (for Windows)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# OpenAI API key configuration (retrieve from environment variable)
openai.api_key = os.getenv("")

# Define fields with patterns at the very beginning
fields = {
    "unique_id": r"Your current unique identifier",
    "prefix_title": r"Prefix title of the contact",
    "first_name": r"First name of the contact",
    "middle_initial": r"Middle initial of contact",
    "last_name": r"Last name of the contact",
    "suffix_title": r"Suffix title of contact",
    "category": r"(General|Researcher|Lawyer)",
    "address_line_1": r"First line of address",
    "address_line_2": r"Second line of address",
    "address_line_3": r"Third line of address",
    "city": r"City",
    "state": r"State|Province",
    "postal_code": r"Postal|Zip Code",
    "country": r"Country",
    "phone_number": r"Phone number",
    "phone_type": r"Phone type",
    "gender": r"Contact's gender",
    "company": r"Company the contact works for",
    "job_title": r"Contact's job title",
    "department": r"Contact's department",
    "speciality": r"Contact's speciality",
    "citizenship_country": r"Country of citizenship",
    "citizenship_nationality": r"Nationality of citizenship",
    "email": r"Email address",
    "personal_website": r"Personal website",
    "interests": r"Contact Interests",
    "document_path": r"URL or Document Path",
    "comments": r"Additional comments",
    "fax_number": r"Fax number",
    "secondary_email": r"Secondary email address",
    "alternate_phone_number": r"Alternate phone number"
}

# Function to perform OCR on a page
def perform_ocr(page):
    pix = page.get_pixmap(dpi=300)  # Set a higher DPI to improve image quality
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    img = img.convert('L')  # Convert to grayscale to improve OCR accuracy
    text = pytesseract.image_to_string(img, lang='eng')
    return text, img

# Function to extract text from a PDF, with OCR fallback
def extract_text_from_pdf(file_path):
    text = ""
    ocr_text_list = []
    with fitz.open(file_path) as pdf:
        for page_number, page in enumerate(pdf, start=1):
            page_text = page.get_text("text")
            if not page_text.strip():
                print(f"Performing OCR on page {page_number}")
                page_text, img = perform_ocr(page)
                ocr_text_list.append(page_text)
            else:
                print(f"Extracted text from page {page_number} without OCR")
                ocr_text_list.append(page_text)
            text += page_text
    save_ocr_text_as_pdf(ocr_text_list, file_path)
    return text

# Function to save OCR text as a searchable PDF
def save_ocr_text_as_pdf(ocr_text_list, original_file_path):
    output_path = original_file_path.replace(".pdf", "_OCR_Searchable.pdf")
    c = canvas.Canvas(output_path, pagesize=letter)

    for page_number, ocr_text in enumerate(ocr_text_list, start=1):
        c.setFont("Helvetica", 10)
        c.drawString(30, 750, f"Page {page_number}")
        text_lines = ocr_text.split('\n')
        y_position = 720
        for line in text_lines:
            c.drawString(30, y_position, line)
            y_position -= 12
            if y_position < 40:
                c.showPage()
                y_position = 720
                c.setFont("Helvetica", 10)
                c.drawString(30, 750, f"Page {page_number}")

        c.showPage()

    c.save()
    print(f"Searchable OCR output saved to {output_path}")

# Function to refine extraction using ChatGPT
def refine_with_chatgpt(text, field, field_description):
    prompt = (
        f"You are a document extraction assistant. Extract the following information if available: '{field_description}'. "
        f"If the '{field_description}' is not explicitly mentioned in the text below, return 'N/A'. "
        f"\n\nText to analyze:\n\n{text}"
    )
    
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an assistant that extracts specific information from the provided text."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=100,
            temperature=0  # Deterministic output
        )
        return response['choices'][0]['message']['content'].strip()
    
    except Exception as e:
        print(f"Error while calling OpenAI API: {e}")
        return "N/A"

# Function to extract specific fields from text
def extract_fields_from_text(text):
    extracted_data = {}
    for field, pattern in fields.items():
        match = re.search(pattern + r":?\s*(.*)", text, re.IGNORECASE)
        if match:
            group_value = match.group(1)
            if group_value:
                extracted_data[field] = group_value.strip()
            else:
                extracted_data[field] = "N/A"
                print(f"Debug: Match found for field '{field}' but value is None or empty.")
        else:
            # Use ChatGPT to refine the extraction if regex fails
            print(f"Debug: No match found for field '{field}' using pattern '{pattern}', attempting to refine with ChatGPT.")
            extracted_data[field] = refine_with_chatgpt(text, field, pattern)
            print(f"ChatGPT refinement result for '{field}': {extracted_data[field]}")
    return extracted_data

# Main batch processing function
def process_batch_pdfs(pdf_files):
    all_data = []
    for pdf_file in pdf_files:
        text = extract_text_from_pdf(pdf_file)
        extracted_data = extract_fields_from_text(text)
        extracted_data["document_path"] = pdf_file
        all_data.append(extracted_data)

    df = pd.DataFrame(all_data)
    df.to_excel("Extracted_Contacts_with_OCR_and_Verification.xlsx", index=False)
    print("Data has been extracted and saved to Extracted_Contacts_with_OCR_and_Verification.xlsx")

# Example usage
pdf_files = [r"C:\Users\xzhao\NDA_info_extract\V4_OCR\file1.pdf"]
process_batch_pdfs(pdf_files)
