import os
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import re
import pandas as pd

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'  # Folder where uploaded files will be stored
app.config['ALLOWED_EXTENSIONS'] = {'pdf', 'jpg', 'jpeg'}  # Allowed file extensions

# Function to check if file extension is allowed
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# Route to render the upload form
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle file upload and processing
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"

    files = request.files.getlist('file')  # Handle multiple files

    for file in files:
        if file.filename == '':
            return "No selected file"

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            # Process uploaded file
            if filename.lower().endswith('.pdf'):
                pages = convert_from_path(file_path)
                ocr_text = ""
                for page in pages:
                    ocr_text += pytesseract.image_to_string(page)
            else:
                image = Image.open(file_path)
                ocr_text = pytesseract.image_to_string(image)

            # Extract information using regex patterns
            sender_name = extract_sender_name(ocr_text)
            invoice_id = extract_invoice_id(ocr_text)
            total_amount = extract_total_amount(ocr_text)
            invoice_date = extract_invoice_date(ocr_text)

            # Append data to Excel file
            excel_file = 'invoices.xlsx'
            append_to_excel(excel_file, sender_name, invoice_id, total_amount, invoice_date)

        else:
            return "File type not allowed"

    return "Data successfully appended to Excel file"

# Function to extract sender's name from OCR text
def extract_sender_name(ocr_text):
    lines = ocr_text.split('\n')
    first_line_words = lines[0].split()
    if len(first_line_words) >= 2:
        sender_name = f"{first_line_words[0]} {first_line_words[1]}"
    else:
        sender_name = first_line_words[0]
    return sender_name

# Function to extract invoice ID from OCR text
def extract_invoice_id(ocr_text):
    invoice_pattern = r"Invoice # (\w+-\d+)"
    invoice_match = re.search(invoice_pattern, ocr_text)
    if invoice_match:
        return invoice_match.group(1)
    else:
        return "Invoice ID not found"

# Function to extract total amount from OCR text
def extract_total_amount(ocr_text):
    total_pattern = r"TOTAL \$([\d,]+\.\d{2})"
    total_match = re.search(total_pattern, ocr_text)
    if total_match:
        return total_match.group(1)
    else:
        return "Total Amount not found"

# Function to extract invoice date from OCR text
def extract_invoice_date(ocr_text):
    date_pattern = r"Invoice Date (\d{2}/\d{2}/\d{4})"
    date_match = re.search(date_pattern, ocr_text)
    if date_match:
        return date_match.group(1)
    else:
        return "Invoice Date not found"

# Function to append data to Excel file
def append_to_excel(excel_file, sender_name, invoice_id, total_amount, invoice_date):
    data = {
        'Name of Sender': [sender_name],
        'Invoice ID': [invoice_id],
        'Total Amount': [total_amount],
        'Invoice Date': [invoice_date]
    }
    df = pd.DataFrame(data)
    
    if os.path.exists(excel_file):
        existing_df = pd.read_excel(excel_file)
        updated_df = pd.concat([existing_df, df], ignore_index=True)
        updated_df.to_excel(excel_file, index=False)
    else:
        df.to_excel(excel_file, index=False)

if __name__ == '__main__':
    app.run(debug=True)
