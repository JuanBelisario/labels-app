import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from barcode import EAN13
from barcode.writer import ImageWriter
from datetime import datetime
from zipfile import ZipFile
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber

# Function to clean up file names
def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', name)

# Function to extract FNSKU from a specific region of the page using pdfplumber
def extract_fnsku_from_page(page):
    bbox = (59.46, 43.07, 102.71600000000002, 51.07)  # Coordinates based on your input
    text = page.within_bbox(bbox).extract_text()
    
    fnsku = None
    if text:
        for line in text.split("\n"):
            if re.match(r"^[A-Z0-9]{10}$", line):  # Assuming FNSKU is 10 alphanumeric characters
                fnsku = line.strip()
                break
    return fnsku if fnsku else "unknown_fnsku"

# Function to split a PDF into multiple PDFs, one per page, using FNSKU as the file name
def split_fnsku_pdf(uploaded_pdf):
    # Reset the file pointer and read the PDF once
    pdf_file = BytesIO(uploaded_pdf.read())  # Convert uploaded file to BytesIO
    input_pdf = PdfReader(pdf_file)
    total_pages = len(input_pdf.pages)

    # Create an output folder
    output_folder = f"Split_FNSKU_{datetime.now().strftime('%Y%m%d')}"
    os.makedirs(output_folder, exist_ok=True)

    progress_bar = st.progress(0)

    # Reset the pdf_file pointer again to use with pdfplumber
    pdf_file.seek(0)
    
    # Use pdfplumber to extract text from the entire PDF, only opening it once
    with pdfplumber.open(pdf_file) as pdf:
        for page_num in range(total_pages):
            writer = PdfWriter()
            writer.add_page(input_pdf.pages[page_num])

            page = pdf.pages[page_num]
            fnsku = extract_fnsku_from_page(page)  # Extract FNSKU from the page

            fnsku_clean = clean_filename(fnsku)
            output_filename = os.path.join(output_folder, f"{fnsku_clean}_page_{page_num + 1}.pdf")
            with open(output_filename, 'wb') as output_pdf:
                writer.write(output_pdf)

            progress_bar.progress((page_num + 1) / total_pages)

    # Compress the split PDFs into a ZIP file
    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for folder_name, subfolders, filenames in os.walk(output_folder):
            for filename in filenames:
                filepath = os.path.join(folder_name, filename)
                zipObj.write(filepath, os.path.basename(filepath))

    return zip_filename

# Streamlit interface
st.title("Label Tools")

option = st.selectbox("Choose an action", ["Generate Labels", "Split FNSKU Labels"])

if option == "Generate Labels":
    st.write("Upload an Excel file with SKU, UPC, and LOT# (if applicable)")
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            if st.button("Generate Labels"):
                zip_path = generate_pdfs_from_excel(df)
                if zip_path:
                    with open(zip_path, "rb") as f:
                        st.download_button("Download ZIP file with Labels", f, file_name=zip_path)
        except Exception as e:
            st.error(f"Error reading the Excel file: {e}")

elif option == "Split FNSKU Labels":
    st.write("Upload a PDF file to split FNSKU labels")
    uploaded_pdf = st.file_uploader("Upload PDF file", type=["pdf"])

    if uploaded_pdf is not None:
        if st.button("Split PDF"):
            try:
                zip_path = split_fnsku_pdf(uploaded_pdf)
                if zip_path:
                    with open(zip_path, "rb") as f:
                        st.download_button("Download ZIP file with Split PDFs", f, file_name=zip_path)
            except Exception as e:
                st.error(f"Error processing the PDF file: {e}")
