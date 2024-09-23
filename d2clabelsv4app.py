import streamlit as st
import pandas as pd
import os
import re
import pdfplumber
from io import BytesIO
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from barcode import EAN13
from barcode.writer import ImageWriter
from datetime import datetime
from zipfile import ZipFile
from PyPDF2 import PdfReader, PdfWriter

# Function to clean up file names
def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', name)

# Function to extract FNSKU from a specific region of the page
def extract_fnsku_from_page(pdf_page):
    # Open the PDF with pdfplumber
    with pdfplumber.open(pdf_page) as pdf:
        first_page = pdf.pages[0]
        
        # Define a bounding box where the FNSKU is located (adjust these values as necessary)
        bbox = (100, 150, 400, 200)  # Adjust coordinates based on where the FNSKU is located on the page
        text = first_page.within_bbox(bbox).extract_text()
        
        fnsku = None
        if text:
            for line in text.split("\n"):
                if re.match(r"^[A-Z0-9]{10}$", line):  # Assuming FNSKU is 10 alphanumeric characters
                    fnsku = line.strip()
                    break
    return fnsku if fnsku else "unknown_fnsku"

# Function to split a PDF into multiple PDFs, one per page, using FNSKU as the file name
def split_fnsku_pdf(uploaded_pdf):
    # Read the uploaded PDF from the in-memory BytesIO object
    pdf_file = BytesIO(uploaded_pdf.read())
    input_pdf = PdfReader(pdf_file)
    total_pages = len(input_pdf.pages)

    # Create an output folder
    output_folder = f"Split_FNSKU_{datetime.now().strftime('%Y%m%d')}"
    os.makedirs(output_folder, exist_ok=True)

    progress_bar = st.progress(0)  # Add progress bar for splitting PDF

    for page_num in range(total_pages):
        writer = PdfWriter()
        writer.add_page(input_pdf.pages[page_num])

        # Extract FNSKU from the page (adjust according to actual PDF content)
        fnsku = extract_fnsku_from_page(pdf_file)  # Extract FNSKU from the page

        # Clean the FNSKU for use as a filename
        fnsku_clean = clean_filename(fnsku)

        # Save each page as a separate PDF using the FNSKU as the filename
        output_filename = os.path.join(output_folder, f"{fnsku_clean}.pdf")
        with open(output_filename, 'wb') as output_pdf:
            writer.write(output_pdf)

        # Update the progress bar
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

# Add option for either generating labels or splitting PDF
option = st.selectbox("Choose an action", ["Generate Labels", "Split FNSKU Labels"])

if option == "Generate Labels":
    st.write("Upload an Excel file with SKU, UPC, and LOT# (if applicable)")

    # Upload the Excel file
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        
        if st.button("Generate Labels"):
            zip_path = generate_pdfs_from_excel(df)
            with open(zip_path, "rb") as f:
                st.download_button("Download ZIP file with Labels", f, file_name=zip_path)

elif option == "Split FNSKU Labels":
    st.write("Upload a PDF file to split FNSKU labels")

    # Upload the PDF file
    uploaded_pdf = st.file_uploader("Upload PDF file", type=["pdf"])

    if uploaded_pdf is not None:
        if st.button("Split PDF"):
            zip_path = split_fnsku_pdf(uploaded_pdf)
            with open(zip_path, "rb") as f:
                st.download_button("Download ZIP file with Split PDFs", f, file_name=zip_path)
