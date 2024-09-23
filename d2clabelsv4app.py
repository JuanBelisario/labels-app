import streamlit as st
import os
import re
import pdfplumber
from io import BytesIO
from datetime import datetime
from zipfile import ZipFile
from PyPDF2 import PdfReader, PdfWriter

# Function to clean up file names
def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', name)

# Function to extract FNSKU from a specific region of the page using pdfplumber
def extract_fnsku_from_page(page):
    # Define a bounding box where the FNSKU is located (adjust these values as necessary)
    bbox = (100, 150, 400, 200)  # Adjust coordinates based on where the FNSKU is located on the page
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
    # Read the uploaded PDF from the in-memory BytesIO object
    pdf_file = BytesIO(uploaded_pdf.read())  # Convert uploaded file to BytesIO
    input_pdf = PdfReader(pdf_file)
    total_pages = len(input_pdf.pages)

    # Create an output folder
    output_folder = f"Split_FNSKU_{datetime.now().strftime('%Y%m%d')}"
    os.makedirs(output_folder, exist_ok=True)

    progress_bar = st.progress(0)  # Add progress bar for splitting PDF

    for page_num in range(total_pages):
        writer = PdfWriter()
        writer.add_page(input_pdf.pages[page_num])

        # Use pdfplumber to extract the FNSKU from the current page
        with pdfplumber.open(pdf_file) as pdf:
            page = pdf.pages[page_num]
            fnsku = extract_fnsku_from_page(page)  # Extract FNSKU from the page

        # Clean the FNSKU for use as a filename
        fnsku_clean = clean_filename(fnsku)

        # Save each page as a separate PDF using the FNSKU as the filename
        output_filename = os.path.join(output_folder, f"{fnsku_clean}_page_{page_num + 1}.pdf")
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

if option == "Split FNSKU Labels":
    st.write("Upload a PDF file to split FNSKU labels")

    # Upload the PDF file
    uploaded_pdf = st.file_uploader("Upload PDF file", type=["pdf"])

    if uploaded_pdf is not None:
        if st.button("Split PDF"):
            zip_path = split_fnsku_pdf(uploaded_pdf)  # Call the function to split the PDF
            if zip_path:
                with open(zip_path, "rb") as f:
                    st.download_button("Download ZIP file with Split PDFs", f, file_name=zip_path)
