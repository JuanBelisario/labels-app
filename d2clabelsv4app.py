import streamlit as st
import pandas as pd
import os
import re
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

# Function to generate UPC labels in PDF format
def generate_label_pdf(sku, upc_code, lot_num, output_path):
    width, height = 60 * mm, 35 * mm
    c = canvas.Canvas(output_path, pagesize=(width, height))

    x_margin = 4.5 * mm
    y_sku = height - 7.75 * mm
    y_barcode = height / 2 - 8 * mm
    y_lot = 4.75 * mm
    barcode_width = 51.5 * mm

    c.setFont("Helvetica", 9.5)
    c.drawCentredString(width / 2, y_sku, sku)

    if len(upc_code) == 12:
        upc_code = '0' + upc_code

    barcode_filename = clean_filename(f"{sku}_barcode")
    barcode_path = f"{barcode_filename}.png"

    options = {
        'module_width': 0.35,
        'module_height': 16,
        'font_size': 7.75,
        'text_distance': 4.5,
        'quiet_zone': 1.25,
        'dpi': 600
    }

    barcode_ean = EAN13(upc_code, writer=ImageWriter())
    barcode_ean.save(barcode_filename, options)

    c.drawImage(barcode_path, (width - barcode_width) / 2, y_barcode, width=barcode_width, height=16 * mm)
    os.remove(barcode_path)

    c.setFont("Helvetica", 9)
    if lot_num:
        lot_box_width = 40 * mm
        lot_box_height = 4 * mm
        x_lot_box = (width - lot_box_width) / 2
        y_lot_box = y_lot - 1.125 * mm
        c.setStrokeColorRGB(0, 0, 0)
        c.rect(x_lot_box, y_lot_box, lot_box_width, lot_box_height, stroke=1, fill=0)
        c.drawCentredString(width / 2, y_lot, lot_num)

    c.save()

# Function to generate PDFs and compress them into a ZIP file
def generate_pdfs_from_excel(df):
    first_sku = df.iloc[0]['SKU']
    current_date = datetime.now().strftime("%Y%m%d")

    output_folder = f"{first_sku}_{current_date}"
    os.makedirs(output_folder, exist_ok=True)

    total_rows = len(df)
    progress_bar = st.progress(0)  # Add progress bar

    for index, row in df.iterrows():
        sku = row['SKU']
        upc_code = str(row['UPC Code']).zfill(12)
        lot_num = row['LOT#'] if pd.notnull(row['LOT#']) else ""
        pdf_filename = clean_filename(f"{sku}.pdf")
        pdf_path = os.path.join(output_folder, pdf_filename)
        generate_label_pdf(sku, upc_code, lot_num, pdf_path)

        # Update the progress bar
        progress_bar.progress((index + 1) / total_rows)

    # Compress all PDFs into a ZIP file
    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for folder_name, subfolders, filenames in os.walk(output_folder):
            for filename in filenames:
                filepath = os.path.join(folder_name, filename)
                zipObj.write(filepath, os.path.basename(filepath))

    return zip_filename

# Function to split a PDF into multiple PDFs, one per page
def split_fnsku_pdf(uploaded_pdf):
    input_pdf = PdfReader(uploaded_pdf)
    total_pages = len(input_pdf.pages)

    # Create an output folder
    output_folder = f"Split_FNSKU_{datetime.now().strftime('%Y%m%d')}"
    os.makedirs(output_folder, exist_ok=True)

    progress_bar = st.progress(0)  # Add progress bar for splitting PDF

    for page_num in range(total_pages):
        writer = PdfWriter()
        writer.add_page(input_pdf.pages[page_num])

        # Save each page as a separate PDF
        output_filename = os.path.join(output_folder, f"page_{page_num + 1}.pdf")
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
