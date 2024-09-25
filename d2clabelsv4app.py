import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128  # Usando Code128 desde reportlab
from datetime import datetime
from zipfile import ZipFile
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber

# Function to clean up file names
def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', name)

# Function to generate UPC labels in PDF format (Excel upload for D2C)
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

    barcode = code128.Code128(upc_code, barHeight=16 * mm, humanReadable=True)  # Usando Code128 desde reportlab
    barcode.drawOn(c, (width - barcode_width) / 2, y_barcode)

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

# New Function to generate FNSKU labels in PDF format using Code128 (Excel upload for AMZ)
def generate_fnsku_label_pdf(sku, fnsku_code, lot_num, output_path):
    width, height = 60 * mm, 35 * mm
    c = canvas.Canvas(output_path, pagesize=(width, height))

    x_margin = 4.5 * mm
    y_sku = height - 7.75 * mm
    y_barcode = height / 2 - 8 * mm
    y_lot = 4.75 * mm
    barcode_width = 51.5 * mm

    c.setFont("Helvetica", 9.5)
    c.drawCentredString(width / 2, y_sku, sku)

    barcode = code128.Code128(fnsku_code, barHeight=16 * mm, humanReadable=True)  # Usando Code128 desde reportlab
    barcode.drawOn(c, (width - barcode_width) / 2, y_barcode)

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

# Function to generate PDFs and compress them into a ZIP file (Excel upload for AMZ)
def generate_amz_pdfs_from_excel(df):
    required_columns = ['SKU', 'UPC Code', 'LOT#']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Missing columns in the Excel file: {', '.join(missing_columns)}")
        return None

    first_sku = df.iloc[0]['SKU']
    current_date = datetime.now().strftime("%Y%m%d")

    output_folder = f"{first_sku}_{current_date}"
    os.makedirs(output_folder, exist_ok=True)

    total_rows = len(df)
    progress_bar = st.progress(0)

    for index, row in df.iterrows():
        sku = row['SKU']
        fnsku_code = str(row['UPC Code']).zfill(12)
        lot_num = row['LOT#'] if pd.notnull(row['LOT#']) else ""
        pdf_filename = clean_filename(f"{sku}.pdf")
        pdf_path = os.path.join(output_folder, pdf_filename)
        generate_fnsku_label_pdf(sku, fnsku_code, lot_num, pdf_path)

        progress_bar.progress((index + 1) / total_rows)

    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for folder_name, subfolders, filenames in os.walk(output_folder):
            for filename in filenames:
                filepath = os.path.join(folder_name, filename)
                zipObj.write(filepath, os.path.basename(filepath))

    return zip_filename

# Streamlit interface
st.title("Label Tools")

option = st.selectbox("Choose an action", ["Generate D2C Labels", "Generate AMZ Labels", "Split FNSKUs PDF"], key="action_select")

if option == "Generate D2C Labels":
    st.write("Upload an Excel file with SKU, UPC, and LOT# (if applicable)")
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], key="excel_uploader")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            if st.button("Generate D2C Labels", key="generate_d2c_labels"):
                zip_path = generate_pdfs_from_excel(df)
                if zip_path:
                    with open(zip_path, "rb") as f:
                        st.download_button("Download ZIP file with Labels", f, file_name=zip_path)
        except Exception as e:
            st.error(f"Error reading the Excel file: {e}")

elif option == "Generate AMZ Labels":
    st.write("Upload an Excel file with SKU, FNSKU (as UPC Code), and LOT# (if applicable)")
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], key="excel_amz_uploader")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            if st.button("Generate AMZ Labels", key="generate_amz_labels"):
                zip_path = generate_amz_pdfs_from_excel(df)
                if zip_path:
                    with open(zip_path, "rb") as f:
                        st.download_button("Download ZIP file with AMZ Labels", f, file_name=zip_path)
        except Exception as e:
            st.error(f"Error reading the Excel file: {e}")

elif option == "Split FNSKUs PDF":
    st.write("Upload a PDF file to split FNSKU labels")
    uploaded_pdf = st.file_uploader("Upload PDF file", type=["pdf"], key="pdf_uploader")

    if uploaded_pdf is not None:
        if st.button("Split PDF", key="split_pdf"):
            try:
                zip_path = split_fnsku_pdf(uploaded_pdf)
                if zip_path:
                    with open(zip_path, "rb") as f:
                        st.download_button("Download ZIP file with Split PDFs", f, file_name=zip_path)
            except Exception as e:
                st.error(f"Error processing the PDF file: {e}")
