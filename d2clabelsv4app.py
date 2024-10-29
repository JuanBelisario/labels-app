import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from barcode import EAN13, Code128
from barcode.writer import ImageWriter
from datetime import datetime
from zipfile import ZipFile
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import textwrap

# Generate D2C Labels Excel template
def generate_d2c_template():
    df = pd.DataFrame(columns=['SKU', 'UPC Code', 'LOT#'])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='D2C Template')
    output.seek(0)
    return output

# Generate FNSKU Labels Excel template
def generate_fnsku_template():
    df = pd.DataFrame(columns=['SKU', 'FNSKU', 'Product Name', 'LOT#'])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='FNSKU Template')
    output.seek(0)
    return output

# Show template download buttons in Streamlit
def show_template_download_buttons():
    st.write("Download Templates for D2C Labels and FNSKU Labels:")
    d2c_template = generate_d2c_template()
    st.download_button("Download D2C Template", data=d2c_template, file_name="d2c_labels_template.xlsx")
    fnsku_template = generate_fnsku_template()
    st.download_button("Download FNSKU Template", data=fnsku_template, file_name="fnsku_labels_template.xlsx")

# Clean up file names
def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', name)

# Generate temporary FNSKU barcode (Code128)
def generate_fnsku_barcode(fnsku, sku):
    barcode = Code128(fnsku, writer=ImageWriter())
    barcode.writer.set_options({
        'module_width': 0.35,
        'module_height': 16,
        'font_size': 7.75,
        'text_distance': 4.5,
        'quiet_zone': 1.25,
        'dpi': 600
    })
    filename = f"{sku}_barcode.png"
    barcode.save(filename)
    return filename

# Generate temporary D2C barcode (EAN13)
def generate_d2c_barcode(upc_code, sku):
    barcode = EAN13(upc_code, writer=ImageWriter())
    filename = f"{sku}_barcode.png"
    barcode.save(filename)
    return filename

# Handle long text in FNSKU label
def wrap_text_to_two_lines(text, max_length, c, start_x, start_y, line_height, max_width):
    text = str(text) if pd.notna(text) else ""
    lines = textwrap.wrap(text[:max_length] + '...' + text[-max_length:], width=max_width)[:2]
    for i, line in enumerate(lines):
        c.drawString(start_x, start_y - i * line_height, line)

# Create FNSKU label PDF
def create_fnsku_pdf(barcode_image, fnsku, sku, product_name, lot, output_folder):
    pdf_filename = os.path.join(output_folder, f"{sku}_fnsku_label.pdf")
    c = canvas.Canvas(pdf_filename, pagesize=(60 * mm, 35 * mm))
    c.drawImage(barcode_image, 4.5 * mm, 10 * mm, width=51.5 * mm, height=16 * mm)
    c.setFont("Helvetica", 9)
    if product_name:
        wrap_text_to_two_lines(product_name, 23, c, start_x=5 * mm, start_y=7.75 * mm, line_height=11, max_width=38)
    if lot:
        c.drawString(5 * mm, 3.5 * mm, f"Lot: {lot}")
    c.showPage()
    c.save()

# Generate D2C label PDF
def generate_label_pdf(sku, upc_code, lot_num, output_path):
    width, height = 60 * mm, 35 * mm
    c = canvas.Canvas(output_path, pagesize=(width, height))
    c.setFont("Helvetica", 9.5)
    c.drawCentredString(width / 2, height - 7.75 * mm, sku)
    barcode_path = generate_d2c_barcode(upc_code, sku)
    c.drawImage(barcode_path, 4.5 * mm, height / 2 - 8 * mm, width=51.5 * mm, height=16 * mm)
    os.remove(barcode_path)
    if lot_num:
        c.setFont("Helvetica", 9)
        c.drawCentredString(width / 2, 4.75 * mm, f"Lot: {lot_num}")
    c.save()

# Generate and ZIP PDFs for D2C
def generate_pdfs_from_excel(df):
    if df.empty:
        st.error("The uploaded Excel file is empty or not valid.")
        return None

    first_sku = df.iloc[0]['SKU']
    output_folder = f"{first_sku}_{datetime.now().strftime('%Y%m%d')}"
    os.makedirs(output_folder, exist_ok=True)

    for index, row in df.iterrows():
        sku = row['SKU']
        upc_code = str(row['UPC Code']).zfill(12)
        lot_num = row['LOT#'] if pd.notnull(row['LOT#']) else ""
        pdf_path = os.path.join(output_folder, f"{clean_filename(sku)}.pdf")
        generate_label_pdf(sku, upc_code, lot_num, pdf_path)

    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for filename in os.listdir(output_folder):
            zipObj.write(os.path.join(output_folder, filename), filename)

    return zip_filename

# Generate and ZIP PDFs for FNSKU
def generate_fnsku_labels_from_excel(df):
    if df.empty:
        st.error("The uploaded Excel file is empty or not valid.")
        return None

    first_fnsku = df.iloc[0]['FNSKU']
    output_folder = f"{first_fnsku}_{datetime.now().strftime('%Y%m%d')}"
    os.makedirs(output_folder, exist_ok=True)

    for index, row in df.iterrows():
        sku = row['SKU']
        fnsku = row['FNSKU']
        product_name = row.get('Product Name', '')
        lot = row.get('LOT#', '')
        
        barcode_image = generate_fnsku_barcode(fnsku, sku)
        create_fnsku_pdf(barcode_image, fnsku, sku, product_name, lot, output_folder)
        os.remove(barcode_image)  # Remove PNG after use

    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for filename in os.listdir(output_folder):
            if "_fnsku_label" in filename:  # Include only correct files
                zipObj.write(os.path.join(output_folder, filename), filename)

    return zip_filename

# Streamlit UI
st.title("Label Tools")
show_template_download_buttons()
option = st.selectbox("Choose an action", ["Generate D2C Labels", "Generate FNSKU Labels"], key="action_select")

# Generate D2C Labels
if option == "Generate D2C Labels":
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], key="excel_uploader")
    if uploaded_file and st.button("Generate D2C Labels"):
        df = pd.read_excel(uploaded_file)
        zip_path = generate_pdfs_from_excel(df)
        if zip_path:
            with open(zip_path, "rb") as f:
                st.download_button("Download ZIP file with Labels", f, file_name=zip_path)

# Generate FNSKU Labels
elif option == "Generate FNSKU Labels":
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], key="excel_fnsku_uploader")
    if uploaded_file and st.button("Generate FNSKU Labels"):
        df = pd.read_excel(uploaded_file)
        zip_path = generate_fnsku_labels_from_excel(df)
        if zip_path:
            with open(zip_path, "rb") as f:
                st.download_button("Download ZIP file with FNSKU Labels", f, file_name=zip_path)
