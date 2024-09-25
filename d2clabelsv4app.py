import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from barcode import Code128
from barcode.writer import ImageWriter
from datetime import datetime
from zipfile import ZipFile
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import textwrap

# Función para generar el archivo Excel de plantilla para D2C Labels
def generate_d2c_template():
    df = pd.DataFrame(columns=['SKU', 'UPC Code', 'LOT#'])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='D2C Template')
    output.seek(0)
    return output

# Función para generar el archivo Excel de plantilla para FNSKU Labels
def generate_fnsku_template():
    df = pd.DataFrame(columns=['SKU', 'FNSKU', 'Product Name', 'LOT#'])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='FNSKU Template')
    output.seek(0)
    return output

# Función para mostrar los botones de descarga de plantillas en Streamlit
def show_template_download_buttons():
    st.write("Download Templates for D2C Labels and FNSKU Labels:")
    d2c_template = generate_d2c_template()
    st.download_button(
        label="Download D2C Template",
        data=d2c_template,
        file_name="d2c_labels_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    fnsku_template = generate_fnsku_template()
    st.download_button(
        label="Download FNSKU Template",
        data=fnsku_template,
        file_name="fnsku_labels_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Función para limpiar nombres de archivos
def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', name)

# Función para generar código de barras FNSKU como imagen temporal
def generate_fnsku_barcode(fnsku, sku):
    fnsku_barcode = Code128(fnsku, writer=ImageWriter())
    fnsku_barcode.writer.set_options({
        'module_width': 0.35,
        'module_height': 16,
        'font_size': 7.75,
        'text_distance': 4.5,
        'quiet_zone': 1.25,
        'dpi': 600
    })
    barcode_filename = f"{sku}_barcode"
    fnsku_barcode.save(barcode_filename)
    return f"{barcode_filename}.png"

# Función para manejar el texto largo del nombre del producto en la etiqueta FNSKU
def wrap_text_to_two_lines(text, max_length, c, start_x, start_y, line_height, max_width):
    text = str(text) if pd.notna(text) else ""
    if len(text) > 2 * max_length:
        text_to_display = text[:max_length] + '...' + text[-max_length:]
    else:
        text_to_display = text
    
    lines = textwrap.wrap(text_to_display, width=max_width)
    if len(lines) > 2:
        lines = lines[:2]
        lines[-1] = lines[-1][:max_width - 3] + '...'

    for i, line in enumerate(lines):
        c.drawString(start_x, start_y - i * line_height, line)

# Función para crear el PDF de la etiqueta FNSKU
def create_fnsku_pdf(barcode_image, fnsku, sku, product_name, lot, output_folder):
    pdf_filename = os.path.join(output_folder, f"{sku}_fnsku_label.pdf")
    c = canvas.Canvas(pdf_filename, pagesize=(60 * mm, 35 * mm))
    c.drawImage(barcode_image, 4.5 * mm, 10 * mm, width=51.5 * mm, height=16 * mm)
    font_size = 9
    c.setFont("Helvetica", font_size)
    if product_name:
        wrap_text_to_two_lines(product_name, max_length=23, c=c, start_x=5 * mm, start_y=7.75 * mm, line_height=font_size + 2, max_width=38)
    if lot:
        c.drawString(5 * mm, 3.5 * mm, f"Lot: {lot}")
    c.showPage()
    c.save()

    # Eliminar el archivo PNG temporal después de usarlo
    if os.path.exists(barcode_image):
        os.remove(barcode_image)

# Función para generar PDFs y comprimirlos en un archivo ZIP (FNSKU)
def generate_fnsku_labels_from_excel(df):
    first_fnsku = df.iloc[0]['FNSKU']
    current_date = datetime.now().strftime("%Y%m%d")
    output_folder = f"{first_fnsku}_{current_date}"
    os.makedirs(output_folder, exist_ok=True)

    total_rows = len(df)
    progress_bar = st.progress(0)

    for index, row in df.iterrows():
        sku = str(row['SKU']) if pd.notna(row['SKU']) else ''
        fnsku = str(row['FNSKU']) if pd.notna(row['FNSKU']) else ''
        product_name = str(row['Product Name']) if pd.notna(row['Product Name']) else ''
        lot = str(row['LOT#']) if pd.notna(row['LOT#']) else ''
        
        # Generar el código de barras FNSKU temporalmente
        barcode_image = generate_fnsku_barcode(fnsku, sku)

        # Crear el PDF con la etiqueta FNSKU y eliminar el PNG después
        create_fnsku_pdf(barcode_image, fnsku, sku, product_name, lot, output_folder)

        progress_bar.progress((index + 1) / total_rows)

    # Comprimir solo los PDFs que tengan el sufijo "_fnsku_label" en el nombre
    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for folder_name, subfolders, filenames in os.walk(output_folder):
            for filename in filenames:
                if "_fnsku_label" in filename:  # Solo incluir los archivos correctos
                    filepath = os.path.join(folder_name, filename)
                    zipObj.write(filepath, os.path.basename(filepath))

    return zip_filename

# Streamlit UI
st.title("Label Tools")

# Mostrar los botones de descarga de plantillas
show_template_download_buttons()

# Opciones en el menú de la app
option = st.selectbox("Choose an action", ["Generate D2C Labels", "Generate FNSKU Labels", "Split FNSKU PDFs"], key="action_select")

# Opción: Generate D2C Labels
if option == "Generate D2C Labels":
    st.write("Upload an Excel file with SKU, UPC, and LOT# (if applicable)")
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], key="excel_uploader")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            if st.button("Generate D2C Labels", key="generate_d2c_labels"):
                zip_path = generate_pdfs_from_excel(df, label_type="D2C")
                if zip_path:
                    with open(zip_path, "rb") as f:
                        st.download_button("Download ZIP file with Labels", f, file_name=zip_path)
        except Exception as e:
            st.error(f"Error reading the Excel file: {e}")

# Opción: Generate FNSKU Labels
elif option == "Generate FNSKU Labels":
    st.write("Upload an Excel file with SKU, FNSKU, and LOT# (if applicable)")
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], key="excel_fnsku_uploader")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            if st.button("Generate FNSKU Labels", key="generate_fnsku_labels"):
                zip_path = generate_fnsku_labels_from_excel(df)
                if zip_path:
                    with open(zip_path, "rb") as f:
                        st.download_button("Download ZIP file with FNSKU Labels", f, file_name=zip_path)
        except Exception as e:
            st.error(f"Error reading the Excel file: {e}")

# Opción: Split FNSKU PDFs
elif option == "Split FNSKU PDFs":
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
