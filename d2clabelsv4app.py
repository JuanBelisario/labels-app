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
    df = pd.DataFrame(columns=['FNSKU', 'Product Name', 'LOT#'])
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

# Función para generar código de barras FNSKU (Code128) como imagen temporal
def generate_fnsku_barcode(fnsku):
    fnsku_barcode = Code128(fnsku, writer=ImageWriter())
    fnsku_barcode.writer.set_options({
        'module_width': 0.35,
        'module_height': 16,
        'font_size': 7.75,
        'text_distance': 4.5,
        'quiet_zone': 1.25,
        'dpi': 600
    })
    barcode_filename = f"{fnsku}_barcode"
    fnsku_barcode.save(barcode_filename)
    return f"{barcode_filename}.png"

# Función para generar código de barras EAN13 (D2C) como imagen temporal
def generate_d2c_barcode(upc_code, sku):
    barcode_ean = EAN13(upc_code, writer=ImageWriter())
    # Clean the SKU to create a valid filename
    clean_sku = clean_filename(sku)
    barcode_filename = f"{clean_sku}_barcode"
    barcode_ean.save(barcode_filename)
    return f"{barcode_filename}.png"

# Función para manejar el texto largo del nombre del producto en la etiqueta FNSKU
def wrap_text_to_two_lines(text, max_length, c, start_x, start_y, line_height, max_width):
    text = str(text) if pd.notna(text) else ""
    
    # Ensure the text is not too long for the given max_length
    if len(text) > 2 * max_length:
        text_to_display = text[:20] + '...' + text[-20:]  # First 20 + ... + Last 20
    else:
        text_to_display = text
    
    # Wrap the text to fit within the max_width
    lines = textwrap.wrap(text_to_display, width=25)
    
    # Ensure we only have two lines
    if len(lines) > 2:
        lines = lines[:2]
    
    # Draw each line on the canvas
    for i, line in enumerate(lines):
        c.drawString(start_x, start_y - i * line_height, line)

# Función para crear el PDF de la etiqueta FNSKU
def create_fnsku_pdf(barcode_image, fnsku, product_name, lot, output_folder):
    pdf_filename = os.path.join(output_folder, f"{fnsku}_fnsku_label.pdf")
    c = canvas.Canvas(pdf_filename, pagesize=(60 * mm, 35 * mm))
    
    # Dibujar código de barras
    c.drawImage(barcode_image, 4.5 * mm, 10 * mm, width=51.5 * mm, height=16 * mm)
    
    # Configurar la fuente y tamaño para el texto
    font_size = 9  # Back to original size
    c.setFont("Helvetica", font_size)

    # Ajustar el nombre del producto
    if product_name:
        wrap_text_to_two_lines(product_name, max_length=22, c=c, start_x=5 * mm, start_y=7.75 * mm, line_height=font_size - 1.5, max_width=25)

    # Añadir el número de lote si está disponible
    if lot:
        c.drawString(5 * mm, 3.5 * mm, f"Lot: {lot}")

    c.showPage()
    c.save()

    # Eliminar el archivo PNG temporal después de usarlo
    if os.path.exists(barcode_image):
        os.remove(barcode_image)

# Función para generar UPC labels (D2C) en PDF
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

    barcode_path = generate_d2c_barcode(upc_code, sku)

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

# Función para generar PDFs para D2C y comprimirlos en un ZIP
def generate_pdfs_from_excel(df):
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
        upc_code = str(row['UPC Code']).zfill(12)
        lot_num = row['LOT#'] if pd.notnull(row['LOT#']) else ""
        pdf_filename = clean_filename(f"{sku}.pdf")
        pdf_path = os.path.join(output_folder, pdf_filename)
        generate_label_pdf(sku, upc_code, lot_num, pdf_path)

        progress_bar.progress((index + 1) / total_rows)

    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for folder_name, subfolders, filenames in os.walk(output_folder):
            for filename in filenames:
                filepath = os.path.join(folder_name, filename)
                zipObj.write(filepath, os.path.basename(filepath))

    return zip_filename

# Función para generar PDFs y comprimirlos en un archivo ZIP (FNSKU)
def generate_fnsku_labels_from_excel(df):
    first_fnsku = df.iloc[0]['FNSKU']
    current_date = datetime.now().strftime("%Y%m%d")
    output_folder = f"{first_fnsku}_{current_date}"
    os.makedirs(output_folder, exist_ok=True)

    total_rows = len(df)
    progress_bar = st.progress(0)

    for index, row in df.iterrows():
        fnsku = str(row['FNSKU']) if pd.notna(row['FNSKU']) else ""
        product_name = str(row['Product Name']) if pd.notna(row['Product Name']) else ""
        lot = str(row['LOT#']) if pd.notna(row['LOT#']) else ""
        
        # Generar el código de barras FNSKU temporalmente
        barcode_image = generate_fnsku_barcode(fnsku)

        # Crear el PDF con la etiqueta FNSKU y eliminar el PNG después
        create_fnsku_pdf(barcode_image, fnsku, product_name, lot, output_folder)

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

# Opciones del menú
option = st.selectbox("Choose an action", ["Generate D2C Labels", "Generate FNSKU Labels"], key="action_select")

# Opción: Generate D2C Labels
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
