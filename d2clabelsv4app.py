import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from barcode import EAN13, Code128  # Incluimos EAN13 y Code128 para FNSKU y UPC
from barcode.writer import ImageWriter
from datetime import datetime
from zipfile import ZipFile
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import textwrap

# Función para generar el archivo Excel de plantilla para D2C Labels
def generate_d2c_template():
    # Crear un DataFrame vacío con los encabezados correctos
    df = pd.DataFrame(columns=['SKU', 'UPC Code', 'LOT#'])
    
    # Guardar el DataFrame como un archivo Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='D2C Template')
    
    output.seek(0)
    return output

# Función para generar el archivo Excel de plantilla para FNSKU Labels
def generate_fnsku_template():
    # Crear un DataFrame vacío con los encabezados correctos
    df = pd.DataFrame(columns=['SKU', 'FNSKU', 'Product Name', 'LOT#'])
    
    # Guardar el DataFrame como un archivo Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='FNSKU Template')
    
    output.seek(0)
    return output

# Función para mostrar los botones de descarga de plantillas en Streamlit
def show_template_download_buttons():
    st.write("Download Templates for D2C Labels and FNSKU Labels:")

    # Botón para descargar el template D2C Labels
    d2c_template = generate_d2c_template()
    st.download_button(
        label="Download D2C Template",
        data=d2c_template,
        file_name="d2c_labels_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Botón para descargar el template FNSKU Labels
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

# Función para generar código de barras FNSKU
def generate_fnsku_barcode(fnsku, sku, output_folder):
    fnsku = str(fnsku) if pd.notna(fnsku) else ""  # Asegurarnos de que FNSKU sea una cadena y no NaN
    fnsku_barcode = Code128(fnsku, writer=ImageWriter())
    fnsku_barcode.writer.set_options({
        'module_width': 0.35,
        'module_height': 16,
        'font_size': 7.75,
        'text_distance': 4.5,
        'quiet_zone': 1.25,
        'dpi': 600
    })
    barcode_filename = os.path.join(output_folder, f"{sku}_barcode")
    fnsku_barcode.save(barcode_filename)
    return f"{barcode_filename}.png"

# Función para manejar el texto largo del nombre del producto en la etiqueta FNSKU
def wrap_text_to_two_lines(text, max_length, c, start_x, start_y, line_height, max_width):
    text = str(text) if pd.notna(text) else ""  # Asegurarnos de que text sea una cadena y no NaN
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

# Función para generar etiqueta FNSKU en PDF
def create_fnsku_pdf(barcode_image, fnsku, sku, product_name, lot, output_folder):
    pdf_filename = os.path.join(output_folder, f"{sku}_fnsku_label.pdf")
    c = canvas.Canvas(pdf_filename, pagesize=(59 * mm, 28.09 * mm))
    
    # Asegurarnos de que los valores no sean NaN
    fnsku = str(fnsku) if pd.notna(fnsku) else ""
    product_name = str(product_name) if pd.notna(product_name) else ""
    lot = str(lot) if pd.notna(lot) else ""

    # Ajustar la imagen del código de barras
    c.drawImage(barcode_image, 4 * mm, 10.5 * mm, width=52 * mm, height=15 * mm)
    
    font_size = 7
    c.setFont("Helvetica", font_size)

    # Ajustar el nombre del producto (2 líneas máximo)
    if product_name:
        wrap_text_to_two_lines(product_name, max_length=23, c=c, start_x=5 * mm, start_y=7.5 * mm, line_height=font_size + 2, max_width=38)
    
    # Añadir número de lote si está disponible
    if lot:
        c.drawString(5 * mm, 3 * mm, f"Lot: {lot}")
    
    c.showPage()
    c.save()

# Función para generar etiquetas D2C en PDF con UPC
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

# Función para generar PDFs y comprimirlos en un archivo ZIP
def generate_pdfs_from_excel(df, label_type="D2C"):
    required_columns = ['SKU', 'UPC Code', 'LOT#'] if label_type == "D2C" else ['SKU', 'FNSKU', 'Product Name', 'LOT#']
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
        sku = str(row['SKU']) if pd.notna(row['SKU']) else ''
        if label_type == "D2C":
            upc_code = str(row['UPC Code']).zfill(12)
            lot_num = str(row['LOT#']) if pd.notna(row['LOT#']) else ""
            pdf_filename = clean_filename(f"{sku}.pdf")
            pdf_path = os.path.join(output_folder, pdf_filename)
            generate_label_pdf(sku, upc_code, lot_num, pdf_path)
        elif label_type == "FNSKU":
            fnsku = str(row['FNSKU']) if pd.notna(row['FNSKU']) else ''
            product_name = str(row['Product Name']) if pd.notna(row['Product Name']) else ''
            lot_num = str(row['LOT#']) if pd.notna(row['LOT#']) else ""
            barcode_image = generate_fnsku_barcode(fnsku, sku, output_folder)
            create_fnsku_pdf(barcode_image, fnsku, sku, product_name, lot_num, output_folder)

        progress_bar.progress((index + 1) / total_rows)

    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for folder_name, subfolders, filenames in os.walk(output_folder):
            for filename in filenames:
                filepath = os.path.join(folder_name, filename)
                zipObj.write(filepath, os.path.basename(filepath))

    return zip_filename

# Función para extraer texto de una página usando pdfplumber
def extract_text_from_page(page):
    text = page.extract_text()
    if text:
        clean_text = re.sub(r'[^\w\s]', '', text)  # Remover caracteres especiales
        clean_text = "_".join(clean_text.split())  # Reemplazar espacios por guiones bajos
        if len(clean_text) > 5:
            return clean_text
    return None

# Función para dividir un PDF en varios PDFs, una página por archivo
def split_fnsku_pdf(uploaded_pdf):
    pdf_file = BytesIO(uploaded_pdf.read())  # Convertir el archivo subido a BytesIO
    input_pdf = PdfReader(pdf_file)
    total_pages = len(input_pdf.pages)

    output_folder = f"Split_FNSKU_{datetime.now().strftime('%Y%m%d')}"
    os.makedirs(output_folder, exist_ok=True)

    progress_bar = st.progress(0)

    pdf_file.seek(0)  # Resetear el puntero para pdfplumber
    with pdfplumber.open(pdf_file) as pdf:
        for page_num in range(total_pages):
            writer = PdfWriter()
            writer.add_page(input_pdf.pages[page_num])

            page = pdf.pages[page_num]
            page_text = extract_text_from_page(page)  # Extraer el texto de la página

            if page_text:  # Saltar páginas sin texto válido
                clean_filename_text = clean_filename(page_text)
                output_filename = os.path.join(output_folder, f"{clean_filename_text}_page_{page_num + 1}.pdf")
                with open(output_filename, 'wb') as output_pdf:
                    writer.write(output_pdf)

            progress_bar.progress((page_num + 1) / total_pages)

    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for folder_name, subfolders, filenames in os.walk(output_folder):
            for filename in filenames:
                if "unknown" not in filename:
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
                zip_path = generate_pdfs_from_excel(df, label_type="FNSKU")
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
