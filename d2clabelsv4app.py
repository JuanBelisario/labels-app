# Instalar las dependencias necesarias
!pip install reportlab python-barcode pandas streamlit > /dev/null 2>&1

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

# Función para limpiar el nombre del archivo
def limpiar_nombre_archivo(nombre):
    return re.sub(r'[<>:"/\\|?*]', '', nombre)

# Función para generar etiquetas en formato PDF
def generar_etiqueta_pdf(sku, upc_code, lot_num, output_path):
    width, height = 60 * mm, 35 * mm
    c = canvas.Canvas(output_path, pagesize=(width, height))

    x_margin = 4.5 * mm
    y_sku = height - 7.75 * mm
    y_barcode = height / 2 - 8 * mm
    y_lot = 4.75 * mm
    ancho_barcode = 51.5 * mm

    c.setFont("Helvetica", 9.5)
    c.drawCentredString(width / 2, y_sku, sku)

    if len(upc_code) == 12:
        upc_code = '0' + upc_code

    barcode_filename = limpiar_nombre_archivo(f"{sku}_barcode")
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

    c.drawImage(barcode_path, (width - ancho_barcode) / 2, y_barcode, width=ancho_barcode, height=16 * mm)
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

# Función para generar los PDFs y comprimirlos
def generar_pdfs_desde_excel(df):
    primer_sku = df.iloc[0]['SKU']
    fecha_actual = datetime.now().strftime("%Y%m%d")

    # Crear una carpeta temporal
    carpeta_salida = f"{primer_sku}_{fecha_actual}"
    os.makedirs(carpeta_salida, exist_ok=True)

    for index, row in df.iterrows():
        sku = row['SKU']
        upc_code = str(row['UPC Code']).zfill(12)
        lot_num = row['LOT#'] if pd.notnull(row['LOT#']) else ""
        nombre_archivo_pdf = limpiar_nombre_archivo(f"{sku}.pdf")
        ruta_pdf = os.path.join(carpeta_salida, nombre_archivo_pdf)
        generar_etiqueta_pdf(sku, upc_code, lot_num, ruta_pdf)

    # Comprimir todos los archivos en un ZIP
    zip_filename = f"{carpeta_salida}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for folder_name, subfolders, filenames in os.walk(carpeta_salida):
            for filename in filenames:
                filepath = os.path.join(folder_name, filename)
                zipObj.write(filepath, os.path.basename(filepath))

    return zip_filename

# Interfaz con Streamlit
st.title("Generador de Etiquetas UPC")

st.write("Cargue un archivo Excel con los SKU, UPC y LOT# (si aplica)")

# Cargar el archivo Excel
uploaded_file = st.file_uploader("Cargar archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    if st.button("Generar etiquetas"):
        zip_path = generar_pdfs_desde_excel(df)
        with open(zip_path, "rb") as f:
            st.download_button("Descargar etiquetas ZIP", f, file_name=zip_path)
