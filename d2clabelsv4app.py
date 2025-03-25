# TOs Hub - Streamlit App with Labels Generator + PL Builder
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

# =====================
# 📁 PL BUILDER MODULE
# =====================
def build_pl_base(df, transformation=False):
    df = df.copy()
    required_cols = [
        'TO', 'FOP SO #', 'From Loc', 'To Loc',
        'SKU External ID', 'Required Qty', 'Shipping Method'
    ]
    if transformation:
        required_cols.append('Destination SKU')

    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        st.error(f"Missing required columns: {', '.join(missing)}")
        return None, None

    # Filename generation
    to = df['TO'].iloc[0]
    so = df['FOP SO #'].iloc[0]
    from_loc = df['From Loc'].iloc[0]
    to_loc = df['To Loc'].iloc[0]
    total_qty = df['Required Qty'].sum()
    filename = f"{to} + {so} + {from_loc} + {to_loc} + {total_qty} Units.xlsx"

    # Add placeholder columns (adjust as needed)
    df['SKU Name'] = ""
    df['Status'] = "Pending"
    df['Notes'] = ""

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='PL')
        workbook = writer.book
        worksheet = writer.sheets['PL']

        # Format header
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'middle',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 18)

    output.seek(0)
    return output, filename

# =====================
# 🧾 LABELS MODULE
# =====================
def generate_d2c_template():
    df = pd.DataFrame(columns=['SKU', 'UPC Code', 'LOT#'])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='D2C Template')
    output.seek(0)
    return output

def generate_fnsku_template():
    df = pd.DataFrame(columns=['FNSKU', 'Product Name', 'LOT#'])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='FNSKU Template')
    output.seek(0)
    return output

def show_template_download_buttons():
    st.write("Download Templates for D2C Labels and FNSKU Labels:")
    d2c_template = generate_d2c_template()
    st.download_button("Download D2C Template", d2c_template, "d2c_labels_template.xlsx")
    fnsku_template = generate_fnsku_template()
    st.download_button("Download FNSKU Template", fnsku_template, "fnsku_labels_template.xlsx")

def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', name)

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

def generate_d2c_barcode(upc_code, sku):
    barcode_ean = EAN13(upc_code, writer=ImageWriter())
    clean_sku = clean_filename(sku)
    barcode_filename = f"{clean_sku}_barcode"
    barcode_ean.save(barcode_filename)
    return f"{barcode_filename}.png"

def wrap_text_to_two_lines(text, max_length, c, start_x, start_y, line_height, max_width):
    text = str(text) if pd.notna(text) else ""
    if len(text) > 2 * max_length:
        text_to_display = text[:20] + '...' + text[-20:]
    else:
        text_to_display = text
    lines = textwrap.wrap(text_to_display, width=25)
    if len(lines) > 2:
        lines = lines[:2]
    for i, line in enumerate(lines):
        c.drawString(start_x, start_y - i * line_height, line)

def create_fnsku_pdf(barcode_image, fnsku, product_name, lot, output_folder):
    pdf_filename = os.path.join(output_folder, f"{fnsku}_fnsku_label.pdf")
    c = canvas.Canvas(pdf_filename, pagesize=(60 * mm, 35 * mm))
    c.drawImage(barcode_image, 4.5 * mm, 10 * mm, width=51.5 * mm, height=16 * mm)
    c.setFont("Helvetica", 9)
    if product_name:
        wrap_text_to_two_lines(product_name, 22, c, 5 * mm, 7.75 * mm, 7.5, 25)
    if lot:
        c.drawString(5 * mm, 3.5 * mm, f"Lot: {lot}")
    c.showPage()
    c.save()
    if os.path.exists(barcode_image):
        os.remove(barcode_image)

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

def generate_pdfs_from_excel(df):
    required_columns = ['SKU', 'UPC Code', 'LOT#']
    if any(col not in df.columns for col in required_columns):
        st.error("Missing required columns in Excel file.")
        return None
    first_sku = df.iloc[0]['SKU']
    current_date = datetime.now().strftime("%Y%m%d")
    output_folder = f"{first_sku}_{current_date}"
    os.makedirs(output_folder, exist_ok=True)
    progress_bar = st.progress(0)
    for index, row in df.iterrows():
        sku = row['SKU']
        upc_code = str(row['UPC Code']).zfill(12)
        lot_num = row['LOT#'] if pd.notnull(row['LOT#']) else ""
        pdf_filename = clean_filename(f"{sku}.pdf")
        pdf_path = os.path.join(output_folder, pdf_filename)
        generate_label_pdf(sku, upc_code, lot_num, pdf_path)
        progress_bar.progress((index + 1) / len(df))
    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for f in os.listdir(output_folder):
            zipObj.write(os.path.join(output_folder, f), f)
    return zip_filename

def generate_fnsku_labels_from_excel(df):
    first_fnsku = df.iloc[0]['FNSKU']
    current_date = datetime.now().strftime("%Y%m%d")
    output_folder = f"{first_fnsku}_{current_date}"
    os.makedirs(output_folder, exist_ok=True)
    progress_bar = st.progress(0)
    for index, row in df.iterrows():
        fnsku = str(row['FNSKU']) if pd.notna(row['FNSKU']) else ""
        product_name = str(row['Product Name']) if pd.notna(row['Product Name']) else ""
        lot = str(row['LOT#']) if pd.notna(row['LOT#']) else ""
        barcode_image = generate_fnsku_barcode(fnsku)
        create_fnsku_pdf(barcode_image, fnsku, product_name, lot, output_folder)
        progress_bar.progress((index + 1) / len(df))
    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for f in os.listdir(output_folder):
            if f.endswith("_fnsku_label.pdf"):
                zipObj.write(os.path.join(output_folder, f), f)
    return zip_filename

# =====================
# 🚀 STREAMLIT APP
# =====================
st.set_page_config(page_title="TOs Hub", layout="wide")
st.title("TOs Hub")

st.sidebar.title("Navigation")
module = st.sidebar.radio("Go to:", ["Labels Generator", "PL Builder"])

if module == "Labels Generator":
    st.header("Labels Generator")
    show_template_download_buttons()
    option = st.selectbox("Choose an action", ["Generate D2C Labels", "Generate FNSKU Labels"], key="action_select")

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

elif module == "PL Builder":
    st.header("📦 Packing List Builder")
    pl_type = st.selectbox("Select PL Type", ["Normal TO PL", "Transformation TO PL"])
    uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])

    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            is_transformation = pl_type == "Transformation TO PL"
            output, filename = build_pl_base(df, transformation=is_transformation)
            if output:
                st.success("PL file generated successfully!")
                st.download_button(
                    label="📥 Download PL Excel",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error processing file: {e}")
