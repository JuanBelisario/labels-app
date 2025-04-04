# TOs Hub - Streamlit App with Labels Generator + PL Builder
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
import textwrap

# ---------- LABEL TEMPLATES ----------
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
    st.download_button("Download D2C Template", generate_d2c_template(), "d2c_labels_template.xlsx")
    st.download_button("Download FNSKU Template", generate_fnsku_template(), "fnsku_labels_template.xlsx")

# ---------- BARCODE UTILITIES ----------
def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', name)

def generate_d2c_barcode(upc_code, sku):
    barcode = Code128(str(upc_code), writer=ImageWriter())
    barcode_filename = clean_filename(sku) + "_barcode"
    barcode.save(barcode_filename)
    return barcode_filename + ".png"

def generate_fnsku_barcode(fnsku):
    barcode = Code128(fnsku, writer=ImageWriter())
    barcode_filename = f"{fnsku}_barcode"
    barcode.save(barcode_filename)
    return f"{barcode_filename}.png"

# ---------- PDF RENDERING ----------
def wrap_text_to_two_lines(text, max_length, c, start_x, start_y, line_height, max_width):
    text = str(text) if pd.notna(text) else ""
    if len(text) > 2 * max_length:
        text_to_display = text[:20] + '...' + text[-20:]
    else:
        text_to_display = text
    lines = textwrap.wrap(text_to_display, width=25)
    lines = lines[:2] if len(lines) > 2 else lines
    for i, line in enumerate(lines):
        c.drawString(start_x, start_y - i * line_height, line)

def create_fnsku_pdf(barcode_image, fnsku, product_name, lot, output_folder):
    c = canvas.Canvas(os.path.join(output_folder, f"{fnsku}_fnsku_label.pdf"), pagesize=(60 * mm, 35 * mm))
    c.drawImage(barcode_image, 4.5 * mm, 10 * mm, width=51.5 * mm, height=16 * mm)
    c.setFont("Helvetica", 9)
    if product_name:
        wrap_text_to_two_lines(product_name, 22, c, 5 * mm, 7.75 * mm, 7.5, 25)
    if lot is not None:
        c.drawString(5 * mm, 3.5 * mm, f"Lot: {str(lot).strip()}")
    c.showPage()
    c.save()
    os.remove(barcode_image)

def generate_label_pdf(sku, upc_code, lot_num, output_path):
    c = canvas.Canvas(output_path, pagesize=(60 * mm, 35 * mm))
    c.setFont("Helvetica", 9.5)
    c.drawCentredString(30 * mm, 27.25 * mm, sku)

    barcode_path = generate_d2c_barcode(upc_code, sku)
    c.drawImage(barcode_path, 4.5 * mm, 9 * mm, width=51.5 * mm, height=16 * mm)
    os.remove(barcode_path)

    c.setFont("Helvetica", 9)
    lot_text = str(lot_num).strip() if lot_num is not None else ""
    c.setStrokeColorRGB(0, 0, 0)
    c.rect(10 * mm, 3.625 * mm, 40 * mm, 4 * mm, stroke=1, fill=0)
    c.drawCentredString(30 * mm, 4.75 * mm, lot_text)
    c.save()

# ---------- GENERATE PDFs ----------
def generate_pdfs_from_excel(df):
    if any(col not in df.columns for col in ['SKU', 'UPC Code', 'LOT#']):
        st.error("Missing required columns in uploaded file.")
        return None
    folder = f"{df.iloc[0]['SKU']}_{datetime.now().strftime('%Y%m%d')}"
    os.makedirs(folder, exist_ok=True)
    progress = st.progress(0)
    for idx, row in df.iterrows():
        generate_label_pdf(row['SKU'], row['UPC Code'], row['LOT#'], os.path.join(folder, clean_filename(f"{row['SKU']}.pdf")))
        progress.progress((idx + 1) / len(df))
    zip_path = folder + ".zip"
    with ZipFile(zip_path, 'w') as z:
        for file in os.listdir(folder):
            z.write(os.path.join(folder, file), file)
    return zip_path

def generate_fnsku_labels_from_excel(df):
    folder = f"{df.iloc[0]['FNSKU']}_{datetime.now().strftime('%Y%m%d')}"
    os.makedirs(folder, exist_ok=True)
    progress = st.progress(0)
    for idx, row in df.iterrows():
        barcode_img = generate_fnsku_barcode(row['FNSKU'])
        create_fnsku_pdf(barcode_img, row['FNSKU'], row['Product Name'], row['LOT#'], folder)
        progress.progress((idx + 1) / len(df))
    zip_path = folder + ".zip"
    with ZipFile(zip_path, 'w') as z:
        for file in os.listdir(folder):
            if "_fnsku_label" in file:
                z.write(os.path.join(folder, file), file)
    return zip_path
# ---------- PL BUILDER ----------
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

    to = df['TO'].iloc[0]
    so = df['FOP SO #'].iloc[0]
    from_loc = df['From Loc'].iloc[0]
    to_loc = df['To Loc'].iloc[0]
    total_qty = int(pd.to_numeric(df['Required Qty'], errors='coerce').sum())
    filename = f"{to} + {so} + {from_loc} + {to_loc} + {total_qty} Units.xlsx"

    headers = [
        "TO", "SO #", "From Loc", "To Loc", "Trafilea SKU",
        "Required Qty", "Shipping Method", "FG", "LOT", "Expiration Date", "CARTONS",
        "UNITS/Ctn", "Total QTY", "Carton Dimensions(inch) ",
        "Carton WEIGHT-LB", "Pallet Dimensions", "Pallet WEIGHT-LB.", "Pallet #"
    ]

    if transformation:
        headers.insert(5, "Destination SKU")

    output_df = pd.DataFrame(columns=headers)
    output_df['TO'] = df['TO']
    output_df['SO #'] = df['FOP SO #']
    output_df['From Loc'] = df['From Loc']
    output_df['To Loc'] = df['To Loc']
    output_df['Trafilea SKU'] = df['SKU External ID']
    output_df['Required Qty'] = df['Required Qty']
    output_df['Shipping Method'] = df['Shipping Method']
    if transformation:
        output_df['Destination SKU'] = df['Destination SKU']

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        output_df.to_excel(writer, index=False, sheet_name='PL')
        workbook = writer.book
        worksheet = writer.sheets['PL']
        dark_blue = workbook.add_format({'bold': True, 'bg_color': '#0C2D63', 'font_color': 'white', 'border': 1})
        light_blue = workbook.add_format({'bold': True, 'bg_color': '#D9EAF7', 'border': 1})
        for col_num, col_name in enumerate(output_df.columns):
            fmt = dark_blue if col_name in ["TO", "SO #", "From Loc", "To Loc", "Trafilea SKU", "Destination SKU", "Required Qty", "Shipping Method"] else light_blue
            worksheet.write(0, col_num, col_name, fmt)
            worksheet.set_column(col_num, col_num, 22)
    output.seek(0)
    return output, filename

# ---------- STREAMLIT UI ----------
st.set_page_config(page_title="TOs Hub", layout="wide")
st.title("TOs Hub")

st.sidebar.title("Navigation")
module = st.sidebar.radio("Go to:", ["Labels Generator", "PL Builder"])

if module == "Labels Generator":
    st.header("Labels Generator")
    show_template_download_buttons()
    option = st.selectbox("Choose an action", ["Generate D2C Labels", "Generate FNSKU Labels"])

    if option == "Generate D2C Labels":
        st.write("Upload an Excel file with SKU, UPC, and LOT# (if applicable)")
        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"], key="excel_uploader")
        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl' if uploaded_file.name.endswith('xlsx') else 'xlrd')
                if st.button("Generate D2C Labels"):
                    zip_path = generate_pdfs_from_excel(df)
                    if zip_path:
                        with open(zip_path, "rb") as f:
                            st.download_button("Download ZIP file with Labels", f, file_name=zip_path)
            except Exception as e:
                st.error(f"Error reading the Excel file: {e}")

    elif option == "Generate FNSKU Labels":
        st.write("Upload an Excel file with FNSKU, Product Name, and LOT#")
        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"], key="excel_fnsku_uploader")
        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl' if uploaded_file.name.endswith('xlsx') else 'xlrd')
                if st.button("Generate FNSKU Labels"):
                    zip_path = generate_fnsku_labels_from_excel(df)
                    if zip_path:
                        with open(zip_path, "rb") as f:
                            st.download_button("Download ZIP file with FNSKU Labels", f, file_name=zip_path)
            except Exception as e:
                st.error(f"Error reading the Excel file: {e}")

elif module == "PL Builder":
    st.header("📦 Packing List Builder")
    st.subheader("Generate PL")

    uploaded_files = st.file_uploader(
        "Upload one or more CSV or Excel files",
        type=["csv", "xls", "xlsx"],
        accept_multiple_files=True
    )

    if uploaded_files:
        zip_buffer = BytesIO()
        with ZipFile(zip_buffer, 'w') as zip_archive:
            for uploaded_file in uploaded_files:
                try:
                    if uploaded_file.name.endswith(".csv"):
                        df = pd.read_csv(uploaded_file)
                    else:
                        df = pd.read_excel(uploaded_file, engine='openpyxl' if uploaded_file.name.endswith('xlsx') else 'xlrd')

                    is_transformation = 'Destination SKU' in df.columns
                    output, filename = build_pl_base(df, transformation=is_transformation)

                    if output:
                        zip_archive.writestr(filename, output.getvalue())

                except Exception as e:
                    st.error(f"Error processing file '{uploaded_file.name}': {e}")

        zip_buffer.seek(0)
        if zip_buffer.getbuffer().nbytes > 0:
            st.success("All PL files processed successfully!")
            st.download_button(
                label="📥 Download All PLs as ZIP",
                data=zip_buffer,
                file_name="packing_lists.zip",
                mime="application/zip"
            )

    st.markdown(
        """
        <a href="https://docs.google.com/forms/d/e/1FAIpQLSfllE2UA33kBQpr5-Nq2tmDwhnYn9DStNyHRcKdONvpw0qTaQ/viewform" target="_blank">
            <button style='padding: 0.5em 1em; font-size: 16px;'>📧 Fill TO Template | Send Email</button>
        </a>
        """,
        unsafe_allow_html=True
    )
