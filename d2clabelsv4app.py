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
# 📁 TEMPLATE FUNCTIONS
# =====================
def generate_d2c_template():
    df = pd.DataFrame(columns=['SKU', 'UPC Code', 'LOT#'])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='D2C Template')
    output.seek(0)
    return output

def generate_fnsku_template():
    df = pd.DataFrame(columns=['SKU', 'FNSKU', 'Product Name', 'LOT#'])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='FNSKU Template')
    output.seek(0)
    return output

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

    # Validate numeric and non-empty fields
    errors = []
    if df['Required Qty'].isnull().any() or not pd.api.types.is_numeric_dtype(df['Required Qty']):
        errors.append("Required Qty must be a numeric and non-null column.")

    for field in ['TO', 'FOP SO #', 'From Loc', 'To Loc', 'SKU External ID']:
        if df[field].isnull().any():
            errors.append(f"Column '{field}' contains missing values.")

    if errors:
        for err in errors:
            st.error(err)
        return None, None

    # Filename generation (fix sum to int)
    to = df['TO'].iloc[0]
    so = df['FOP SO #'].iloc[0]
    from_loc = df['From Loc'].iloc[0]
    to_loc = df['To Loc'].iloc[0]
    total_qty = int(df['Required Qty'].sum())
    filename = f"{to} + {so} + {from_loc} + {to_loc} + {total_qty} Units.xlsx"

    # Define full header structure for output
    headers = [
        "TO", "SO #", "From Loc", "To Loc", "Trafilea SKU", "Destination SKU", "Required Qty",
        "Shipping Method", "FG", "Trafilea SKU", "LOT", "Expiration Date", "CARTONS",
        "UNITS/Ctn", "Total QTY", "Carton Dimensions(inch) ", "Carton WEIGHT-LB",
        "Pallet Dimensions", "Pallet WEIGHT-LB.", "Pallet #"
    ]

    output_df = pd.DataFrame(columns=headers)

    # Fill known data into the correct headers
    output_df['TO'] = df['TO']
    output_df['SO #'] = df['FOP SO #']
    output_df['From Loc'] = df['From Loc']
    output_df['To Loc'] = df['To Loc']
    output_df['Trafilea SKU'] = df['SKU External ID']
    output_df['Required Qty'] = df['Required Qty']
    output_df['Shipping Method'] = df['Shipping Method']

    # If transformation, copy destination SKU
    if transformation and 'Destination SKU' in df.columns:
        output_df['Destination SKU'] = df['Destination SKU']

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        output_df.to_excel(writer, index=False, sheet_name='PL')
        workbook = writer.book
        worksheet = writer.sheets['PL']

        dark_blue = workbook.add_format({
            'bold': True,
            'bg_color': '#0C2D63',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        light_blue = workbook.add_format({
            'bold': True,
            'bg_color': '#D9EAF7',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        for col_num, column_name in enumerate(output_df.columns):
            if column_name in ["TO", "SO #", "From Loc", "To Loc", "Trafilea SKU", "Destination SKU", "Required Qty", "Shipping Method"]:
                worksheet.write(0, col_num, column_name, dark_blue)
            else:
                worksheet.write(0, col_num, column_name, light_blue)
            worksheet.set_column(col_num, col_num, 22)

    output.seek(0)
    return output, filename

# =====================
# 📁 LABEL GENERATION FUNCTIONS
# =====================
def clean_filename(filename):
    # Remove invalid characters from filename
    return re.sub(r'[<>:"/\\|?*]', '', filename)

def generate_label_pdf(sku, upc_code, lot_num, output_path):
    # Create PDF with barcode
    c = canvas.Canvas(output_path, pagesize=(mm(100), mm(50)))
    
    # Draw barcode
    ean = EAN13(upc_code, writer=ImageWriter())
    barcode_path = f"temp_{sku}.png"
    ean.save(barcode_path)
    
    # Add barcode image to PDF
    c.drawImage(barcode_path, mm(5), mm(5), width=mm(40), height=mm(20))
    
    # Add text
    c.setFont("Helvetica-Bold", 12)
    c.drawString(mm(50), mm(30), f"SKU: {sku}")
    c.drawString(mm(50), mm(20), f"UPC: {upc_code}")
    if lot_num:
        c.drawString(mm(50), mm(10), f"LOT: {lot_num}")
    
    # Clean up temporary file
    os.remove(barcode_path)
    c.save()

def generate_fnsku_label_pdf(sku, fnsku, product_name, lot_num, output_path):
    # Create PDF with barcode
    c = canvas.Canvas(output_path, pagesize=(mm(100), mm(50)))
    
    # Draw barcode
    code128 = Code128(fnsku, writer=ImageWriter())
    barcode_path = f"temp_{fnsku}.png"
    code128.save(barcode_path)
    
    # Add barcode image to PDF
    c.drawImage(barcode_path, mm(5), mm(5), width=mm(40), height=mm(20))
    
    # Add text
    c.setFont("Helvetica-Bold", 12)
    c.drawString(mm(50), mm(30), f"SKU: {sku}")
    c.drawString(mm(50), mm(20), f"FNSKU: {fnsku}")
    c.drawString(mm(50), mm(15), f"Product: {product_name}")
    if lot_num:
        c.drawString(mm(50), mm(10), f"LOT: {lot_num}")
    
    # Clean up temporary file
    os.remove(barcode_path)
    c.save()

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

def generate_fnsku_labels_from_excel(df):
    required_columns = ['SKU', 'FNSKU', 'Product Name', 'LOT#']
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
        fnsku = row['FNSKU']
        product_name = row['Product Name']
        lot_num = row['LOT#'] if pd.notnull(row['LOT#']) else ""
        pdf_filename = clean_filename(f"{fnsku}.pdf")
        pdf_path = os.path.join(output_folder, pdf_filename)
        generate_fnsku_label_pdf(sku, fnsku, product_name, lot_num, pdf_path)
        progress_bar.progress((index + 1) / total_rows)

    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for folder_name, subfolders, filenames in os.walk(output_folder):
            for filename in filenames:
                filepath = os.path.join(folder_name, filename)
                zipObj.write(filepath, os.path.basename(filepath))

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
        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"], key="excel_uploader")
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
        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"], key="excel_fnsku_uploader")
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
    uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xls", "xlsx"])

    if uploaded_file is not None:
        try:
            # Read the file content into a BytesIO object
            file_content = BytesIO(uploaded_file.getvalue())
            
            if uploaded_file.name.endswith(".csv"):
                # For CSV files, read as string first to handle encoding
                file_content.seek(0)
                df = pd.read_csv(file_content, encoding='utf-8')
            else:
                # For Excel files
                file_content.seek(0)
                df = pd.read_excel(file_content, engine='openpyxl')

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
            st.write("Please make sure your file is in the correct format (CSV or Excel) and contains the required columns.")
