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
from supabase import create_client

# Initialize Supabase client
@st.cache_resource
def init_connection():
    url = st.secrets["supabase_url"]
    key = st.secrets["supabase_key"]
    return create_client(url, key)

supabase = init_connection()

# Database functions
def update_sku_database_from_file(file):
    # Only show this function if user is admin
    if not st.session_state.get('is_admin', False):
        st.error("Unauthorized access")
        return
        
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
            
        if 'SKU' not in df.columns or 'UPC' not in df.columns:
            return False, "File must contain 'SKU' and 'UPC' columns"
        
        # Convert DataFrame to records
        records = [{
            'sku': str(row['SKU']),
            'upc': str(row['UPC'])
        } for _, row in df.iterrows()]
        
        # Delete all existing records with a proper filter
        supabase.table('sku_products').delete().neq('sku', '').execute()
        
        # Insert new records
        supabase.table('sku_products').insert(records).execute()
        
        return True, f"Successfully updated {len(df)} SKU products"
    except Exception as e:
        return False, f"Error updating database: {str(e)}"

def update_fnsku_database_from_file(file):
    # Only show this function if user is admin
    if not st.session_state.get('is_admin', False):
        st.error("Unauthorized access")
        return
        
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
            
        if 'FNSKU' not in df.columns or 'Product Name' not in df.columns:
            return False, "File must contain 'FNSKU' and 'Product Name' columns"
        
        # Convert DataFrame to records
        records = [{
            'fnsku': str(row['FNSKU']),
            'product_name': str(row['Product Name'])
        } for _, row in df.iterrows()]
        
        # Delete all existing records with a proper filter
        supabase.table('fnsku_products').delete().neq('fnsku', '').execute()
        
        # Insert new records
        supabase.table('fnsku_products').insert(records).execute()
        
        return True, f"Successfully updated {len(df)} FNSKU products"
    except Exception as e:
        return False, f"Error updating database: {str(e)}"

def get_upc_for_sku(sku: str) -> str:
    response = supabase.table('sku_products').select('upc').eq('sku', sku).execute()
    return response.data[0]['upc'] if response.data else None

def get_product_name_for_fnsku(fnsku: str) -> str:
    response = supabase.table('fnsku_products').select('product_name').eq('fnsku', fnsku).execute()
    return response.data[0]['product_name'] if response.data else None

# Simple admin authentication
def check_password():
    """Returns `True` if the user had the correct password."""
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["admin_password"]:
            st.session_state["password_correct"] = True
            st.session_state["is_admin"] = True
        else:
            st.session_state["password_correct"] = False
            st.session_state["is_admin"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct.
        return True

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
    barcode_filename = f"{fnsku}_barcode"
    fnsku_barcode.save(barcode_filename)
    return f"{barcode_filename}.png"

# Función para generar código de barras EAN13 (D2C) como imagen temporal
def generate_d2c_barcode(upc_code, sku):
    barcode_ean = EAN13(upc_code, writer=ImageWriter())
    barcode_filename = f"{sku}_barcode"
    barcode_ean.save(barcode_filename)
    return f"{barcode_filename}.png"

# Función para manejar el texto largo del nombre del producto en la etiqueta FNSKU
def wrap_text_to_two_lines(text, max_length, c, start_x, start_y, line_height, max_width):
    text = str(text) if pd.notna(text) else ""
    
    # Ensure the text is not too long for the given max_length
    if len(text) > 2 * max_length:
        text_to_display = text[:max_length] + '...' + text[-max_length:]
    else:
        text_to_display = text
    
    # Wrap the text to fit within the max_width
    lines = textwrap.wrap(text_to_display, width=max_width)
    
    # Ensure we only have two lines, truncate if necessary
    if len(lines) > 2:
        lines = lines[:2]
        lines[-1] = lines[-1][:max_width - 3] + '...'

    # Draw each line on the canvas
    for i, line in enumerate(lines):
        c.drawString(start_x, start_y - i * line_height, line)

# Función para crear el PDF de la etiqueta FNSKU
def create_fnsku_pdf(barcode_image, fnsku, sku, product_name, lot, output_folder):
    pdf_filename = os.path.join(output_folder, f"{fnsku}_fnsku_label.pdf")
    c = canvas.Canvas(pdf_filename, pagesize=(60 * mm, 35 * mm))
    
    # Dibujar código de barras
    c.drawImage(barcode_image, 4.5 * mm, 10 * mm, width=51.5 * mm, height=16 * mm)
    
    # Configurar la fuente y tamaño para el texto
    font_size = 9
    c.setFont("Helvetica", font_size)

    # Ajustar el nombre del producto
    if product_name:
        wrap_text_to_two_lines(product_name, max_length=21, c=c, start_x=5 * mm, start_y=7.75 * mm, line_height=font_size - 1.5, max_width=33.5)

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

    barcode_filename = clean_filename(f"{sku}_barcode")
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
        barcode_image = generate_fnsku_barcode(fnsku, fnsku)

        # Crear el PDF con la etiqueta FNSKU y eliminar el PNG después
        create_fnsku_pdf(barcode_image, fnsku, fnsku, product_name, lot, output_folder)

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

def main():
    st.title("Label Tools")
    
    # Show admin features only if logged in as admin
    if 'is_admin' not in st.session_state:
        st.session_state.is_admin = False
    
    # Admin login in sidebar
    with st.sidebar:
        if st.button("Admin Login"):
            check_password()
    
    # Show database management only to admin
    if st.session_state.is_admin:
        st.header("Database Management")
        db_tab1, db_tab2 = st.tabs(["SKU Database", "FNSKU Database"])
        
        with db_tab1:
            st.subheader("Update SKU-UPC Database")
            sku_file = st.file_uploader("Upload CSV/Excel file with SKU and UPC columns", 
                                      type=["csv", "xlsx"], 
                                      key="sku_db_upload")
            if sku_file:
                if st.button("Update SKU Database"):
                    success, message = update_sku_database_from_file(sku_file)
                    if success:
                        st.success(message)
                    else:
                        st.error(message)
        
        with db_tab2:
            st.subheader("Update FNSKU Database")
            fnsku_file = st.file_uploader("Upload CSV/Excel file with FNSKU and Product Name columns", 
                                        type=["csv", "xlsx"], 
                                        key="fnsku_db_upload")
            if fnsku_file:
                if st.button("Update FNSKU Database"):
                    success, message = update_fnsku_database_from_file(fnsku_file)
                    if success:
                        st.success(message)
                    else:
                        st.error(message)
    
    # Label Generation Section
    st.header("Label Generation")
    show_template_download_buttons()
    
    option = st.selectbox("Choose an action", ["Generate D2C Labels", "Generate FNSKU Labels"])
    
    if option == "Generate D2C Labels":
        input_method = st.radio("Choose input method", 
                              ["Upload File", "Enter SKUs manually"],
                              key="d2c_input_method")
        
        if input_method == "Upload File":
            uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], key="d2c_uploader")
            if uploaded_file:
                try:
                    df = pd.read_excel(uploaded_file)
                    if 'SKU' not in df.columns:
                        st.error("Excel file must contain 'SKU' column")
                    else:
                        # Add UPC codes from database
                        df['UPC Code'] = df['SKU'].apply(get_upc_for_sku)
                        missing_upcs = df[df['UPC Code'].isna()]['SKU'].tolist()
                        if missing_upcs:
                            st.error(f"UPC codes not found for SKUs: {', '.join(missing_upcs)}")
                        elif st.button("Generate D2C Labels"):
                            zip_path = generate_pdfs_from_excel(df)
                            if zip_path:
                                with open(zip_path, "rb") as f:
                                    st.download_button("Download ZIP file with Labels", 
                                                     f, 
                                                     file_name=zip_path)
                except Exception as e:
                    st.error(f"Error processing file: {str(e)}")
        
        else:  # Manual SKU entry
            skus_input = st.text_area("Enter SKUs (comma-separated)")
            lots_input = st.text_area("Enter LOT numbers in same order (comma-separated, leave blank if not applicable)")
            
            if skus_input and st.button("Generate D2C Labels"):
                skus = parse_input_list(skus_input)
                lots = parse_input_list(lots_input) if lots_input.strip() else [''] * len(skus)
                
                # Pad or truncate lots list to match SKUs length
                lots = (lots + [''] * len(skus))[:len(skus)]
                
                # Create DataFrame
                df = pd.DataFrame({
                    'SKU': skus,
                    'LOT#': lots
                })
                df['UPC Code'] = df['SKU'].apply(get_upc_for_sku)
                
                missing_upcs = df[df['UPC Code'].isna()]['SKU'].tolist()
                if missing_upcs:
                    st.error(f"UPC codes not found for SKUs: {', '.join(missing_upcs)}")
                else:
                    zip_path = generate_pdfs_from_excel(df)
                    if zip_path:
                        with open(zip_path, "rb") as f:
                            st.download_button("Download ZIP file with Labels", 
                                             f, 
                                             file_name=zip_path)
    
    elif option == "Generate FNSKU Labels":
        input_method = st.radio("Choose input method", 
                              ["Upload File", "Enter FNSKUs manually"],
                              key="fnsku_input_method")
        
        if input_method == "Upload File":
            uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], key="fnsku_uploader")
            if uploaded_file:
                try:
                    df = pd.read_excel(uploaded_file)
                    if 'FNSKU' not in df.columns:
                        st.error("Excel file must contain 'FNSKU' column")
                    else:
                        # Add Product Names from database
                        df['Product Name'] = df['FNSKU'].apply(get_product_name_for_fnsku)
                        missing_names = df[df['Product Name'].isna()]['FNSKU'].tolist()
                        if missing_names:
                            st.error(f"Product names not found for FNSKUs: {', '.join(missing_names)}")
                        elif st.button("Generate FNSKU Labels"):
                            zip_path = generate_fnsku_labels_from_excel(df)
                            if zip_path:
                                with open(zip_path, "rb") as f:
                                    st.download_button("Download ZIP file with FNSKU Labels", 
                                                     f, 
                                                     file_name=zip_path)
                except Exception as e:
                    st.error(f"Error processing file: {str(e)}")
        
        else:  # Manual FNSKU entry
            fnskus_input = st.text_area("Enter FNSKUs (comma-separated)")
            lots_input = st.text_area("Enter LOT numbers in same order (comma-separated, leave blank if not applicable)")
            
            if fnskus_input and st.button("Generate FNSKU Labels"):
                fnskus = parse_input_list(fnskus_input)
                lots = parse_input_list(lots_input) if lots_input.strip() else [''] * len(fnskus)
                
                # Pad or truncate lots list to match FNSKUs length
                lots = (lots + [''] * len(fnskus))[:len(fnskus)]
                
                # Create DataFrame
                df = pd.DataFrame({
                    'FNSKU': fnskus,
                    'LOT#': lots
                })
                df['Product Name'] = df['FNSKU'].apply(get_product_name_for_fnsku)
                
                missing_names = df[df['Product Name'].isna()]['FNSKU'].tolist()
                if missing_names:
                    st.error(f"Product names not found for FNSKUs: {', '.join(missing_names)}")
                else:
                    zip_path = generate_fnsku_labels_from_excel(df)
                    if zip_path:
                        with open(zip_path, "rb") as f:
                            st.download_button("Download ZIP file with FNSKU Labels", 
                                             f, 
                                             file_name=zip_path)

if __name__ == "__main__":
    main()
