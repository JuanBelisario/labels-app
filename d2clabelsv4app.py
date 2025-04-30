# ✅ FIXED VERSION: PL Builder module added, label logic preserved untouched from original version.

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
import difflib  # 🔁 ADD THIS NEW IMPORT
import urllib.parse

def normalize_column_names(df):
    """
    Identifies and renames variations of 'Destination SKU' to a standard column name.
    Accepts: 'destination sku', 'DestinationSKU', 'dest_sku', 'destinationsku', etc.
    """
    target_normalized = "destinationsku"

    for col in df.columns:
        col_normalized = col.strip().lower().replace(" ", "").replace("_", "")
        if target_normalized in col_normalized:
            df.rename(columns={col: "Destination SKU"}, inplace=True)
            return df

    # Fallback: fuzzy match if substring wasn't enough
    import difflib
    normalized_map = {col: col.strip().lower().replace(" ", "").replace("_", "") for col in df.columns}
    match = difflib.get_close_matches(target_normalized, normalized_map.values(), n=1, cutoff=0.75)

    if match:
        for original_col, normalized in normalized_map.items():
            if normalized == match[0]:
                df.rename(columns={original_col: "Destination SKU"}, inplace=True)
                break

    return df



# --- ORIGINAL LABEL LOGIC STARTS HERE (UNCHANGED) ---

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
    font_size = 9
    c.setFont("Helvetica", font_size)
    if product_name:
        wrap_text_to_two_lines(product_name, max_length=22, c=c, start_x=5 * mm, start_y=7.75 * mm, line_height=font_size - 1.5, max_width=25)
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
        barcode_image = generate_fnsku_barcode(fnsku)
        create_fnsku_pdf(barcode_image, fnsku, product_name, lot, output_folder)
        progress_bar.progress((index + 1) / total_rows)
    zip_filename = f"{output_folder}.zip"
    with ZipFile(zip_filename, 'w') as zipObj:
        for folder_name, subfolders, filenames in os.walk(output_folder):
            for filename in filenames:
                if "_fnsku_label" in filename:
                    filepath = os.path.join(folder_name, filename)
                    zipObj.write(filepath, os.path.basename(filepath))
    return zip_filename

# --- PL BUILDER STARTS HERE ---
# Optional: clean up inconsistent location names for standardization
LOCATION_MAP = {
    "JD NJ : JD NJ - AMZ FBA": "JD NJ - AMZ FBA",
    "JD Canada : JD Canada - AMZ FBA": "JD CANADA - AMZ FBA",
    "JD UK : JD UK - AMAZON FBA": "JD UK - AMZ FBA",
    "JD AU : JD AU - AMAZON FBA": "JD AU - AMZ FBA",
    "JD - Belk": "JD CA - BELK",
    "JD - Showcase": "JD CANADA - SHOWCASE",
    "JD Canada": "JD CANADA",
    "JD CA - Walmart": "JD CA - WALMART",
    "JD - Nordstrom.com": "JD CA - NORDSTROM.COM",
    "JD - Nordstrom Stores": "JD CA - NORDSTROM STORES",
    "JD CA - Macy's": "JD CA - MACY'S",
    "JD NJ - Macy's": "JD CA - MACY'S",
    "JD ATL - Macy's": "JD CA - MACY'S",
    "Kenz - SALASA": "SALAZA SA - KENZ",
    "Kenz - SALASA": "SALAZA SA",
    "JD Canada - Walmart": "JD CANADA - WALMART",
    "Lateral TJ": "LATERAL TJ"
    


    
    # You can keep expanding this list as needed
}

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

    # === FILE NAMING LOGIC ===
    to = df['TO'].iloc[0]
    so = df['FOP SO #'].iloc[0]
    from_loc = df['From Loc'].iloc[0]
    to_loc = df['To Loc'].iloc[0]

    total_qty = None

    # Try to extract from 'Total QTY' if there's a 'Total' row
    if 'Total QTY' in df.columns:
        total_row = df[df['TO'].astype(str).str.lower().str.strip() == 'total']
        if total_row.empty and 'Trafilea SKU' in df.columns:
            total_row = df[df['Trafilea SKU'].astype(str).str.lower().str.strip() == 'total']

        if not total_row.empty:
            qty_val = total_row.iloc[0].get('Total QTY')
            if pd.notna(qty_val) and float(qty_val) > 0:
                total_qty = int(float(qty_val))

    # Fallback if Total QTY is missing or invalid
    if total_qty is None:
        filtered_df = df[~df['TO'].astype(str).str.lower().str.strip().eq('total')]
        total_qty = int(pd.to_numeric(filtered_df['Required Qty'], errors='coerce').sum())

    filename = f"{to} + {so} + {from_loc} + {to_loc} + {total_qty} Units.xlsx"

    # === OUTPUT COLUMNS ===
    headers = [
        "TO", "SO #", "From Loc", "To Loc", "Trafilea SKU", "Required Qty",
        "Shipping Method", "FG", "LOT", "Expiration Date", "CARTONS",
        "UNITS/Ctn", "Total QTY", "Carton Dimensions(inch) ", "Carton WEIGHT-LB",
        "Pallet Dimensions", "Pallet WEIGHT-LB.", "Pallet #"
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

    if transformation and 'Destination SKU' in df.columns:
        output_df['Destination SKU'] = df['Destination SKU']
    if 'Total QTY' in df.columns:
        output_df['Total QTY'] = df['Total QTY']

    # === EXCEL EXPORT ===
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        output_df.to_excel(writer, index=False, sheet_name='PL')
        workbook = writer.book
        worksheet = writer.sheets['PL']

        dark_blue = workbook.add_format({
            'bold': True, 'bg_color': '#0C2D63', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        light_blue = workbook.add_format({
            'bold': True, 'bg_color': '#D9EAF7', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })

        for col_num, col_name in enumerate(output_df.columns):
            header_format = dark_blue if col_name in [
                "TO", "SO #", "From Loc", "To Loc", "Trafilea SKU", "Destination SKU", "Required Qty", "Shipping Method"
            ] else light_blue
            worksheet.write(0, col_num, col_name, header_format)
            worksheet.set_column(col_num, col_num, 22)

    output.seek(0)
    return output, filename

# --- STREAMLIT APP UI ---

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
                df = pd.read_excel(uploaded_file, engine='openpyxl' if uploaded_file.name.endswith('xlsx') else 'xlrd')
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
                df = pd.read_excel(uploaded_file, engine='openpyxl' if uploaded_file.name.endswith('xlsx') else 'xlrd')
                if st.button("Generate FNSKU Labels", key="generate_fnsku_labels"):
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
        st.success(f"{len(uploaded_files)} file(s) uploaded successfully.")
        st.markdown("### 📝 Processed Packing Lists")

        for uploaded_file in uploaded_files:
            try:
                if uploaded_file.name.endswith(".csv"):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(
                        uploaded_file,
                        engine='openpyxl' if uploaded_file.name.endswith('xlsx') else 'xlrd'
                    )

                # Normalize column names
                df.columns = [col.strip() for col in df.columns]
                is_transformation = any("destination sku" in col.lower() for col in df.columns)
                output, filename = build_pl_base(df, transformation=is_transformation)

                if output:
                    # --- Extract values for prefill ---
                    raw_to = df['TO'].iloc[0]
                    raw_so = df['FOP SO #'].iloc[0]
                    raw_from_loc = df['From Loc'].iloc[0]
                    raw_to_loc = df['To Loc'].iloc[0]
                    raw_shipping = df['Shipping Method'].iloc[0]

                    from_loc = LOCATION_MAP.get(str(raw_from_loc).strip(), str(raw_from_loc).strip())
                    to_loc = LOCATION_MAP.get(str(raw_to_loc).strip(), str(raw_to_loc).strip())

                    filtered_df = df[~df['TO'].astype(str).str.lower().str.strip().eq('total')]
                    qty = int(pd.to_numeric(filtered_df['Required Qty'], errors='coerce').sum())
                    sku_count = filtered_df['SKU External ID'].nunique()

                    def enc(val): return urllib.parse.quote_plus(str(val))
                    form_link = (
                        "https://docs.google.com/forms/d/e/1FAIpQLSelQ08zk5O1py2t5czsuW4jnpVYO22LAtMskBxlbk__WuRgmA/viewform"
                        f"?entry.811040286={enc(raw_to)}"
                        f"&entry.771037158={enc(raw_so)}"
                        f"&entry.75050938={qty}"
                        f"&entry.2087058692={sku_count}"
                        f"&entry.227202689={enc(from_loc)}"
                        f"&entry.855389021={enc(to_loc)}"
                        f"&entry.105986750={enc(raw_shipping)}"
                    )

                    with st.container():
                        st.markdown(
                            f"<p style='margin-bottom: 0.25em;'><strong>📄 {filename}</strong></p>",
                            unsafe_allow_html=True
                        )

                        col1, col2 = st.columns([1, 1])
                        with col1:
                            st.markdown(
                                f"""
                                <a href="{form_link}" target="_blank" style="text-decoration: none;">
                                    <button style='
                                        padding: 0.4em 1em;
                                        font-size: 14px;
                                        border: 1px solid #999;
                                        border-radius: 6px;
                                        background-color: #f4f4f4;
                                        color: #000;
                                        width: 100%;
                                    '>📝 Fill TO Template</button>
                                </a>
                                """,
                                unsafe_allow_html=True
                            )
                        with col2:
                            st.download_button(
                                label="⬇️ Download PL Excel",
                                data=output,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=filename,
                                use_container_width=True
                            )

            except Exception as e:
                st.error(f"❌ Error processing file '{uploaded_file.name}': {e}")

    st.markdown(
        """
        <br>
        <a href="https://docs.google.com/forms/d/e/1FAIpQLSelQ08zk5O1py2t5czsuW4jnpVYO22LAtMskBxlbk__WuRgmA/viewform" target="_blank">
            <button style='padding: 0.5em 1em; font-size: 14px;'>📧 Fill TO Template | Send Email</button>
        </a>
        """,
        unsafe_allow_html=True
    )
