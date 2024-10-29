# Updated function to handle SKUs with "/" for D2C label generation
def generate_pdfs_from_excel(df):
    required_columns = ['SKU', 'UPC Code', 'LOT#']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Missing columns in the Excel file: {', '.join(missing_columns)}")
        return None

    # Use a cleaned SKU value for filenames, but keep the original for display
    first_sku = df.iloc[0]['SKU'].replace("/", "_")
    current_date = datetime.now().strftime("%Y%m%d")

    output_folder = f"{first_sku}_{current_date}"
    os.makedirs(output_folder, exist_ok=True)

    total_rows = len(df)
    progress_bar = st.progress(0)

    for index, row in df.iterrows():
        sku = row['SKU'].replace("/", "_")  # Replace "/" in SKU for filenames and barcodes
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

# Updated function to generate D2C label PDF with cleaned SKU values
def generate_label_pdf(sku, upc_code, lot_num, output_path):
    width, height = 60 * mm, 35 * mm
    c = canvas.Canvas(output_path, pagesize=(width, height))

    x_margin = 4.5 * mm
    y_sku = height - 7.75 * mm
    y_barcode = height / 2 - 8 * mm
    y_lot = 4.75 * mm
    barcode_width = 51.5 * mm

    c.setFont("Helvetica", 9.5)
    display_sku = sku.replace("_", "/")  # Replace back to "/" for display if needed
    c.drawCentredString(width / 2, y_sku, display_sku)

    barcode_path = generate_d2c_barcode(upc_code, sku)  # Generate barcode with cleaned SKU
    if os.path.exists(barcode_path):
        c.drawImage(barcode_path, (width - barcode_width) / 2, y_barcode, width=barcode_width, height=16 * mm)
        os.remove(barcode_path)

    if lot_num:
        c.setFont("Helvetica", 9)
        c.drawCentredString(width / 2, 4.75 * mm, f"Lot: {lot_num}")
    c.save()
