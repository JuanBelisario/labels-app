# Patch the PL Builder to conditionally exclude "Destination SKU" from Normal TO PL

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

    # Define headers based on PL type
    headers = [
        "TO", "SO #", "From Loc", "To Loc", "Trafilea SKU", "Required Qty", "Shipping Method",
        "FG", "LOT", "Expiration Date", "CARTONS",
        "UNITS/Ctn", "Total QTY", "Carton Dimensions(inch) ", "Carton WEIGHT-LB",
        "Pallet Dimensions", "Pallet WEIGHT-LB.", "Pallet #"
    ]
    if transformation:
        headers.insert(5, "Destination SKU")  # Insert after "Trafilea SKU"

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
