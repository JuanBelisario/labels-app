# Label Generation Tool

A Streamlit application for generating D2C and FNSKU labels with database integration.

## Features

### D2C Labels
- Generate labels using SKU and LOT# (optional)
- Automatic UPC lookup from database
- Support for both file upload and manual SKU entry
- Batch label generation with ZIP download

### FNSKU Labels
- Generate labels using FNSKU and LOT# (optional)
- Automatic Product Name lookup from database
- Support for both file upload and manual FNSKU entry
- Batch label generation with ZIP download

### Admin Features
- Database management interface
- Bulk update of SKU-UPC database
- Bulk update of FNSKU-Product Name database
- Secure admin authentication

## Setup
1. Install dependencies:
   bash
pip install -r requirements.txt

2. Configure Streamlit secrets with Supabase credentials

3. Run the application:
   bash
   streamlit run app_finalversion_labels.py
   
## Database Structure
- SKU Database: Stores SKU and UPC code pairs
- FNSKU Database: Stores FNSKU and Product Name pairs

## File Upload Formats
- Supported formats: CSV, Excel (.xlsx)
- Required columns for SKU database: 'SKU', 'UPC'
- Required columns for FNSKU database: 'FNSKU', 'Product Name'
