import streamlit as st
import pandas as pd
from pypdf import PdfReader
import io
import re
import concurrent.futures

st.set_page_config(page_title="Bulk Sales Order PDF to Excel Converter", layout="wide")

st.title("⚡ High-Performance Bulk SO to Excel Converter")
st.markdown("Upload multiple Sales Order PDFs below. The app processes them concurrently and extracts structured data instantly.")

uploaded_files = st.file_uploader("Choose Sales Order PDF files", type=["pdf"], accept_multiple_files=True)

def extract_fields_from_text(text, filename):
    """Smart regex extraction for Indo Amines Sales Orders"""
    
    # Helper function to find patterns safely
    def find_pattern(pattern, default=""):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else default

    # Extracting common SO fields using flexible patterns
    so_number = find_pattern(r'(?:SO\s*No\.?|Sales\s*Order\s*No\.?|LSO)[^\d]*(\d{8,12})', filename.split('.')[0])
    so_date = find_pattern(r'(\d{2}/\d{2}/\d{4})')
    customer = find_pattern(r'(?:Customer|Consignee)[^\w]*(?:Name)?[^\w]*([A-Z\s.,]+(?:PRIVATE LIMITED|LTD|CORP|COMPANY))')
    product_code = find_pattern(r'(ASTFG\d+|ASTPF\d+)')
    
    # Quantity, Rate, Amount extraction fallback heuristics
    qty = find_pattern(r'(\d{1,3}(?:,\d{3})*\s*(?:KG|MT|DRUMS))')
    rate = find_pattern(r'(?:Rate|Price)[^\d]*([\d,]+\.?\d*)')
    amount = find_pattern(r'(?:Total|Amount)[^\d]*([\d,]+\.?\d*)')

    return {
        "Source File": filename,
        "SO Number": so_number,
        "SO Date": so_date,
        "Customer Name": customer,
        "Product Code": product_code,
        "Extracted Quantity": qty,
        "Rate": rate,
        "Total Amount": amount,
        "Raw Preview": text[:200].replace('\n', ' ')
    }

def process_single_pdf(file):
    try:
        reader = PdfReader(file)
        text = "".join([page.extract_text() + "\n" for page in reader.pages])
        return extract_fields_from_text(text, file.name)
    except Exception as e:
        return {
            "Source File": file.name,
            "SO Number": "Error",
            "SO Date": "",
            "Customer Name": f"Failed to parse: {e}",
            "Product Code": "",
            "Extracted Quantity": "",
            "Rate": "",
            "Total Amount": "",
            "Raw Preview": ""
        }

if uploaded_files:
    st.success(f"Loaded {len(uploaded_files)} PDF file(s).")
    
    if st.button("🚀 Process All PDFs Concurrently"):
        with st.spinner("Processing PDFs in parallel..."):
            # Multi-threading for maximum speed
            with concurrent.futures.ThreadPoolExecutor() as executor:
                extracted_data = list(executor.map(process_single_pdf, uploaded_files))
            
            df = pd.DataFrame(extracted_data)
            st.dataframe(df, use_container_width=True)
            
            # Excel export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sales_Orders_Summary')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Combined Excel Summary",
                data=processed_data,
                file_name="Sales_Order_Summary_Pro.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
