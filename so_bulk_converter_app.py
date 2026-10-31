import streamlit as st
import pandas as pd
from pypdf import PdfReader
import io

st.set_page_config(page_title="Bulk Sales Order PDF to Excel Converter", layout="wide")

st.title("📦 Bulk Sales Order PDF to Excel Converter")
st.markdown("Upload multiple Sales Order PDFs below to automatically parse and combine them into a single Excel sheet.")

uploaded_files = st.file_uploader("Choose Sales Order PDF files", type=["pdf"], accept_multiple_files=True)

def extract_so_data(pdf_file):
    try:
        reader = PdfReader(pdf_file)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        
        # Basic placeholder extraction logic (customizable based on your exact PDF layout)
        # You can enhance this parsing logic to match your specific fields.
        return {
            "Filename": pdf_file.name,
            "RawTextSnippet": text[:300] + "..." # Preview of text
        }
    except Exception as e:
        return {
            "Filename": pdf_file.name,
            "RawTextSnippet": f"Error reading PDF: {e}"
        }

if uploaded_files:
    st.success(f"Loaded {len(uploaded_files)} PDF file(s).")
    
    if st.button("🚀 Process and Extract Data"):
        with st.spinner("Extracting data from PDFs..."):
            extracted_data = []
            for file in uploaded_files:
                data = extract_so_data(file)
                extracted_data.append(data)
            
            df = pd.DataFrame(extracted_data)
            st.dataframe(df)
            
            # Convert to Excel for download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sales_Orders')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Combined Excel Summary",
                data=processed_data,
                file_name="Sales_Order_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
