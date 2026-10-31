
import io
import re
from datetime import datetime
import pdfplumber
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Bulk SO PDF to Excel", page_icon="📄", layout="wide")

COLUMNS = [
    "SO Type", "SO Number", "SO Date", "Customer Name", "Customer GST",
    "Customer PO No.", "Delivery Location", "Marketing Person", "Contact Person",
    "Contact Phone", "Product Code", "HSN", "Product Name", "Unit",
    "Quantity", "Rate (INR)", "Net Amount (INR)", "Delivery Date",
    "Packing Details", "Drum Type", "Drum Quantity", "Packing Size (KG)", "Specification Type",
    "Transporter", "Freight", "Payment Terms", "Incoterms", "Tax",
    "Special Instruction", "Prepared By", "Approved By", "Source File",
    "Extraction Status"
]

def clean(s):
    return re.sub(r"\s+", " ", (s or "")).strip(" :\t\r\n")

def first(patterns, text, flags=re.I | re.S):
    for pat in patterns:
        m = re.search(pat, text, flags)
        if m:
            return clean(m.group(1))
    return ""

def number(value):
    if not value:
        return ""
    value = value.replace(",", "").strip()
    try:
        return float(value) if "." in value else int(value)
    except Exception:
        return value

def extract_pdf(uploaded_file):
    filename = uploaded_file.name
    with pdfplumber.open(io.BytesIO(uploaded_file.getvalue())) as pdf:
        pages = [(p.extract_text() or "") for p in pdf.pages]
    text = "\n".join(pages)
    flat = clean(text)

    so = first([r"\b(LSO|SO)\s*[-:]?\s*(\d{8,})"], text)
    # Separate SO type and number so it also works with LSO - 123...
    m = re.search(r"\b(LSO|SO)\s*[-:]?\s*(\d{8,})", text, re.I)
    so_type, so_number = (m.group(1).upper(), m.group(2)) if m else ("", "")

    # Customer name must come from the first non-empty line in the SHIP TO block.
    # Do not use BILL TO / supplier GST blocks as customer details.
    ship_block = re.search(r"SHIP\s+TO\s*:\s*\n?(.*?)(?=\n\s*(?:NOTIFY ADDRESS|ADDRESS FOR ORIGINAL DOCUMENTS|BILL\s+TO))", text, re.I | re.S)
    customer = ""
    if ship_block:
        for line in ship_block.group(1).splitlines():
            candidate = clean(line)
            if candidate and not re.search(r"^(?:Plot|B/H|Tal:|Contact Person|PH\s*:|FAX|MO\.?\s*:|EMAIL|GST\s*NO)", candidate, re.I):
                # The right-hand PO column can share the same PDF text line.
                candidate = re.split(r"\s+Purchase\s+Order\s+No", candidate, flags=re.I)[0]
                customer = clean(candidate)
                break

    # Product row: Sr No, item code, HSN/Description, Unit, Qty, Rate, Net Amount, Delivery Date
    product_code = first([r"\b( AST[A-Z0-9]{4,})\b", r"\b(AST[A-Z0-9]{4,})\b"], text)
    hsn = first([r"HSN\s*:\s*(\d{6,})"], text)
    # Product description can continue on the next line. Prefer the complete
    # continuation description when it is clearly part of the same item row.
    product_name = ""
    item_desc = re.search(r"^\s*1\s+AST[A-Z0-9]+\s+HSN\s*:\s*\d+\s*[-–]\s*([^\n]+)(?:\n\s*([^\n]+))?", text, re.I | re.M)
    if item_desc:
        first_desc = clean(item_desc.group(1))
        next_desc = clean(item_desc.group(2) or "")
        product_name = next_desc if next_desc and re.search(r"[A-Za-z]", next_desc) else first_desc
        if next_desc and next_desc.lower() not in product_name.lower() and len(next_desc) < 60:
            product_name = clean(product_name)
    quantity = first([r"\bKG\s+([\d,]+\.\d{2,})", r"\b([\d,]+\.\d{2,})\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2,}"], text)
    rate = first([r"\b([\d,]+\.\d{2})\s+[\d,]+\.\d{2,}\s+\d{2}/\d{2}/\d{4}", r"\bRate\s*:\s*([\d,]+(?:\.\d+)?)"], text)
    amount = first([r"Net Amount\s+Delivery Date.*?\n.*?([\d,]+\.\d{2})", r"\bTotal\s*:\s*INR\s*([\d,]+\.\d{2})"], text)
    # For standard line layout, find numeric columns on item line.
    item_line = re.search(r"^\s*1\s+(AST[A-Z0-9]+).*?\b(KG|LTR|MT)\b\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+\.\d{2})", text, re.I | re.M)
    if item_line:
        product_code = item_line.group(1)
        quantity, rate, amount = item_line.group(3), item_line.group(4), item_line.group(5)


    date_match = re.search(r"Date\s*:\s*(\d{2}/\d{2}/\d{4})", text, re.I)
    so_date = date_match.group(1) if date_match else ""
    delivery_date = first([r"Delivery Date\s*\n?\s*(\d{2}/\d{2}/\d{4})",
                           r"([\d]{2}/[\d]{2}/[\d]{4})\s*$"], text)
    po = first([r"Purchase Order No\./Dt\.\s*:\s*([^\n]+)"], text)
    delivery_location = first([r"Delivery Location\.\s*:\s*([^\n]+)"], text)
    mkt = first([r"MKT Person\.\s*:\s*([^\n]+)"], text)
    # Customer GST is explicitly labelled GST NO in the SHIP TO block.
    # GST IN elsewhere on this form can be Indo Amines' own GST number.
    gst = first([r"GST\s*NO\s*:\s*([A-Z0-9]{15})"], text)
    contact = first([r"Contact Person\s*:\s*([^\n]+)"], text)
    phone = first([r"(?:PH|MO)\s*:\s*([+\d][\d -]{8,})"], text)
    packing = first([r"PACKING DETAILS\s*:\s*([^\n]+)"], text)
    drum = first([r"DRUM TYPE\s*:\s*([^\n]+(?:\n\s+WITH[^\n]+)?)"], text)
    drum = clean(re.split(r"\s+PACKING\s+MATERIAL\b", drum, flags=re.I)[0])
    drum_qty = first([r"TOTAL NO OF\s*\n?\s*(\d+)"], text)
    # Product pack size is the KG amount per package (e.g. 180KG), not
    # the physical drum capacity (e.g. 210 LTR).
    pack_size = first([
        r"PACKING\s+GROUP\s+(\d+)\s*KG\b",
        r"Volume\s*:\s*[\d,]+\s*kgs?\s*\(\s*\d+\s*\*\s*(\d+)\s*kgs?"
    ], text)
    specification_type = first([r"Specifications\s*:\s*([^\n]+)"], text)
    transporter = first([r"Transporter\s*:\s*([^\n]+)"], text)
    freight = first([r"Incoterms\s*:\s*([^\n]+)", r"FREIGHT\s*:\s*([^\n]+)"], text)
    payment = first([r"Payment Terms\s*:\s*([^\n]+)"], text)
    incoterms = first([r"Incoterms\s*:\s*([^\n]+)"], text)
    tax = first([r"Tax Details\s*:\s*([^\n]+)"], text)
    special = first([r"Special Instructions\s*:\s*([^\n]+(?:\n(?!\s*(?:FUMIGATION|3RD PARTY|TERMS AND CONDITIONS)).*)?)"], text)
    prepared = first([r"Prepared By\s*:\s*\d+\s*\n?\s*Designation\s*:\s*([^\n]+)"], text)
    approved = first([r"Reviewed\s*&\s*Approved By\s*:\s*\d+\s*\n?\s*Designation\s*:\s*([^\n]+)"], text)

    # For some PDFs, customer GST appears beside the customer block; use ship-to GST as fallback.
    if not gst:
        gst = first([r"GST IN\s*:\s*([A-Z0-9]{15})"], text)

    data = {
        "SO Type": so_type, "SO Number": so_number, "SO Date": so_date,
        "Customer Name": customer, "Customer GST": gst, "Customer PO No.": po,
        "Delivery Location": delivery_location, "Marketing Person": mkt,
        "Contact Person": contact, "Contact Phone": phone,
        "Product Code": product_code, "HSN": hsn, "Product Name": product_name,
        "Unit": "KG" if re.search(r"\bKG\b", text, re.I) else "",
        "Quantity": number(quantity), "Rate (INR)": number(rate),
        "Net Amount (INR)": number(amount), "Delivery Date": delivery_date,
        "Packing Details": packing, "Drum Type": drum,
        "Drum Quantity": number(drum_qty), "Packing Size (KG)": number(pack_size),
        "Specification Type": specification_type,
        "Transporter": transporter, "Freight": freight, "Payment Terms": payment,
        "Incoterms": incoterms, "Tax": tax, "Special Instruction": special,
        "Prepared By": prepared, "Approved By": approved,
        "Source File": filename,
        "Extraction Status": "Text extracted" if len(text.strip()) > 80 else "No readable text - OCR required"
    }
    return data

st.title("📄 Bulk Sales Order PDF → Excel")
st.caption("Upload multiple Indo Amines Sales Order PDFs and create one combined Excel file. Review extracted rows before downloading.")

with st.sidebar:
    st.header("How it works")
    st.markdown("1. Select multiple PDF files.\\n2. Click **Extract SOs**.\\n3. Review/edit the table.\\n4. Download one Excel file.")
    st.info("Best for text-based PDFs. Scanned/image PDFs need OCR (Tesseract) and may require checking.")

files = st.file_uploader("Choose Sales Order PDF files", type=["pdf"], accept_multiple_files=True)

if "so_data" not in st.session_state:
    st.session_state.so_data = []

if st.button("Extract SOs", type="primary", disabled=not files):
    rows, errors = [], []
    progress = st.progress(0)
    for i, f in enumerate(files):
        try:
            rows.append(extract_pdf(f))
        except Exception as e:
            errors.append(f"{f.name}: {e}")
        progress.progress((i + 1) / len(files))
    st.session_state.so_data = rows
    if errors:
        st.error("\\n".join(errors))
    st.success(f"Processed {len(rows)} of {len(files)} PDF file(s).")

if st.session_state.so_data:
    st.subheader("Review / edit extracted data")
    df = pd.DataFrame(st.session_state.so_data, columns=COLUMNS)
    edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", key="editor")
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        edited.to_excel(writer, index=False, sheet_name="SO Summary")
        ws = writer.book["SO Summary"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        from openpyxl.styles import Font, PatternFill, Alignment
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[1].height = 32
        for col in ws.columns:
            letter = col[0].column_letter
            max_len = max((len(str(c.value or "")) for c in col[:100]), default=12)
            ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 45)
    st.download_button("⬇️ Download combined Excel", data=out.getvalue(),
                       file_name="Sales_Order_Summary.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    st.caption("Tip: Keep one row per SO. Check blank or unusual fields against the original PDF before using the summary.")
