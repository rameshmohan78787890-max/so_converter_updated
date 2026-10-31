UPDATED BULK SO PDF TO EXCEL APP

Files:
- so_bulk_converter_app.py : updated Streamlit application
- requirements.txt         : required Python packages

GitHub / Streamlit Cloud update:
1. Replace your existing so_bulk_converter_app.py with this updated file.
2. Make sure requirements.txt is in the same repository root folder.
3. Commit and push both files to GitHub.
4. In Streamlit Community Cloud, open the app and click Reboot (or wait for redeploy).
5. Main file path: so_bulk_converter_app.py

Changes validated against the supplied sample SO PDF:
- Customer name taken from the SHIP TO block; strips adjacent PO-column text.
- Product name reads the continuation line (INDOCID OA 02).
- Packing size reads KG per package (180 KG), not drum capacity (210 LTR).
- Added Specification Type column, extracted from the Specifications field.
- Customer GST reads the GST NO field rather than the separate GST IN field.

Important: This was tested against the supplied sample PDF only. Test other SO layouts and review the extracted table before using the Excel output.
