import os
import re
import subprocess
import zipfile
from datetime import datetime  # <-- Added for current date
from io import BytesIO
import pandas as pd
import streamlit as st
from docx import Document

# Configuration
EXCEL_FILE = "data.xlsx"  # Path to your Excel file
TEMPLATE_FILE = "template.docx"  # Path to your Word template


@st.cache_data
def load_data():
    """Load and clean Excel data."""
    if not os.path.exists(EXCEL_FILE):
        st.error(f"Excel file '{EXCEL_FILE}' not found!")
        return pd.DataFrame()
    return pd.read_excel(EXCEL_FILE, dtype=str)


def generate_pdf_bytes(row_data, flat_number):
    """Populates Word template and converts to PDF using LibreOffice."""
    if not os.path.exists(TEMPLATE_FILE):
        st.error(f"Template file '{TEMPLATE_FILE}' not found!")
        return None

    doc = Document(TEMPLATE_FILE)

    # Automatically inject the current date into the data dictionary
    # Format 'B' is full month name, 'd' is day, 'Y' is 4-digit year
    row_data["Date"] = datetime.now().strftime("%B %d, %Y")

    # Replace placeholders in paragraphs
    for paragraph in doc.paragraphs:
        for key, value in row_data.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in paragraph.text:
                for run in paragraph.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))

    # Replace placeholders in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in row_data.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in paragraph.text:
                            paragraph.text = paragraph.text.replace(
                                placeholder, str(value)
                            )

    # Use Linux /tmp directory for processing
    tmp_docx = f"/tmp/Letter_Flat_{flat_number}.docx"
    doc.save(tmp_docx)

    # Convert DOCX to PDF using LibreOffice headless command line
    try:
        libreoffice_path = "/usr/bin/libreoffice"
        cmd = [
            libreoffice_path,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            "/tmp",
            tmp_docx,
        ]
        subprocess.run(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True,
            env={"HOME": "/tmp"},
        )

        pdf_path = f"/tmp/Letter_Flat_{flat_number}.pdf"

        # Read the generated PDF into memory
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

        # Clean up temporary files
        if os.path.exists(tmp_docx):
            os.remove(tmp_docx)
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

        return pdf_bytes
    except Exception as e:
        st.error(f"Error converting Flat {flat_number} via LibreOffice: {e}")
        return None


# --- STREAMLIT FRONTEND ---
st.title("🏢 GHKW Parking Allotments")

df = load_data()

if not df.empty:
    if "Flat Number" not in df.columns:
        st.error("Excel sheet must contain a 'Flat Number' column.")
    else:
        col1, col2 = st.columns(2)

        with col1:
            st.header("Single Flat Print")
            flat_list = df["Flat Number"].dropna().unique().tolist()
            selected_flat = st.selectbox("Select Flat Number:", flat_list)

            if st.button("Generate Letter"):
                row = df[df["Flat Number"] == selected_flat].iloc[0].to_dict()
                with st.spinner("Processing PDF..."):
                    pdf_data = generate_pdf_bytes(row, selected_flat)
                    if pdf_data:
                        st.success("PDF ready for download!")
                        st.download_button(
                            label="⬇️ Save PDF to Desktop",
                            data=pdf_data,
                            file_name=f"Letter_Flat_{selected_flat}.pdf",
                            mime="application/pdf",
                        )

        with col2:
            st.header("Bulk Parking Print")
            st.write(
                "Compile all letters with assigned parking spaces into a ZIP file."
            )

            if st.button("🚀 Prepare All Parkings"):
                if "Parking" in df.columns:
                    parking_df = df[
                        df["Parking"].notna() & (df["Parking"] != "")
                    ]
                else:
                    parking_df = df

                if parking_df.empty:
                    st.warning("No records found with parking details.")
                else:
                    zip_buffer = BytesIO()
                    progress_bar = st.progress(0)
                    total = len(parking_df)

                    with zipfile.ZipFile(
                        zip_buffer, "w", zipfile.ZIP_DEFLATED
                    ) as zip_file:
                        for index, (_, row) in enumerate(parking_df.iterrows()):
                            flat_num = row["Flat Number"]
                            pdf_data = generate_pdf_bytes(
                                row.to_dict(), flat_num
                            )

                            if pdf_data:
                                zip_file.writestr(
                                    f"Letter_Flat_{flat_num}.pdf", pdf_data
                                )

                            progress_bar.progress((index + 1) / total)

                    st.success("ZIP package generated successfully!")
                    st.download_button(
                        label="⬇️ Download All PDFs (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="All_Parking_Letters.zip",
                        mime="application/zip",
                    )