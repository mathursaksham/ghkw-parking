import os
import re
import subprocess
import zipfile
from datetime import datetime
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

    # Ensure the Linux /tmp environment directory exists entirely
    os.makedirs("/tmp", exist_ok=True)

    doc = Document(TEMPLATE_FILE)

    # Automatically inject the current date into the data dictionary
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

    # Temporary system locations
    tmp_docx = f"/tmp/Letter_Flat_{flat_number}.docx"
    pdf_path = f"/tmp/Letter_Flat_{flat_number}.pdf"
    
    doc.save(tmp_docx)

    # Convert DOCX to PDF using LibreOffice headless command line
    try:
        libreoffice_cmd = "libreoffice"
        cmd = [
            libreoffice_cmd,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            "/tmp",
            tmp_docx,
        ]
        
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=True,
            env={"HOME": "/tmp"},
        )

        if not os.path.exists(pdf_path):
            st.error(f"LibreOffice execution completed but no PDF found for Flat {flat_number}.")
            return None

        # Read the generated PDF into memory
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

        # Clean up temporary storage files immediately
        if os.path.exists(tmp_docx):
            os.remove(tmp_docx)
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

        return pdf_bytes
        
    except Exception as e:
        st.error(f"Error converting Flat {flat_number} via LibreOffice: {e}")
        return None


# --- STREAMLIT FRONTEND ---
st.set_page_config(layout="wide")
st.title("🏢 GHKW Parking Allotments")

df = load_data()

if not df.empty:
    if "Flat Number" not in df.columns:
        st.error("Excel sheet must contain a 'Flat Number' column.")
    else:
        # Initialize session state to track checked flats persistently across searches
        if "selected_flats_tracker" not in st.session_state:
            st.session_state.selected_flats_tracker = set()

        col1, col2 = st.columns([3, 2])

        with col1:
            st.header("🔍 Search & Select Flats")
            
            # 1. Search filter input text box
            search_query = st.text_input(
                "Search by Flat / Block (e.g., 'C5', 'A', 'A1-003'):", 
                placeholder="Type to filter rows..."
            ).strip()

            # 2. Filter the dataframe based on search query
            if search_query:
                # Case-insensitive partial string matching on 'Flat Number'
                filtered_df = df[df["Flat Number"].str.contains(search_query, case=False, na=False)].copy()
            else:
                filtered_df = df.copy()

            # 3. Apply state back to the filtered dataframe column
            filtered_df.insert(
                0, 
                "Select", 
                filtered_df["Flat Number"].apply(lambda x: x in st.session_state.selected_flats_tracker)
            )

            # 4. Action buttons to select/deselect filtered rows at once
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                if st.button("✅ Check All Filtered Rows"):
                    for f_num in filtered_df["Flat Number"]:
                        st.session_state.selected_flats_tracker.add(f_num)
                    st.rerun()
            with btn_col2:
                if st.button("❌ Uncheck All Filtered Rows"):
                    for f_num in filtered_df["Flat Number"]:
                        st.session_state.selected_flats_tracker.discard(f_num)
                    st.rerun()

            # 5. Render interactive data table
            edited_df = st.data_editor(
                filtered_df,
                hide_index=True,
                column_config={
                    "Select": st.column_config.CheckboxColumn(
                        "Select",
                        help="Check to include this flat",
                        default=False,
                    )
                },
                disabled=[col for col in filtered_df.columns if col != "Select"],
                use_container_width=True,
                key="flat_data_editor"
            )

            # 6. Capture individual checklist edits from user interactions
            if edited_df is not None:
                for _, row in edited_df.iterrows():
                    f_num = row["Flat Number"]
                    if row["Select"]:
                        st.session_state.selected_flats_tracker.add(f_num)
                    else:
                        st.session_state.selected_flats_tracker.discard(f_num)

            # Get full data rows for all verified checked flats
            current_selections = list(st.session_state.selected_flats_tracker)
            selected_rows = df[df["Flat Number"].isin(current_selections)]

            st.write(f"📂 **Total unique flats selected overall:** {len(current_selections)}")
            if len(current_selections) > 0:
                with st.expander("See selected list"):
                    st.write(", ".join(sorted(current_selections)))

            # 7. File Compiler Trigger
            if st.button("Generate Letters for Chosen Rows"):
                if not current_selections:
                    st.warning("Please select or search-check at least one flat checkbox above.")
                
                # Case A: Exactly 1 checked -> Direct PDF download
                elif len(current_selections) == 1:
                    flat = current_selections[0]
                    row = selected_rows[selected_rows["Flat Number"] == flat].iloc[0].to_dict()
                    
                    with st.spinner(f"Processing PDF for Flat {flat}..."):
                        pdf_data = generate_pdf_bytes(row, flat)
                        if pdf_data:
                            st.success(f"PDF for Flat {flat} ready!")
                            st.download_button(
                                label="⬇️ Save PDF to Desktop",
                                data=pdf_data,
                                file_name=f"Letter_Flat_{flat}.pdf",
                                mime="application/pdf",
                            )
                
                # Case B: Multiple checked -> Package to ZIP file
                else:
                    with st.spinner("Processing chosen layout configurations..."):
                        zip_buffer = BytesIO()
                        progress_bar = st.progress(0)
                        total = len(current_selections)
                        processed_count = 0

                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for index, flat in enumerate(current_selections):
                                row = selected_rows[selected_rows["Flat Number"] == flat].iloc[0].to_dict()
                                pdf_data = generate_pdf_bytes(row, flat)
                                if pdf_data:
                                    zip_file.writestr(f"Letter_Flat_{flat}.pdf", pdf_data)
                                    processed_count += 1

                                progress_bar.progress((index + 1) / total)

                        if processed_count > 0:
                            st.success(f"Successfully packaged {processed_count} letters!")
                            st.download_button(
                                label="⬇️ Download Selected PDFs (ZIP)",
                                data=zip_buffer.getvalue(),
                                file_name="Selected_Parking_Letters.zip",
                                mime="application/zip",
                            )
                        else:
                            st.error("Failed to generate PDFs for the chosen flats.")

        with col2:
            st.header("🚀 Bulk Parking Print")
            st.write(
                "Compile all letters with assigned parking spaces into a ZIP file."
            )

            if st.button("Prepare All Parkings"):
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