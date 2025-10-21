import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Replace with your actual GitHub username and repo name
HEADER_URL = "https://raw.githubusercontent.com/yourusername/excel-split/main/excel%20header.xlsx"

st.title("Excel Splitter & Cleaner")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Load fixed header row from GitHub (first row only, no column names)
        header_df = pd.read_excel(HEADER_URL, header=None, dtype=str)
        fixed_header = header_df.iloc[[0]].copy()

        # Load uploaded file (skip its first row)
        df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)
        df = df_raw.iloc[1:].reset_index(drop=True)

        # Clean special characters
        df = df.applymap(lambda x: re.sub(r"[&'<]", '', x) if isinstance(x, str) else x)

        # Split into chunks
        chunk_size = 1999
        num_chunks = (len(df)) // (chunk_size - 1) + 1

        st.success(f"File loaded. Splitting into {num_chunks} parts...")

        for i in range(num_chunks):
            start = i * (chunk_size - 1)
            end = start + (chunk_size - 1)
            chunk = df.iloc[start:end]

            # Combine fixed header + chunk
            combined = pd.concat([fixed_header, chunk], ignore_index=True)

            # Save to Excel in memory
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                combined.to_excel(writer, index=False, header=False)
            buffer.seek(0)

            # Load workbook and left-align first row
            wb = load_workbook(buffer)
            ws = wb.active
            for cell in ws[1]:
                cell.alignment = Alignment(horizontal='left')

            # Save updated workbook
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            st.download_button(
                label=f"Download split_part_{i+1}.xlsx",
                data=buffer.getvalue(),
                file_name=f"split_part_{i+1}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
