import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import Alignment
from openpyxl import load_workbook

st.title("Excel Splitter & Cleaner")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Load the Excel file as strings to preserve formatting
        df = pd.read_excel(uploaded_file, dtype=str)

        # Clean special characters from all string cells
        df = df.applymap(lambda x: re.sub(r"[&'<]", '', x) if isinstance(x, str) else x)

        chunk_size = 1999
        num_chunks = (len(df) - 1) // (chunk_size - 1) + 1

        st.success(f"File loaded. Splitting into {num_chunks} parts...")

        for i in range(num_chunks):
            start = i * (chunk_size - 1)
            end = start + (chunk_size - 1)
            chunk = df.iloc[start:end]

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                chunk.to_excel(writer, index=False, header=True)
            buffer.seek(0)

            # Load workbook and apply left alignment to header row
            wb = load_workbook(buffer)
            ws = wb.active
            for cell in ws[1]:
                cell.alignment = Alignment(horizontal='left')

            # Save updated workbook back to buffer
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
