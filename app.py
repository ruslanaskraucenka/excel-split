import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import io

st.title("Excel Splitter with Header and Formatting Preservation")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

MAX_ROWS = 1999

def adjust_column_widths(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    header = df.columns.tolist()
    num_chunks = (len(df) + MAX_ROWS - 1) // MAX_ROWS

    st.success(f"File uploaded successfully! Splitting into {num_chunks} files...")

    for i in range(num_chunks):
        start = i * MAX_ROWS
        end = start + MAX_ROWS
        chunk = df.iloc[start:end]

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        # Write header
        ws.append(header)

        # Write data rows
        for row in dataframe_to_rows(chunk, index=False, header=False):
            ws.append(row)

        # Adjust column widths
        adjust_column_widths(ws)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        filename = f"{uploaded_file.name.replace('.xlsx', '')}_part_{i+1}.xlsx"
        st.download_button(
            label=f"Download {filename}",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
