import streamlit as st
import pandas as pd
import os

st.title("Excel Splitter: Max 1999 Rows per File")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    max_rows = 1999
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    num_chunks = (len(df) + max_rows - 1) // max_rows

    st.success(f"File uploaded successfully! Splitting into {num_chunks} files...")

    for i in range(num_chunks):
        start = i * max_rows
        end = start + max_rows
        chunk = df.iloc[start:end]
        output_filename = f"{uploaded_file.name.replace('.xlsx', '')}_part_{i+1}.xlsx"
        chunk.to_excel(output_filename, index=False)
        with open(output_filename, "rb") as f:
            st.download_button(
                label=f"Download {output_filename}",
                data=f,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        os.remove(output_filename)
