import pandas as pd
import streamlit as st
import os

# Set page title
st.title("Excel File Splitter")

# File uploader
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

# Process the uploaded file
if uploaded_file is not None:
    # Read the Excel file
    df = pd.read_excel(uploaded_file, engine='openpyxl')

    # Get the base filename without extension
    base_filename = os.path.splitext(uploaded_file.name)[0]

    # Split the dataframe into chunks of 1999 rows
    chunk_size = 1999
    chunks = [df[i:i + chunk_size] for i in range(0, len(df), chunk_size)]

    # Save each chunk as a separate Excel file
    for idx, chunk in enumerate(chunks, start=1):
        output_filename = f"{base_filename}_part_{idx}.xlsx"
        chunk.to_excel(output_filename, index=False, engine='openpyxl')
        st.success(f"Generated file: {output_filename}")
        with open(output_filename, "rb") as f:
            st.download_button(label=f"Download {output_filename}", data=f, file_name=output_filename)

    st.info(f"Successfully split into {len(chunks)} files with up to {chunk_size} rows each.")
