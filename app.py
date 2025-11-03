import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Replace with your actual GitHub username and repo name
HEADER_URL = "https://raw.githubusercontent.com/ruslanaskraucenka/excel-split/main/excel%20header.xlsx"
# URL to your fixed header file in GitHub
HEADER_URL = "https://raw.githubusercontent.com/yourusername/excel-split/main/excel%20header.xlsx"

st.title("Excel Splitter & Cleaner")


        # Load fixed header row from GitHub (first row only, no column names)
        header_df = pd.read_excel(HEADER_URL, header=None, dtype=str)
        fixed_header = header_df.iloc[[0]].copy()
        fixed_header_row = header_df.iloc[0].tolist()  # Extract as list

        # Load uploaded file (skip its first row)
        df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)
            end = start + (chunk_size - 1)
            chunk = df.iloc[start:end]

            # Combine fixed header + chunk
            combined = pd.concat([fixed_header, chunk], ignore_index=True)
            # Prepend fixed header row as first row of data
            chunk_with_header = pd.DataFrame([fixed_header_row], columns=None)
            chunk_combined = pd.concat([chunk_with_header, chunk], ignore_index=True)

            # Save to Excel in memory
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                combined.to_excel(writer, index=False, header=False)
                chunk_combined.to_excel(writer, index=False, header=False)
            buffer.seek(0)

            # Load workbook and left-align first row
