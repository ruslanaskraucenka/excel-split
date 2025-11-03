import pandas as pd
import os

# Configuration
input_file = "LTL250002595.xlsx"  # Replace with your actual file path
output_dir = "split_files"
max_rows = 1999

# Create output directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

# Load the Excel file
df = pd.read_excel(input_file, engine="openpyxl")

# Calculate the number of chunks
num_chunks = (len(df) + max_rows - 1) // max_rows

# Split and save each chunk
for i in range(num_chunks):
    start = i * max_rows
    end = start + max_rows
    chunk = df.iloc[start:end]
    output_file = os.path.join(output_dir, f"{os.path.splitext(input_file)[0]}_part_{i+1}.xlsx")
    chunk.to_excel(output_file, index=False)

print(f"Split into {num_chunks} files in '{output_dir}' directory.")
``
