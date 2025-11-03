import openpyxl
from openpyxl import Workbook
import os

# Load the original Excel file
source_file = "LTL250002595.xlsx"
wb = openpyxl.load_workbook(source_file)
ws = wb.active

# Load the header from the copilot-generated file
copilot_file = "LTL250002595_part_3 (2) copilot.xlsx"
copilot_wb = openpyxl.load_workbook(copilot_file)
copilot_ws = copilot_wb.active

# Extract header values from the copilot file
header = [cell.value for cell in copilot_ws[1]]

# Determine the number of rows and chunk size
max_rows = 1999
total_rows = ws.max_row - 1  # excluding header
num_chunks = (total_rows + max_rows - 1) // max_rows

# Create output directory
output_dir = "split_correct_format"
os.makedirs(output_dir, exist_ok=True)

# Split and save each chunk
for i in range(num_chunks):
    start_row = i * max_rows + 2  # +2 to skip original header
    end_row = min(start_row + max_rows - 1, ws.max_row)

    new_wb = Workbook()
    new_ws = new_wb.active

    # Write header
    new_ws.append(header)

    # Write data rows
    for row in ws.iter_rows(min_row=start_row, max_row=end_row, values_only=True):
        new_ws.append(row)

    # Save the chunk
    output_file = os.path.join(output_dir, f"LTL250002595_part_{i+1}.xlsx")
    new_wb.save(output_file)

print(f"Successfully split into {num_chunks} files in '{output_dir}' directory.")
