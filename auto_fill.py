import openpyxl
import random

# Define the file path
file_path = "./Psikotest.xlsx"  # Update with the actual path
output_path = "./Psikotest_filled.xlsx"  # Update with the desired output path

# Load the Excel file
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Define the starting and ending row for column Q
start_row = 24
end_row = 292

# Iterate over the merged cell ranges in the sheet
for merged_range in sheet.merged_cells.ranges:
    # Get the top-left cell of the merged range
    top_left_cell = merged_range.start_cell
    
    # Check if the row is within the specified range for Q column
    if start_row <= top_left_cell.row <= end_row and top_left_cell.column == 17:
        # Assign random "A" or "B" to the top-left cell of the merged range
        top_left_cell.value = random.choice(["A", "B"])

# Save the modified workbook
workbook.save(output_path)
