import xlrd
import json

# Replace 'your_excel_file.xls' with the actual filename and path
excel_file = '1.xls'

# Open the Excel file
workbook = xlrd.open_workbook(excel_file)

# Choose a specific sheet if needed
sheet = workbook.sheet_by_index(0)  # 0 represents the first sheet

# Create a JSON object to store the data
data = []

# Iterate through rows and columns, and convert to JSON
for row in range(sheet.nrows):
    row_data = {}
    for col in range(sheet.ncols):
        cell_value = sheet.cell_value(row, col)
        row_data[sheet.cell_value(0, col)] = cell_value  # Assuming the first row contains column headers
    data.append(row_data)

# Write the JSON data to a file
# Replace 'output.json' with the desired output filename and path
with open('output.json', 'w') as json_file:
    json.dump(data, json_file, indent=4)

print("Excel file converted to JSON successfully.")
