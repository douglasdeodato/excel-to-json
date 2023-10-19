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

# Assuming the first row contains column headers
headers = [sheet.cell_value(0, col) for col in range(sheet.ncols)]

# Iterate through rows, skipping the first (header) row
for row in range(1, sheet.nrows):
    row_data = {}
    for col in range(sheet.ncols):
        if headers[col] == "Name":
            row_data["name"] = sheet.cell_value(row, col)
        elif headers[col] == "Email":
            row_data["email"] = sheet.cell_value(row, col)
    data.append(row_data)

# Write the JSON data to a file
# Replace 'output.json' with the desired output filename and path
with open('output.json', 'w') as json_file:
    json.dump(data, json_file, indent=4)

print("Excel file converted to JSON successfully.")
