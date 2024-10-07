import openpyxl

# Load the Excel file
workbook = openpyxl.load_workbook('src/Excel/Project 1.2.xlsx')

# Select the active sheet (or you can select a specific sheet by name)
sheet = workbook['Sales Person']  # Or workbook['SheetName']

# Get the value of cell A1
value = sheet['AK99'].value

print(value)
