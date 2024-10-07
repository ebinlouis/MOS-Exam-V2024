import openpyxl

# Load the Excel file
workbook = openpyxl.load_workbook('../Files/Excel/Project 1.1.xlsx')

# Select the active sheet (or you can select a specific sheet by name)
sheet = workbook['Ship Mode']  # Or workbook['SheetName']

# Get the value of cell A1
value = sheet['AK99'].value

print(value)
