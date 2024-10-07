import win32com.client as win32

# Connect to an already running instance of Excel
excel = win32.Dispatch("Excel.Application")

# Access the active workbook (the one that's already open)
workbook = excel.ActiveWorkbook

# Save the workbook
workbook.Save()

# Close the workbook
workbook.Close(SaveChanges=False)  # Use SaveChanges=True if you want to save changes before closing

# Quit Excel application
excel.Quit()