import xlwings as xw
import os

def openExcel():
    try:
        # Specify the path to your Excel file
        file_path = '../Files/Excel/Project 1.1.xlsx'
        
        # Check if the file exists
        if not os.path.isfile(file_path):
            # Get the current working directory
            current_dir = os.getcwd()
            raise FileNotFoundError(f"The file {file_path} does not exist. Current working directory: {current_dir}.")
        
        # Open the Excel file
        wb = xw.Book(file_path)
        
        # Make the Excel window full screen
        app = wb.app
        app.api.WindowState = xw.constants.WindowState.xlMaximized
    
    except FileNotFoundError as fnf_error:
        print(fnf_error)
    except Exception as e:
        print(f"An error occurred: {e}")

openExcel()
