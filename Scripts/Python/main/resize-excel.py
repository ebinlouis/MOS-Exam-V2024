import pygetwindow as gw
import tkinter as tk

def resizeWindow():
    # Get all windows with 'Excel' in the title
    windows = gw.getWindowsWithTitle('Excel')

    # Debug: Print out all window titles
    for i, win in enumerate(windows):
        print(f"Window {i}: {win.title}")

    root = tk.Tk()
    root.withdraw()
    height = root.winfo_screenheight() * 0.72
    heightForExcelExam = int(height)
    root.destroy()

    if len(windows) > 0:
        # Assuming there's only one relevant Excel window, or manually choose one based on titles
        excel_window = windows[0]  # You may need to change the index based on the printed titles

        # Set custom height
        custom_height = heightForExcelExam
        excel_window.resizeTo(excel_window.width, custom_height)
        
        print("Excel window height set to:", custom_height)
    else:
        print("Excel window not found.")

resizeWindow()

