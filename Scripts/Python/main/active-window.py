import pygetwindow as gw

# Look for windows that contain 'Excel' in their title
windows = [win for win in gw.getAllTitles() if 'Excel' in win]

if windows:
    # Get the first Excel window found
    excel_window = gw.getWindowsWithTitle(windows[0])[0]
    
    # Activate the window
    excel_window.activate()
    print(f"Excel window '{windows[0]}' activated.")
else:
    print("No Excel window found.")
