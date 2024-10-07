import win32com.client
import win32gui
import win32con
import os

def alwaysOnTop(enable=True):
    def set_window_always_on_top(window_title, enable):
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32gui.SetWindowPos(
                hwnd,
                win32con.HWND_TOPMOST if enable else win32con.HWND_NOTOPMOST,
                0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
            )
        else:
            print(f"Window with title '{window_title}' not found.")

    try:
        # Attach to the existing instance of Excel
        excel = win32com.client.GetObject(None, "Excel.Application")
    except Exception as e:
        print("No running instance of Excel found.")
        exit()

    # Get the window title of the active Excel window
    full_title = excel.ActiveWindow.Caption

    # Extract the filename without the extension
    filename = os.path.splitext(full_title)[0]

    print(filename)

    # Append " - Excel" to the filename
    window_title = filename + " - Excel"

    # Set the Excel window always on top or not based on the 'enable' parameter
    set_window_always_on_top(window_title, enable)

# To enable always on top
alwaysOnTop(enable=False)
