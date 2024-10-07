import xlwings as xw

# Connect to existing Excel instance
app = xw.apps.active

# Get the active workbook
workbook = app.books.active

# Define your VBA code
vba_code = """
Sub P2T5()

Dim ws As Worksheet

On Error GoTo ErrorHandler

Set ws = ThisWorkbook.Sheets("Sales Person")

If ws.Range("B2").Formula = "=VLOOKUP(A2,$K$3:$L$11,2,FALSE)" And _
    ws.Range("B8").Formula = "=VLOOKUP(A8,$K$3:$L$11,2,FALSE)" Or _
        ws.Range("B2").Formula = "=VLOOKUP(A2,$K$2:$L$11,2,FALSE)" And _
            ws.Range("B8").Formula = "=VLOOKUP(A8,$K$2:$L$11,2,FALSE)" Then
    ws.Range("AK99").Value = "Correct"
    Exit Sub
Else
    ws.Range("AK99").Value = "Wrong"
End If

ErrorHandler:
    ws.Range("AK99").Value = "Wrong"

End Sub
"""

# Add the VBA code to the workbook
workbook_vba_project = workbook.api.VBProject
workbook_vba_module = workbook_vba_project.VBComponents.Add(1)
workbook_vba_module.CodeModule.AddFromString(vba_code)

# Run the macro
app.api.Run("P2T5")

# Delete the VBA module
workbook_vba_project.VBComponents.Remove(workbook_vba_module)
