import xlwings as xw

# Connect to existing Excel instance
app = xw.apps.active

# Get the active workbook
workbook = app.books.active

# Define your VBA code
vba_code = """
Sub P2T1()

Dim ws As Worksheet

On Error GoTo ErrorHandler

Set ws = ThisWorkbook.Sheets("Ship Mode")

If ws.Range("M2").Formula = "=MAXIFS($E$2:$E$40,$D$2:$D$40,H2,$C$2:$C$40,I2)" And _
    ws.Range("M5").Formula = "=MAXIFS($E$2:$E$40,$D$2:$D$40,H5,$C$2:$C$40,I5)" Or _
        ws.Range("M2").Formula = "=MAXIFS($E$2:$E$40,$C$2:$C$40,I2,$D$2:$D$40,H2)" And _
            ws.Range("M5").Formula = "=MAXIFS($E$2:$E$40,$C$2:$C$40,I5,$D$2:$D$40,H5)" Then
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
app.api.Run("P2T1")

# Delete the VBA module
workbook_vba_project.VBComponents.Remove(workbook_vba_module)
