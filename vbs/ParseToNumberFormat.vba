Sub NumberFormat(file_path As String, sheet_name As String, col_idx As Integer)
    ' Desactivar la actualización de pantalla y alertas
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Open file path
    Workbooks.Open file_path
    ' Select the sheet by the name
    Sheets(sheet_name).Select
    
    ' Apply format number to the column
    Columns(col_idx).NumberFormat = "0"
    
    ' Close and save the book modified
    ActiveWorkbook.Close SaveChanges:=True
    
    ' Reactivar la actualización de pantalla y alertas
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub