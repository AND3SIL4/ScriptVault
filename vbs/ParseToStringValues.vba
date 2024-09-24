Sub ConvertirColumnaAString(rutaArchivo As String, nombreHoja As String, indiceColumna As Integer)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim ultimaFila As Long
    
    ' Abre el archivo especificado
    Set wb = Workbooks.Open(rutaArchivo)
    
    ' Establece la hoja especificada
    Set ws = wb.Sheets(nombreHoja)
    
    ' Encuentra la última fila con datos en la columna especificada
    ultimaFila = ws.Cells(ws.Rows.Count, indiceColumna).End(xlUp).Row
    
    ' Establece el rango de la columna
    Set rng = ws.Range(ws.Cells(1, indiceColumna), ws.Cells(ultimaFila, indiceColumna))
    
    ' Convierte cada celda en el rango a tipo string
    For Each cell In rng
        cell.Value = CStr(cell.Value)
    Next cell
    
    ' Guarda y cierra el archivo
    wb.Close SaveChanges:=True
End Sub