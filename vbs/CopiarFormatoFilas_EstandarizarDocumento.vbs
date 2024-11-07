Sub CopiarFormatoFila2(Libro As String, Hoja As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim rangoDestino As Range

    ' Desactivar actualizaciones en pantalla y alertas
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Abrir el libro especificado y asignar la hoja
    On Error Resume Next
    Set wb = Workbooks(Libro)
    If wb Is Nothing Then
        Set wb = Workbooks.Open(Libro, UpdateLinks:=0)
    End If
    On Error GoTo 0

    ' Verificar si el libro se abrió correctamente
    If wb Is Nothing Then
        MsgBox "No se pudo abrir el libro especificado.", vbExclamation
        GoTo LimpiarYSalir
    End If

    ' Asignar la hoja especificada
    On Error Resume Next
    Set ws = wb.Sheets(Hoja)
    On Error GoTo 0

    ' Verificar si la hoja existe
    If ws Is Nothing Then
        MsgBox "La hoja especificada no existe en el libro.", vbExclamation
        GoTo LimpiarYSalir
    End If

    ' Encontrar la última fila con datos en la hoja
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Definir el rango donde se aplicará el formato (todas las filas hasta la última con datos, en la columna AM)
    Set rangoDestino = ws.Range("A3:AM" & ultimaFila)
    
    ' Copiar solo el formato de la fila 2 y aplicarlo en el rango de destino
    ws.Rows(2).Copy
    rangoDestino.PasteSpecial Paste:=xlPasteFormats

    ' Limpiar la selección de copiado
    Application.CutCopyMode = False

LimpiarYSalir:
    ' Restaurar actualizaciones en pantalla y alertas
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Cerrar el libro si fue abierto en esta macro
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=True
    End If

    ' Liberar objetos
    Set ws = Nothing
    Set wb = Nothing
    Set rangoDestino = Nothing
End Sub