Sub FiltrarPorFecha(rutaLibro As String, nombreHoja As String, nombreColumna As String, fechaInicio As Date, fechaFin As Date)
    ' Desactivar la actualización de pantalla y alertas
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Abrir el libro especificado
    Workbooks.Open rutaLibro
    
    ' Seleccionar la hoja especificada
    Sheets(nombreHoja).Select
    
    ' Encontrar el número de columna basado en el nombre de la columna
    Dim colNum As Integer
    colNum = Sheets(nombreHoja).Rows(1).Find(What:=nombreColumna, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    ' Aplicar el filtro
    ActiveSheet.Range("$A$1:$DH$9029").AutoFilter Field:=colNum, Criteria1:=">=" & fechaInicio, Operator:=xlAnd, Criteria2:="<=" & fechaFin
    
    ' Desplazar la ventana si es necesario
    ActiveWindow.SmallScroll Down:=9028
    
    ' Reactivar la actualización de pantalla y alertas
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
