' Definir la función que recibe un array de parámetros
Function EliminarFiltros(parametros)
    ' Desempaquetar los parámetros del array
    Dim file_path, sheetName
    file_path = parametros(0)
    sheetName = parametros(1)

    ' Crear un objeto Excel
    Set excelObj = CreateObject("Excel.Application")
    
    ' Abrir el archivo de Excel
    Set objWorkbook = excelObj.Workbooks.Open(file_path)
    ' MsgBox "Archivo abierto con éxito: " & file_path

    ' Seleccionar la hoja específica
    Set objWorksheet = objWorkbook.Sheets(sheetName)
    ' MsgBox "Hoja seleccionada: " & sheetName

    ' Verificar y eliminar los filtros
    If objWorksheet.AutoFilterMode Then 
        objWorksheet.AutoFilterMode = False
        ' MsgBox "Filtros eliminados en la hoja: " & sheetName
    ' Else
        ' MsgBox "No se encontraron filtros para eliminar en la hoja: " & sheetName
    End If

    ' Guardar los cambios y cerrar el archivo
    objWorkbook.Save
    objWorkbook.Close
    ' MsgBox "Archivo guardado y cerrado."

    ' Cerrar Excel
    excelObj.Quit

    ' Liberar los objetos
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    Set excelObj = Nothing

    ' Confirmar la ejecución completa
    ' aMsgBox "Script ejecutado con éxito."
End Function