Call validarReporteConWorkday(nameSheet, reporte)
Function validarReporteConWorkday(nameSheet As String, reporte As String) As Boolean
    Dim nameSheet As Worksheet
    Dim sheetReporte As Worksheet
    Dim ultimaFilaSheet As Long
    Dim ultimaFilaReporte As Long
    Dim i As Long
    Dim j As Long
    Dim encontrado As Boolean
    Dim employeeID As String

    Set nameSheet = ThisWorkbook.Sheets(nameSheetO)
    Set sheetReporte = ThisWorkbook.Sheets(reporte)
    ultimaFilaSheet = nameSheet.Cells(nameSheet.Rows.Count, "I").End(xlUp).Row
    ultimaFilaReporte = sheetReporte.Cells(sheetReporte.Rows.Count, "K").End(xlUp).Row
    For i = 2 To ultimaFilaSheet
        encontrado = False
        For j = 2 To ultimaFilaReporte
            If nameSheet.Cells(i, "I").Value = sheetReporte.Cells(j, "K").Value Then
                
                encontrado = True
                Exit For
            End If
        Next j
        If Not encontrado Then
            nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 22)).Interior.Color = RGB(255, 255, 0)
            nameSheet.Cells(i, "W").Value = "No existe en la hoja Reporte"
            validarReporteConWorkday = False
        Else
            validarReporteConWorkday = True
        End If
    Next i
End Function

For i = 2 To sheetLastRow
        Set sheetCell = sheetMacro.Cells(i, 9)
        Set found = reporteSheetRange.Find(What:=sheetCell.Value, LookIn:=xlValues, LookAt:=xlWhole)
        
        If found Is Nothing Then
            sheetMacro.Range(sheetMacro.Cells(i, 1), sheetMacro.Cells(i, 22)).Interior.Color = RGB(255, 255, 0)
            sheetMacro.Cells(i, "W").Value = "No aparece en el Reporte Joiner"
        End If
    Next i
        