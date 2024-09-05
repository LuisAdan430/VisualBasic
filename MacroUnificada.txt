Sub Pruebas()
    Dim nameMacro As String
    Dim nameWorkday As String
    Dim nameReporte As String

    nameWorkday = "WORKDAY"
    nameReporte = "REPORTE"
    nameMacro = "REHIRE"
    crearHojaSheet(nameMacro)
    copiarTituloWorkday(nameWorkday, nameMacro)
    creacionComentarios(nameMacro)
    copiadoRehireA_Workday(nameWorkday, nameMacro)
    quitarDuplicados(nameMacro)
    validarReporteConWorkday(nameMacro,nameReporte)
End Sub

Function crearHojaSheet(nameMacro As String) As String
    Dim sheetDestino As Worksheet
    Dim sheetExist As Boolean
    sheetExist = False
    For Each sheetDestino In ThisWorkbook.Worksheets
        If sheetDestino.Name = nameMacro Then
            sheetExist = True
            Exit For
        End If
    Next sheetDestino
    If Not sheetExist Then
        Set sheetDestino = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        sheetDestino.Name = nameMacro
        sheetDestino.Tab.Color = RGB(0, 0, 255)
    Else
        Set sheetDestino = ThisWorkbook.Sheets(nameMacro)
        sheetDestino.Cells.Clear
    End If
    crearHojaSheet = "Se ha creado exitosamente la hoja"
End Function

Function copiarTituloWorkday(nameWorkday As String, nameMacro As String) As String
    Dim sheetWorkday As Worksheet
    Dim sheetDestino As Worksheet

    Set sheetWorkday = ThisWorkbook.Sheets(nameWorkday)
    Set sheetDestino = ThisWorkbook.Sheets(nameMacro)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    sheetWorkday.Rows(1).Copy Destination:=sheetDestino.Rows(1)

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    copiarTituloWorkday = "Se copio el titulo correctamente del Workday"

End Function
Function creacionComentarios(nameMacro As String) As String
    Dim sheetDestino As Worksheet
    Set sheetDestino = ThisWorkbook.Sheets(nameMacro)
    sheetDestino.Columns("W:W").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    sheetDestino.Range("W1").Value = "Comentarios"
    creacionComentarios = "Se crearon correctamente los comentarios"
End Function


Function copiadoRehireA_Workday(nameWorkday As String, nameMacro As String) As String
    Dim sheetWorkday As Worksheet
    Dim sheetDestino As Worksheet
    Dim ultimaFilaWorkday As Long
    Dim a As Long
    Dim filaDestinoSheet As Long

    Set sheetWorkday = ThisWorkbook.Sheets(nameWorkday)
    Set sheetDestino = ThisWorkbook.Sheets(nameMacro)
    ultimaFilaWorkday = sheetWorkday.Cells(sheetWorkday.Rows.Count, "G").End(xlUp).Row
    filaDestinoSheet = 2
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For a = 2 To ultimaFilaWorkday
        If sheetWorkday.Cells(a, "G").Value = "A" And sheetWorkday.Cells(a, "A").Value <> "" Then
            sheetWorkday.Rows(a).Copy Destination:=sheetDestino.Rows(filaDestinoSheet)
            filaDestinoSheet = filaDestinoSheet + 1
        End If
    Next a

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    copiadoRehireA_Workday = "Se hizo el copiado correctamente"
End Function

Function quitarDuplicados(nameMacro As String) As String 
    
    Dim sheetDuplicate As Worksheet
    Dim lastRowSheetDuplicate As Long
    Dim rng As Range
    Set sheetDuplicate = ThisWorkbook.Sheets(nameMacro)
    lastRowSheetDuplicate = sheetDuplicate.Cells(sheetDuplicate.Rows.Count, "V").End(xlUp).Row
    Set rng = sheetDuplicate.Range("A1:V" & lastRowSheetDuplicate)
    rng.RemoveDuplicates Columns:=22, Header:=xlYes
    quitarDuplicados = "Se quito el duplicado correctamente"
    
End Function

Function validarReporteConWorkday(nameMacro As String, nameReporte As String) As String
    Dim nameSheet As Worksheet
    Dim sheetReporte As Worksheet
    Dim ultimaFilaSheet As Long
    Dim ultimaFilaReporte As Long
    Dim i As Long
    Dim j As Long
    Dim encontrado As Boolean
    Dim employeeID As String
    
    Set nameSheet = ThisWorkbook.Sheets(nameMacro)
    Set sheetReporte = ThisWorkbook.Sheets(nameReporte)
    ultimaFilaSheet = nameSheet.Cells(nameSheet.Rows.Count, "I").End(xlUp).Row
    ultimaFilaReporte = sheetReporte.Cells(sheetReporte.Rows.Count,"K").End(xlUp).Row

    For i = 2 To ultimaFilaSheet
        encontrado = False
        For j = 2 To ultimaFilaReporte
            If nameSheet.Cells(i,"I").Value = sheetReporte.Cells(j,"K").Value Then
                encontrado = True
                Exit For
            End If
        Next j
        If Not encontrado Then
            nameSheet.Range(nameSheet.Cells(i,1), nameSheet.Cells(i,23)).Interior.Color = RGB(255,255,0)
            nameSheet.Cells(i,"W").Value = "No se encuentra el usuario en el reporte pero si en Workday"
        End If
    Next i 

validarReporteConWorkday = "Se hizo la validacion con reporte y workday correctamente"
End Function

