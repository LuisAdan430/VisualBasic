Sub MacroHrMover()
    Dim nameSheet As String
    Dim reporte As String
    Dim nameWorkday As String

    nameSheet = "HrMover"
    reporte = "Reporte"
    nameWorkday = "Workday"

    Call crearHojaSheet(nameSheet)
    Call copiarTituloWorkday(nameWorkday, nameSheet)
    Call creacionComentarios(nameSheet)
    Call valoresWorkday(nameSheet, nameWorkday)
    Call quitarDuplicados(nameSheet)
    Call validarReporteConWorkday(nameSheet, reporte)
End Sub





Sub crearHojaSheet(nameSheet As String)
    Dim sheetDestino As Worksheet
    Dim sheetExist As Boolean

    sheetExist = False
    For Each sheetDestino In ThisWorkbook.Worksheets
        If sheetDestino.Name = nameSheet Then
            sheetExist = True
            Exit For
        End If
    Next sheetDestino

    If Not sheetExist Then
       Set sheetDestino = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
       sheetDestino.Name = nameSheet
       sheetDestino.Tab.Color = RGB(0, 0, 255)
    Else
        Set sheetDestino = ThisWorkbook.Sheets(nameSheet)
        sheetDestino.Cells.Clear
    End If
End Sub
Sub copiarTituloWorkday(nameWorkday As String, nameSheet As String)
    Dim sheetWorkday As Worksheet
    Dim sheetDestino As Worksheet
    Set sheetWorkday = ThisWorkbook.Sheets(nameWorkday)
    Set sheetDestino = ThisWorkbook.Sheets(nameSheet)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    sheetWorkday.Rows(1).Copy Destination:=sheetDestino.Rows(1)
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
Sub creacionComentarios(nameSheet As String)
    Dim sheetDestino As Worksheet
    Set sheetDestino = ThisWorkbook.Sheets(nameSheet)
    sheetDestino.Columns("W:W").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    sheetDestino.Range("W1").Value = "Comentario"
End Sub
Sub valoresWorkday(nameSheet As String, nameWorkday As String)
    Dim sheetWorkday As Worksheet
    Dim sheetDestino As Worksheet
    Dim ultimaFilaWorkday As Long
    Dim a As Long
    Dim filaDestinoSheet As Long

    Set sheetWorkday = ThisWorkbook.Sheets(nameWorkday)
    Set sheetDestino = ThisWorkbook.Sheets(nameSheet)
    ultimaFilaWorkday = sheetWorkday.Cells(sheetWorkday.Rows.Count, "G").End(xlUp).Row
    filaDestinoSheet = 2
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For a = 2 To ultimaFilaWorkday
        If sheetWorkday.Cells(a, "G").Value = "C" Then
            sheetWorkday.Rows(a).Copy Destination:=sheetDestino.Rows(filaDestinoSheet)
            filaDestinoSheet = filaDestinoSheet + 1
        End If
    Next a
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub quitarDuplicados(nameSheet As String)
    Dim sheetDuplicate As Worksheet
    Dim lastRowSheetDuplicate As Long
    Dim rng As Range
    Set sheetDuplicate = ThisWorkbook.Sheets(nameSheet)
    lastRowSheetDuplicate = sheetDuplicate.Cells(sheetDuplicate.Rows.Count, "V").End(xlUp).Row
    Set rng = sheetDuplicate.Range("A1:V" & lastRowSheetDuplicate)
    rng.RemoveDuplicates Columns:=22, Header:=xlYes
End Sub

Sub validarReporteConWorkday(nameSheetO As String, reporte As String)
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
                    If validarResultado(nameSheetO, reporte, i, j) Then
                        If validarDominio(reporte, j, nameSheetO, i) Then
                             nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.Color = RGB(0, 255, 0)  ' Green
                             
                        End If
                    End If
                Exit For
            End If
        Next j
        If Not encontrado Then
            nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.Color = RGB(255, 255, 0)  ' Yellow
            nameSheet.Cells(i, "W").Value = "Validar tipo de movimiento ( Mover o HrMover)"
        
        End If
    Next i
End Sub

Function validarResultado(nameSheetO As String, reporte As String, i As Long, j As Long) As Boolean
    Dim nameSheet As Worksheet
    Dim sheetReporte As Worksheet
    Dim resultadoReporte As String

    Set nameSheet = ThisWorkbook.Sheets(nameSheetO)
    Set sheetReporte = ThisWorkbook.Sheets(reporte)

    resultadoReporte = sheetReporte.Cells(j, "E").Value
    If resultadoReporte <> "Correcto" Then
        validarResultado = False
        nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.Color = RGB(0, 0, 255)  ' Blue
        nameSheet.Cells(i, "W").Value = "Evento incorrecto. "
    Else
        validarResultado = True
    End If

End Function

Function validarDominio(reporte As String, j As Long, nameSheetO As String, i As Long) As Boolean
    Dim sheetReporte As Worksheet
    Dim nombreUsuario As String
    Dim letraNombreUsuario As String
    Dim domainName As String

    Set nameSheet = ThisWorkbook.Sheets(nameSheetO)
    Set sheetReporte = ThisWorkbook.Sheets(reporte)
    ' L : nombre de usuario : Reporte
    domainName = sheetReporte.Cells(j, "J").Value
    nombreUsuario = sheetReporte.Cells(j, "L").Value
    letraNombreUsuario = extraerPrimeraLetra(nombreUsuario)
    ' MsgBox nombreUsuario & " " & letraNombreUsuario
    If validarDomainName(reporte, j, nameSheetO, i) Then
        If letraNombreUsuario = "S" And domainName = "SucurSales" Then
            validarDominio = True
        ElseIf letraNombreUsuario <> "S" And domainName <> "SucurSales" Then
            validarDominio = True
        ElseIf letraNombreUsuario = "S" And domainName <> "SucurSales" Then
            nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.Color = RGB(255, 0, 0)  ' Red
            nameSheet.Cells(i, "W").Value = "Expediente incorrecto"
            validarDominio = False
        ElseIf letraNombreUsuario <> "S" And domainName = "SucurSales" Then
            nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.Color = RGB(255, 0, 0)  ' Red
            nameSheet.Cells(i, "W").Value = "Expediente incorrecto"
            validarDominio = False
        End If
    End If
End Function

Function extraerPrimeraLetra(nombreUsuario As String) As String
    Dim e As Long
    Dim letrasNombreUsuario As Variant
    Dim letraNombreUsuario As String
    
    ReDim tempArray(1 To Len(nombreUsuario))
        For e = 1 To Len(nombreUsuario)
            tempArray(e) = Mid(nombreUsuario, e, 1)
        Next e

    letrasNombreUsuario = tempArray
    letraNombreUsuario = letrasNombreUsuario(1)
    extraerPrimeraLetra = letraNombreUsuario
End Function

Function validarDomainName(reporte As String, j As Long, nameSheetO As String, i As Long) As Boolean
    'Domain Name: J: Reporte
    Dim sheetReporte As Worksheet
    Dim domainName As String
    Dim nameSheet As Worksheet

    Set nameSheet = ThisWorkbook.Sheets(nameSheetO)
    Set sheetReporte = ThisWorkbook.Sheets(reporte)
    domainName = sheetReporte.Cells(j, "J").Value

    If domainName = "No Domain" Or domainName = "" Then
        nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.Color = RGB(255, 0, 0)  ' Red
        nameSheet.Cells(i, "W").Value = "El dominio esta vacio o tiene el valor Domain Name"
        validarDomainName = False
    Else
        validarDomainName = True
    End If
End Function
