Sub validar()
    Dim nameSheet As String
    Dim nameWorkday As String
    Dim reporte As String

    reporte = "reporte"
    nameWorkday = "Workday"
    nameSheet = "MOVER"

    Call crearHojaSheet(nameSheet)
    Call copiarTituloWorkday(nameWorkday, nameSheet)
    Call valoresWorkday(nameSheet, nameWorkday)
    Call quitarDuplicados(nameSheet)
    Call creacionComentarios(nameSheet)
    Call validarReporteConWorkday(nameSheet, reporte)
    Call OrdenarPorColor
    Call CopiarValores
    MsgBox "Ha terminado la macro"

End Sub
Function validarAplicaciones(aplicaciones As String) As Boolean
    Dim aplicacionesConversion As Variant
    Dim aplicacionesManager As Variant
    aplicacionesConversion = Split(aplicaciones, ",")
    aplicacionesManager = Array("Workday", "LDAP ALHAMBRA MEXICO", "Azure Active Directory", "GIM", "AD - Corporativo", "AD - Produban", "AD - Sucursales", "AD - Contact Center", "AD - Altec")
    Dim i As Long
    Dim j As Long
    Dim valorDiferente As Boolean
    For i = LBound(aplicacionesConversion) To UBound(aplicacionesConversion)
        valorDiferente = False
        For j = LBound(aplicacionesManager) To UBound(aplicacionesManager)
            If aplicacionesConversion(i) = aplicacionesManager(j) Then
                valorDiferente = True
                Exit For
            End If
        Next j
        If Not valorDiferente Then
            validarAplicaciones = False
            Exit Function
        End If
    Next i
    validarAplicaciones = True

End Function


Function validarTipoMovimiento(j As Long, reporte As String, i As Long, nameSheetO As String) As Boolean
    Dim sheetReporte As Worksheet
    Dim nameSheet As Worksheet
    Dim tipoMovimiento As String
    Dim aplicaciones As String

    Set sheetReporte = ThisWorkbook.Sheets(reporte)
    Set nameSheet = ThisWorkbook.Sheets(nameSheetO)
    tipoMovimiento = sheetReporte.Cells(j, "M").Value
    aplicaciones = sheetReporte.Cells(j, "F").Value
    If tipoMovimiento <> "C" Then
        'validarTipoMovimiento = True
        If validarAplicaciones(aplicaciones) Then
           validarTipoMovimiento = False
           nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.color = RGB(255, 255, 255) ' White
           nameSheet.Cells(i, "W").Value = "eventos sin certificación"
        Else
          nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.color = RGB(255, 255, 0) ' Yellow
          nameSheet.Cells(i, "W").Value = "certificación enviada a Manager"
          validarTipoMovimiento = True
        End If
    Else
        validarTipoMovimiento = False
        nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.color = RGB(0, 255, 255) ' Cyan
        nameSheet.Cells(i, "W").Value = "No detono el evento"
    End If


End Function


Function validarNombreCertificacion(j As Long, reporte As String, i As Long, nameSheetO As String) As Boolean
    Dim sheetReporte As Worksheet
    Dim nameSheet As Worksheet
    Dim nombreCertificacion As String
    Dim buscar As String
    Set sheetReporte = ThisWorkbook.Sheets(reporte)
    Set nameSheet = ThisWorkbook.Sheets(nameSheetO)
    nombreCertificacion = sheetReporte.Cells(j, "P").Value
    buscar = "Mover Event Certification"
    If InStr(1, nombreCertificacion, buscar, vbTextCompare) > 0 Then
        validarNombreCertificacion = True
    Else
        validarNombreCertificacion = False
        nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.color = RGB(205, 92, 92)  ' Indian Red
        nameSheet.Cells(i, "W").Value = "Movimiento Invalido"
    End If

End Function



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
                If validarNombreCertificacion(j, reporte, i, nameSheetO) Then
                    If validarTipoMovimiento(j, reporte, i, nameSheetO) Then

                    End If
                End If
            
            
            
            encontrado = True
                Exit For
            End If
        Next j
        If Not encontrado Then
            nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.color = RGB(204, 153, 255)  ' Msuve is equals light purple
            nameSheet.Cells(i, "W").Value = "No existe en la hoja Reporte"
        End If
    Next i
End Sub
Sub creacionComentarios(nameSheet As String)
    Dim sheetDestino As Worksheet
    Set sheetDestino = ThisWorkbook.Sheets(nameSheet)
    sheetDestino.Columns("W:W").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    sheetDestino.Range("W1").Value = "Comentario"
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
       sheetDestino.Tab.color = RGB(0, 0, 255)
    Else
        Set sheetDestino = ThisWorkbook.Sheets(nameSheet)
        sheetDestino.Cells.Clear
    End If

    
End Sub


Sub OrdenarPorColor()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim rng As Range
    Dim celda As Range
    Dim colorDict As Object
    Dim i As Integer
    Dim j As Integer
    Dim lastRow As Long
    Dim tempRow As Long
    Dim colorOrden As Variant
    Dim totalCols As Long
    Dim colorHex As Long
    Dim orderNumber As Long
    Dim auxColExists As Boolean

    Set ws = ThisWorkbook.Sheets("MOVER")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set rng = ws.Range("A2:A" & lastRow)
    colorOrden = Array(RGB(205, 92, 92), RGB(255, 255, 0), RGB(255, 255, 255), RGB(0, 255, 255), RGB(204, 153, 255))

    Set colorDict = CreateObject("Scripting.Dictionary")

    For i = LBound(colorOrden) To UBound(colorOrden)
        colorDict.Add colorOrden(i), i + 1
    Next i

    helperCol = "Z"
    auxColExists = Not IsEmpty(ws.Range(helperCol & "1").Value)
    
    If Not auxColExists Then
        totalCols = ws.Cells(1, helperCol).Column
        
    Else
        totalCols = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    End If

    For Each celda In rng
        colorHex = celda.Interior.color
        If colorDict.Exists(colorHex) Then
            orderNumber = colorDict(colorHex)
        Else
            orderNumber = colorDict.Count + 1
        End If
        ws.Cells(celda.Row, helperCol).Value = orderNumber
    Next celda

    Dim sortRange As Range
    Set sortRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, totalCols))
    sortRange.Sort Key1:=ws.Range(helperCol & "2"), Order1:=xlAscending, Header:=xlNo

    ws.Columns(helperCol).ClearContents
    
    Exit Sub
ErrorHandler:
       ' MsgBox "An error ocurred: " & Err.Description, vbExclamation
End Sub


Sub CopiarValores()
    Dim wsReporte As Worksheet
    Dim wsMover As Worksheet
    Dim lastRowReporte As Long
    Dim lastRowMover As Long
    Dim i As Long
    Dim j As Long
    Dim empleadoIDReporte As Variant
    Dim empleadoIDMover As Variant
    Dim sheetWork As String

    sheetWork = "MOVER"
    Set wsReporte = ThisWorkbook.Sheets("Reporte")
    Set wsMover = ThisWorkbook.Sheets(sheetWork)
    
    Call agregarTitulos(sheetWork)

    lastRowReporte = wsReporte.Cells(wsReporte.Rows.Count, "K").End(xlUp).Row
    lastRowMover = wsMover.Cells(wsMover.Rows.Count, "I").End(xlUp).Row
    
    For i = 2 To lastRowReporte
        empleadoIDReporte = wsReporte.Cells(i, "K").Value
        For j = 2 To lastRowMover
            empleadoIDMover = wsMover.Cells(j, "I").Value
            If empleadoIDReporte = empleadoIDMover Then
                wsMover.Cells(j, "Y").Value = wsReporte.Cells(i, "F").Value
                wsMover.Cells(j, "Z").Value = wsReporte.Cells(i, "K").Value
                wsMover.Cells(j, "AA").Value = wsReporte.Cells(i, "N").Value
                wsMover.Cells(j, "AB").Value = wsReporte.Cells(i, "M").Value
                Exit For
            End If
        Next j
    Next i
    'MsgBox "Valores copiados exitosamente."
End Sub

Sub agregarTitulos(sheetWork As String)
    Dim wsMover As Worksheet
    Set wsMover = ThisWorkbook.Sheets(sheetWork)
    wsMover.Cells(1, "Y").Value = "APLICACIONES"
    wsMover.Range("Y1").Interior.Color = RGB(217,217,217) '

    wsMover.Cells(1, "Z").Value = "EMPLOYEE ID"
    wsMover.Range("Z1").Interior.Color = RGB(217,217,217)

    wsMover.Cells(1, "AA").Value = "NOMBRE DE USUARIO"
    wsMover.Range("AA1").Interior.Color = RGB(217,217,217)

    wsMover.Cells(1, "AB").Value = "TIPO DE MOVIMIENTO"
    wsMover.Range("AB1").Interior.Color = RGB(217,217,217)


End Sub
