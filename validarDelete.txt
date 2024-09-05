Sub ValidarDelete()
    Dim nameSheet As String
    Dim nameWorkday As String
    Dim reporte As String

    reporte = "reporte"
    nameWorkday = "Workday"
    nameSheet = "Delete"

    Call crearHojaSheet(nameSheet)
    Call copiarTituloWorkday(nameWorkday, nameSheet)
    Call valoresWorkday(nameSheet, nameWorkday)
    Call quitarDuplicados(nameSheet)
    Call creacionComentarios(nameSheet)
    Call validarReporteConWorkday(nameSheet, reporte)
    Call Concatenar()
    MsgBox "Ha terminado la macro"

End Sub

Function validarResultado(j As Long, reporte As String, i As Long, nameSheetO As String) As Boolean
'Resultado columna E
Dim nameSheet As Worksheet
Dim sheetReporte As Worksheet
Dim resultado As String
Set nameSheet = ThisWorkbook.Sheets(nameSheetO)
Set sheetReporte = ThisWorkbook.Sheets(reporte)
resultado = sheetReporte.Cells(j,"E").Value
If resultado = "Correcto" Then
    validarResultado = True
Else
    validarResultado = False
    nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 22)).Interior.Color = RGB(205, 92, 92)
    nameSheet.Cells(i, "W").Value = "Resultado Incorrecto"
End If


End Function

Function obtenerFechaActual() As String
    Dim fechaActual As Date
    Dim fechaFormato As String
    fechaActual = Date
    fechaFormato = Format(fechaActual, "DD/MM/YYYY")
    obtenerFechaActual = fechaFormato
End Function



Function validarAccountDeletionDate(j As Long, reporte As String, i As Long, nameSheetO As String) As Boolean
'Account Deletion Date : Columna P 
Dim sheetReporte As Worksheet
Dim accountDeletionDate As String
Dim fechaActuala As String
Dim nameSheet As Worksheet

Set nameSheet = ThisWorkbook.Sheets(nameSheetO)
Set sheetReporte = ThisWorkbook.Sheets(reporte)
accountDeletionDate = sheetReporte.Cells(j,"P").Value
fechaActual = obtenerFechaActual()
If fechaActual = accountDeletionDate Then
    validarAccountDeletionDate = True
    nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 22)).Interior.Color = RGB(80, 200, 120)
Else 
    validarAccountDeletionDate = False
    
    nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 22)).Interior.Color = RGB(255, 165, 0)
    nameSheet.Cells(i, "W").Value = "El Accoun Deletion Date esta Incorrecto"
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
    ultimaFilaReporte = sheetReporte.Cells(sheetReporte.Rows.Count, "H").End(xlUp).Row
    For i = 2 To ultimaFilaSheet
        encontrado = False
        For j = 2 To ultimaFilaReporte
            If nameSheet.Cells(i, "I").Value = sheetReporte.Cells(j, "H").Value Then
                If validarResultado(j,reporte, i, nameSheetO) Then
                    If validarAccountDeletionDate(j,reporte, i, nameSheetO) Then
                    End If
                End If
            encontrado = True
                Exit For
            End If
        Next j
        If Not encontrado Then
            nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 22)).Interior.Color = RGB(255, 255, 0)
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
        If sheetWorkday.Cells(a, "G").Value = "B" Then
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
       sheetDestino.Tab.Color = RGB(0, 0, 255)
    Else
        Set sheetDestino = ThisWorkbook.Sheets(nameSheet)
        sheetDestino.Cells.Clear
    End If

    
End Sub


Sub Concatenar()
    Dim wsReporte As Worksheet
    Dim wsDelete As Worksheet
    Dim lastRowReporte As Long
    Dim lastRowDelete As Long
    Dim i As Long
    Dim j As Long
    Dim empleadoIDReporte As Variant
    Dim empleadoIDDelete As Variant
    Dim empID As String
    Set wsReporte = ThisWorkbook.Sheets("Reporte")
    Set wsDelete = ThisWorkbook.Sheets("DELETE")
    
    lastRowReporte = wsReporte.Cells(wsReporte.Rows.Count, "A").End(xlUp).Row
    lastRowDelete = wsDelete.Cells(wsDelete.Rows.Count, "A").End(xlUp).Row
    
    wsDelete.Cells(1,"Z").Value = "EmpID"
    wsDelete.Range("Z1").Interior.Color = RGB(217,217,217)
    For j = 2 To lastRowDelete
        If j <> lastRowDelete Then
            empID = "empId = """ & wsDelete.Cells(j, "I").Value & """ || "
        Else
            empID = "empId = """ & wsDelete.Cells(j, "I").Value & """  "
        End If
        wsDelete.Cells(j, "Z").Value = empID
    Next j

    Call concatenacionFinal()
End Sub

Sub concatenacionFinal()
    Dim wsDelete As Worksheet
    Dim lastRow As Long
    Dim resultado As String
    Dim cell As Range

    Set wsDelete = ThisWorkbook.Sheets("DELETE")
    wsDelete.Cells(1,"AA").Value = "Concatenacion Final"
    wsDelete.Range("AA1").Interior.Color = RGB(217,217,217)

    lastRow = wsDelete.Cells(wsDelete.Rows.Count, "Z").End(xlUp).Row
    resultado = ""
    For Each cell In wsDelete.Range("Z2:Z" & lastRow)
        If cell.Value <> "" Then
            resultado = resultado & cell.Value & " "
        End If 
    Next cell
    wsDelete.Range("AA2").Value = "( " & Trim(resultado) & " ) "
    
End Sub





