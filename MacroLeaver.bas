Sub validarLeaver()
    Dim nameSheet As String
    Dim nameWorkday As String
    Dim reporte As String
    Dim nameSantanderTerminaciones As String

    nameSheet = "Leaver"
    reporte = "reporte"
    nameWorkday = "Workday"
    nameSantanderTerminaciones = "SantanderTerminaciones"

    Call crearHojaSheet(nameSheet)
    Call copiarTituloWorkday(nameWorkday, nameSheet)
    Call creacionComentarios(nameSheet)
    Call valoresWorkday(nameSheet, nameWorkday)
    Call quitarDuplicados(nameSheet)
    Call validarReporteConWorkday(nameSheet, reporte, nameSantanderTerminaciones)
End Sub
Sub validarReporteConWorkday(nameSheetO As String, reporte As String, nameSantanderTerminaciones As String)
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
    ultimaFilaSheet = nameSheet.Cells(nameSheet.Rows.Count, "V").End(xlUp).Row
    ultimaFilaReporte = sheetReporte.Cells(sheetReporte.Rows.Count, "K").End(xlUp).Row
    For i = 2 To ultimaFilaSheet
        encontrado = False
        For j = 2 To ultimaFilaReporte
            If nameSheet.Cells(i, "V").Value = sheetReporte.Cells(j, "K").Value Then
                encontrado = True
                   Call validarLeaverConSantanderTerminaciones(nameSantanderTerminaciones, nameSheetO, i, reporte, j)
                   Exit For
            End If
        Next j
        If Not encontrado Then
            nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.Color = RGB(255, 0, 0)  ' Red
            nameSheet.Cells(i, "W").Value = "No se lanzo el evento"
        
        End If
    Next i
End Sub

Sub validarLeaverConSantanderTerminaciones(nameSantanderTerminaciones As String, nameSheetO As String, i As Long, reporte As String, j As Long)
    Dim sheetSantanderTerminaciones As Worksheet
    Dim nameSheet As Worksheet
    Dim ultimaFilaSantaderTerminacion As Long
    Dim z As Long
    Dim encontrado As Boolean

    Set nameSheet = ThisWorkbook.Sheets(nameSheetO)
    Set sheetSantanderTerminaciones = ThisWorkbook.Sheets(nameSantanderTerminaciones)
    ultimaFilaSantaderTerminacion = sheetSantanderTerminaciones.Cells(sheetSantanderTerminaciones.Rows.Count, "D").End(xlUp).Row

    
        encontrado = False
        For z = 2 To ultimaFilaSantaderTerminacion
            If nameSheet.Cells(i, "V").Value = sheetSantanderTerminaciones.Cells(z, "D").Value Then
                encontrado = True
                Call validarResultado(nameSheetO, reporte, i, j)
                
            End If
        Next z
        If Not encontrado Then
            nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.Color = RGB(0, 0, 255)  ' Yellow
            nameSheet.Cells(i, "W").Value = "No esta en el informe santantader terminaciones "
       
        End If
End Sub

Sub validarResultado(nameSheetO As String, reporte As String, i As Long, j As Long)
    Dim nameSheet As Worksheet
    Dim sheetReporte As Worksheet
    Dim resultadoReporte As String

    Set nameSheet = ThisWorkbook.Sheets(nameSheetO)
    Set sheetReporte = ThisWorkbook.Sheets(reporte)

    resultadoReporte = sheetReporte.Cells(j, "E").Value
    If resultadoReporte <> "Correcto" Then
        nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.Color = RGB(255, 255, 0)  ' Yellow
        nameSheet.Cells(i, "W").Value = "Evento Incorrecto "
    Else
        nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 23)).Interior.Color = RGB(0, 255, 0)  ' Blue
        nameSheet.Cells(i, "W").Value = "Evento Correcto."
    End If

End Sub
' Casos de uso:
'Tipos de movimiento B
'Quitar duplicados
'Comparar el usuario V de workday con K de reporte, Comentario:no se lanzo el evento, Color :Rojo=Resultado
'Resultado comparar con fila santander terminaciones comparar con la letra D , comentario: ,No esta en el informe santantader terminaciones Color: Azul
'La columna de reporte (resultado), si es incorrecto colocar de color amarillo, comentario: Evento Incorrecto. En caso contrario colocar de color verde ,comentario : Evento Correcto.


' Valores
' Hojas: Inicio, SantanderTerminaciones, reporte, Workday, Leaver
' reporte Rehire:
' -> A : Nombre del evemnto (flujo de trabajo)
' -> B : Iniciado por
' -> C : Iniciado
' -> D : Completado
' -> E : Resultado
' -> F : Aplicaciones
' -> G : Business
' -> H : Codigo Puesto
' -> I : Cost Center
' -> J : Domain Name
' -> K : Employee ID
' -> L : Nombre de Usuario
' -> M : Tipo de Movimiento
' -> N : Nombre Completo
' -> O : Is Delete
' -> P : Inactive
' -> Q : Puesto
' -> R : Ubicacion
' -> S : Departamento
' -> T : Effective Date
' -> U : Account Deletion Date
' -> V : Email
' Workday
' -> G : Tipo de Movimiento
' -> V : Employee ID
' Santander Terminaciones
' -> A : Identity
' -> B : Display Name
' -> C : Tipo de Movimiento
' -> D : Employee Id
' -> E : Inactive
' -> F : Application
' -> G : Account Deletion Date
' -> H : Disable

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
        If sheetWorkday.Cells(a, "G").Value = "B" Then
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

