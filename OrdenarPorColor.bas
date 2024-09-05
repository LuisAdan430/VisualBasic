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
        MsgBox "An error ocurred: " & Err.Description, vbExclamation
End Sub

