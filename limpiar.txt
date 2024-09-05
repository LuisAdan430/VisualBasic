Sub Limpiar()
    Dim ws As Worksheet
    Dim hojasALimpiar As Variant
    Dim hoja As Variant

    hojasALimpiar = Array("Workday","santanderRehireEventReportJc","REHIRE")
    For Each hoja In hojasALimpiar
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(hoja)
        If Not ws In Nothing Then
            ws.Cells.Clear
        End If
        On Error GoTo 0
End Sub
