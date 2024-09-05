Sub ValidarJoiner()
    Dim hojaWorkday As Worksheet
    Dim sheetMacro As Worksheet

    Dim hojaReporte As Worksheet
    Dim createNombre As Variant
    Dim i As Long
    Dim a As Long
    Dim sheetLastRow As Long
    Dim reporteSheetLastRow As Long
    Dim ultimaFilaWorkday As Long
    Dim nextRow As Long
    Dim nextRowSheet As Long
    Dim reporteSheetRange As Range
    Dim sheetCell As Range
    Dim found As Range
    Dim rangoWorkday As Range
    Dim rangoSheetMacro As Range
    Dim domainName As String
    Dim nombreUsuarioReporteSheet As String
    Dim letrasnombreUsuarioRR As Variant
    Dim letraUnicaDeUsuario As String
    Dim dominioLetraPertenece As String
    Dim employeeIdReporteSheet As String
    Dim employeeIdSheetMacro As String
    Dim f As Long
    Dim d As Long
    Dim e As Long
    Dim g As Long
    Dim ADCorrespondiente As String

    Dim nombreMacro As String
    Dim nombreReporte As String
    
    nombreMacro = "REHIRE"
    nombreReporte = "santanderRehireEventReportJc"

    createNombre = Array(nombreMacro)
    Call crearHoja(createNombre)


    Set sheetMacro = ThisWorkbook.Sheets(nombreMacro)
    sheetMacro.Cells.Clear
    Set hojaWorkday = ThisWorkbook.Sheets("Workday")
    
    Set rangoWorkday = hojaWorkday.Rows(1)
    Set rangoSheetMacro = sheetMacro.Rows(1)
    rangoWorkday.Copy
    rangoSheetMacro.PasteSpecial Paste:=xlPasteAll
    ultimaFilaWorkday = hojaWorkday.Cells(hojaWorkday.Rows.Count, "B").End(xlUp).Row
    nextRowSheet = sheetMacro.Cells(sheetMacro.Rows.Count, "A").End(xlUp).Row + 1

            If nombreMacro = "JOINER" Then
                 For i = 1 To ultimaFilaWorkday
                    If hojaWorkday.Cells(i, "A").Value = " " Or IsEmpty(hojaWorkday.Cells(i, "A").Value) Then
                    hojaWorkday.Rows(i).Copy sheetMacro.Rows(nextRowSheet)
                    nextRowSheet = nextRowSheet + 1
                    End If
                 Next i
            Else
                 For i = 1 To ultimaFilaWorkday
                 If hojaWorkday.Cells(i, "A").Value = " " Or IsEmpty(hojaWorkday.Cells(i, "A").Value) Then
                 Else
                    hojaWorkday.Rows(i).Copy sheetMacro.Rows(nextRowSheet)
                    nextRowSheet = nextRowSheet + 1
                 End If
                 Next i
            End If
            
    
    Dim lastRowRehire As Long
    lastRowRehire = sheetMacro.Cells(sheetMacro.Rows.Count, "B").End(xlUp).Row
    For a = lastRowRehire To 2 Step -1
        If sheetMacro.Cells(a, "G").Value = "A" Then
        Else
           sheetMacro.Rows(a).Delete
        End If
    Next a

    'Aqui se colocara el quitado de duplicados.
    Call quitarDuplicados(nombreMacro)


    Set hojaReporte = ThisWorkbook.Sheets(nombreReporte)
    
    sheetMacro.Columns("W:W").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromatLeftOrAbove
    sheetMacro.Range("W1").Value = "Comentario"
    
    sheetLastRow = sheetMacro.Cells(sheetMacro.Rows.Count, "I").End(xlUp).Row
    reporteSheetLastRow = hojaReporte.Cells(hojaReporte.Rows.Count, "K").End(xlUp).Row
    
    Set reporteSheetRange = hojaReporte.Range("K2:K" & reporteSheetLastRow)
    
    


    If validarReporteConWorkday(nombreMacro, nombreReporte) Then
 
        For d = 2 To reporteSheetLastRow
        nombreUsuarioReporteSheet = hojaReporte.Cells(d, "L").Value
        employeeIdReporteSheet = hojaReporte.Cells(d, "K").Value
            If Trim(nombreUsuarioReporteSheet) = "" Then
                GoTo NextRecord
            End If
        ReDim tempArray(1 To Len(nombreUsuarioReporteSheet))
            For e = 1 To Len(nombreUsuarioReporteSheet)
                tempArray(e) = Mid(nombreUsuarioReporteSheet, e, 1)
            Next e
        letrasnombreUsuarioRR = tempArray
        letraUnicaDeUsuario = letrasnombreUsuarioRR(1)
        dominioLetraPertenece = obtenerLetraNombreUsuario(letraUnicaDeUsuario)
         For f = 2 To sheetLastRow
                employeeIdSheetMacro = sheetMacro.Cells(f, "I").Value
                If employeeIdReporteSheet = employeeIdSheetMacro Then
                    If validarTipoMovimiento(d, nombreReporte) Then
                              If hojaReporte.Cells(d, "J").Value = "No Domain" Or hojaReporte.Cells(d, "J").Value = "" Then
                                 ' Incorrecta
                                    sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(255, 0, 0)
                                    sheetMacro.Cells(f, "W").Value = "El campo Domain Name esta vacio o tiene el valor [ No Domain ]"

                               ElseIf hojaReporte.Cells(d, "J").Value = "SucurSales" And letraUnicaDeUsuario = "S" And hojaReporte.Cells(d, "J").Value <> "No Domain" Then
                        
                                     domainName = obtenerDomainName(d, nombreReporte)
                                     ADCorrespondiente = ADCorrespondienteMetodo(domainName)

                                     If obtenerAplicaciones(d, domainName, nombreReporte) Then
                                         If validarAplicaciones(d, ADCorrespondiente, nombreReporte) Then
                                            If validarCorreo(d, nombreReporte) Then
                                                sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(0, 255, 0) ' ROJO
                                            Else
                                                sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(255, 0, 0) ' ROJO
                                                sheetMacro.Cells(f, "W").Value = "El correo esta erroneo"
                                            End If
                                         Else
                                                    
                                                sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(255, 0, 0) ' ROJO
                                                sheetMacro.Cells(f, "W").Value = "El campo aplicaciones tiene aplicaciones Erroneas"
                                         End If

                                     Else
                                             sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(255, 0, 0) ' ROJO
                                             sheetMacro.Cells(f, "W").Value = "El campo aplicaciones no tiene el aplicativo correcto o no lo tiene"
                                     End If



                              ElseIf hojaReporte.Cells(d, "J").Value <> "SucurSales" And letraUnicaDeUsuario <> "S" And hojaReporte.Cells(d, "J").Value <> "No Domain" Then
                                     domainName = obtenerDomainName(d, nombreReporte)
                                     ADCorrespondiente = ADCorrespondienteMetodo(domainName)

                                     If obtenerAplicaciones(d, domainName, nombreReporte) Then
                                        If validarAplicaciones(d, ADCorrespondiente, nombreReporte) Then
                                             If validarCorreo(d, nombreReporte) Then
                                                sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(0, 255, 0) ' ROJO
                                            Else
                                                sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(255, 0, 0) ' ROJO
                                                sheetMacro.Cells(f, "W").Value = "El correo esta erroneo"
                                            End If
                                        Else
                                            sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(255, 0, 0) ' ROJO
                                            sheetMacro.Cells(f, "W").Value = "El campo aplicaciones tiene aplicaciones Erroneas"
                                        End If

                                    Else
                                            sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(255, 0, 0) ' ROJO
                                            sheetMacro.Cells(f, "W").Value = "El campo aplicaciones no tiene el aplicativo correcto o no lo tiene"
                                    End If
                    ElseIf hojaReporte.Cells(d, "J").Value = "SucurSales" And letraUnicaDeUsuario <> "S" And hojaReporte.Cells(d, "J").Value <> "No Domain" Then
                        ' Incorrecta
                         sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(255, 0, 0)
                         sheetMacro.Cells(f, "W").Value = "No coincide el Dominio con expediente local"

                    ElseIf hojaReporte.Cells(d, "J").Value <> "SucurSales" And letraUnicaDeUsuario = "S" And hojaReporte.Cells(d, "J").Value <> "No Domain" Then
                        ' Incorrecta
                         sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(255, 0, 0)
                         sheetMacro.Cells(f, "W").Value = "No coincide el Dominio con expediente local"

                    End If
                Else
                         sheetMacro.Range(sheetMacro.Cells(f, 1), sheetMacro.Cells(f, 22)).Interior.Color = RGB(255, 255, 0) ' Y
                         sheetMacro.Cells(f, "W").Value = "Movimiento diferente a A "
                End If

                   
                    
                End If
           Next f
NextRecord:

    Next d
       
    
       
    End If
     Call redirigir(nombreMacro)
       MsgBox "Ha terminado la macro"
End Sub
 

Sub redirigir(nombreMacro As String)
    Dim nombreHoja As String
    nombreHoja = nombreMacro
    On Error Resume Next
    Worksheets(nombreHoja).Activate
    On Error GoTo 0
    If Err.Number <> 0 Then
        MsgBox "No existe la hoja a la que se desea redirigir"
        Err.Clear
    End If

End Sub

Sub crearHoja(createNombre As Variant)
    Dim numHojas As Integer
    Dim i As Long
    Dim validacionHojaExistente As Boolean
    Dim workInicio As Worksheet
    Dim nombreHoja As String
    nombreHoja = createNombre(i)
    numHojas = UBound(createNombre) - LBound(createNombre) + 1
    For i = 1 To numHojas
        validacionHojaExistente = False
        For Each workInicio In ThisWorkbook.Sheets
            If workInicio.Name = createNombre(i - 1) Then
                validacionHojaExistente = True
            Exit For
            End If
    Next workInicio
        If validacionHojaExistente Then
            Select Case workInicio.Name
                Case nombreHoja
                    workInicio.Tab.Color = RGB(0, 81, 200)
            End Select
        Else
            Set workInicio = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            workInicio.Name = createNombre(i - 1)
            Select Case workInicio.Name
                Case nombreHoja
                    workInicio.Tab.Color = RGB(0, 81, 200)
            End Select
        End If
    Next i



End Sub
'Funcion Destinada a obtener la letra Nombre Usuario.
Function obtenerLetraNombreUsuario(letraUnicaDeUsuario As String) As String
    Select Case letraUnicaDeUsuario
        Case "S"
                obtenerLetraNombreUsuario = "SucurSales"
        Case Else
                obtenerLetraNombreUsuario = "DOMAIN"
    End Select
End Function

Function obtenerDomainName(numUser As Long, nombreReporte As String) As String
    Dim domainName As String
    Dim hojaReporte As Worksheet
    Set hojaReporte = ThisWorkbook.Sheets(nombreReporte)
    domainName = hojaReporte.Cells(numUser, "J").Value
    obtenerDomainName = domainName
End Function


Function obtenerAplicaciones(numUser As Long, domainName As String, nombreReporte As String) As Boolean
    Dim hojaReporte As Worksheet
    Dim aplicaciones As String
    Dim objectsA As Variant
    Dim isValid As Boolean
    Dim aplicacion As Variant
    Dim ADCorrespondiente As String

    ADCorrespondiente = ADCorrespondienteMetodo(domainName)

    Set hojaReporte = ThisWorkbook.Sheets(nombreReporte)
    aplicaciones = hojaReporte.Cells(numUser, "F")
    
        objectsA = Split(aplicaciones, ",")
        isValid = False
        For Each aplicacion In objectsA
             If Trim(aplicacion) = Trim(ADCorrespondiente) Then
                isValid = True
                 Exit For
             End If
         Next aplicacion

         If isValid Then
            obtenerAplicaciones = True
        Else
            obtenerAplicaciones = False
        End If
   
End Function

Function ADCorrespondienteMetodo(domainName As String) As String
    Select Case domainName
        Case "Corporativo"
            ADCorrespondienteMetodo = "AD - Corporativo"
        Case "Produban"
            ADCorrespondienteMetodo = "AD - Produban"
        Case "SucurSales"
            ADCorrespondienteMetodo = "AD - Sucursales"
        Case "ContactCenter"
            ADCorrespondienteMetodo = "AD - Contact Center"
        Case "Altec"
            ADCorrespondienteMetodo = "AD - Altec"
        Case "No Domain "
            ADCorrespondienteMetodo = "Error"
    End Select


End Function

Function validarAplicaciones(numUser As Long, ADCorrespondiente As String, nombreReporte As String) As String
Dim hojaReporte As Worksheet
Dim aplicaciones As String
Dim valoresPermitidosAplicacion As Variant
Dim aplicacionesConversion As Variant
'Dim isValid As Boolean

Set hojaReporte = ThisWorkbook.Sheets(nombreReporte)
aplicaciones = hojaReporte.Cells(numUser, "F")

If ADCorrespondiente = "AD - Produban" Then
    Dim ADObligatorio As String
    ADObligatorio = "AD - Altec"
    aplicacionesConversion = Split(aplicaciones, ",")
    If validarDominioProduban(aplicacionesConversion, ADObligatorio) Then
         valoresPermitidosAplicacion = Array("Workday", "GIM", "Azure Active Directory", "AD - Altec", ADCorrespondiente)
         validarAplicaciones = validarAplicacionesMetodo(aplicacionesConversion, valoresPermitidosAplicacion)
    Else
        validarAplicaciones = False
    End If
Else
    valoresPermitidosAplicacion = Array("Workday", "GIM", "Azure Active Directory", ADCorrespondiente)
    aplicacionesConversion = Split(aplicaciones, ",")
    validarAplicaciones = validarAplicacionesMetodo(aplicacionesConversion, valoresPermitidosAplicacion)
End If

End Function


Function validarDominioProduban(aplicacionesConversion As Variant, ADObligatorio As String) As Boolean
    Dim i As Long
    validarDominioProduban = False
    
    For i = LBound(aplicacionesConversion) To UBound(aplicacionesConversion)
        If aplicacionesConversion(i) = ADObligatorio Then
            validarDominioProduban = True
            Exit Function
        End If
    Next i
End Function

Function validarAplicacionesMetodo(aplicacionesConversion As Variant, valoresPermitidosAplicacion As Variant) As Boolean
    Dim a As Long
    Dim b As Long
    Dim valorPermitido As Boolean

    For a = LBound(aplicacionesConversion) To UBound(aplicacionesConversion)
        valorPermitido = False

        For b = LBound(valoresPermitidosAplicacion) To UBound(valoresPermitidosAplicacion)
            If aplicacionesConversion(a) = valoresPermitidosAplicacion(b) Then
                valorPermitido = True
                    Exit For
            End If
        Next b

        If Not valorPermitido Then
            validarAplicacionesMetodo = False
            Exit Function
        End If
    Next a
    validarAplicacionesMetodo = True
End Function

Sub quitarDuplicados(nombreMacro As String)
    Dim sheetDuplicate As Worksheet
    Dim lastRowSheetDuplicate As Long
    Dim Rng As Range
    Set sheetDuplicate = ThisWorkbook.Sheets(nombreMacro)
    lastRowSheetDuplicate = sheetDuplicate.Cells(sheetDuplicate.Rows.Count, "V").End(xlUp).Row
    Set Rng = sheetDuplicate.Range("A1:V" & lastRowSheetDuplicate)
    Rng.RemoveDuplicates Columns:=22, Header:=xlYes
End Sub


Function validarTipoMovimiento(d As Long, nombreReporte As String) As Boolean
    Dim hojaReporte As Worksheet
    Dim tipoDeMovimiento As String

    Set hojaReporte = ThisWorkbook.Sheets(nombreReporte)
    tipoDeMovimiento = hojaReporte.Cells(d, "M").Value

    If (tipoDeMovimiento = "A") Then
        validarTipoMovimiento = True
    Else
        validarTipoMovimiento = False
    End If
End Function


Function validarCorreo(d As Long, nombreReporte As String) As Boolean
    Dim hojaReporte As Worksheet
    Dim emailReporte As String
    Dim domainNameReporte As String
    Dim BusinessUnitDescription As String

    Set hojaReporte = ThisWorkbook.Sheets(nombreReporte)
    
    emailReporte = hojaReporte.Cells(d, "N").Value
    domainName = hojaReporte.Cells(d, "J").Value
    BusinessUnitDescription = hojaReporte.Cells(d, "G").Value

    If domainName = "SucurSales" And emailReporte = "" Then
        validarCorreo = True
    ElseIf domainName = "SucurSales" And emailReporte <> "" Then
        validarCorreo = False
    ElseIf BusinessUnitDescription = "SAM Asset Management, S.A" And emailReporte = "" Then
        validarCorreo = True
    ElseIf BusinessUnitDescription = "SAM Asset Management, S.A" And emailReporte <> "" Then
        validarCorreo = False
    ElseIf BusinessUnitDescription <> "SAM Asset Management, S.A" And emailReporte <> "" Then
        validarCorreo = True
    ElseIf domainName <> "SucurSales" And emailReporte = "" Then
        validarCorreo = False
    ElseIf domainName <> "SucurSales" And emailReporte <> "" Then
        validarCorreo = True
    End If

End Function


Function validarReporteConWorkday(nameSheetO As String, reporte As String) As Boolean
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

        If encontrado Then
             validarReporteConWorkday = True
        Else 
            nameSheet.Range(nameSheet.Cells(i, 1), nameSheet.Cells(i, 22)).Interior.Color = RGB(255, 255, 0)
            nameSheet.Cells(i, "W").Value = "No existe en la hoja Reporte"
           
        End If
    Next i
End Function










