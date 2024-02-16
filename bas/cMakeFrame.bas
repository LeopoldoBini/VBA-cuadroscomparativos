Attribute VB_Name = "cMakeFrame"
Option Explicit
Option Base 1

Public Function generarCuadros() As Worksheet

Dim nProv, nReng As Integer
Dim strObjcont As String, strTipoProc As String, strNumP As String, strAnoP As String
Dim cOfertas As Collection
Dim loTablaProv As ListObject, loTablaRenglones As ListObject
Dim rgNombresProv As Range
Dim arrNombresP, arrDetalleRenglones


Dim colOfShNames As New Collection

    nProv = Range("cantProv").Value2
    nReng = Range("cantReng").Value2
    strObjcont = Range("objetoProc").Value2
    strTipoProc = Range("tipoProc").Value
    strNumP = Range("numProc").Value2
    strAnoP = Range("anoProc").Value2

Set cOfertas = New Collection
Set loTablaProv = tableroProv.ListObjects("tablaProveedores")
Set rgNombresProv = loTablaProv.DataBodyRange.Columns(2)

If rgNombresProv.Rows.Count > 1 Then
    arrNombresP = WorksheetFunction.Transpose(rgNombresProv.Value2)
Else
    arrNombresP = Array(rgNombresProv.Value2)
End If
Set loTablaRenglones = tableroProv.ListObjects("tablaRenglones")
    arrDetalleRenglones = loTablaRenglones.DataBodyRange.Value2
Dim p As Integer

    For p = 1 To nProv
        colOfShNames.Add (p & " - " & Left(arrNombresP(p), 15) & "..")
    Next p
    
Dim hojaProv As Worksheet

Dim colOfertas As New Collection
Dim colCondiciones As New Collection
Dim oferta As Object
Dim condicion As Object

''''''''''''''''''''''''itero para cada proveedor y consigo coll de ofertas
    For p = 1 To nProv
            
        Set hojaProv = Worksheets(colOfShNames(p))
            hojaProv.Unprotect
        Dim lastrow As Long
            lastrow = hojaProv.Range("a4").CurrentRegion.Rows.Count
                    
        Dim rgOferta As Range
        Set rgOferta = hojaProv.Range("a5:g" & lastrow)
        Dim validator As Boolean
            validator = offerShValidator(rgOferta)
        If validator = False Then
            hojaProv.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
            Exit Function
        End If
        
        Dim arrOferta As Variant
            arrOferta = rgOferta.Value2
        
                        
        Dim r As Long
        For r = 1 To lastrow - 4
        
            If arrOferta(r, 4) <> Empty Then
                Set oferta = New clsOferta
                
                    With oferta
                        .nAlt = arrOferta(r, 3)
                        .nOrden = arrOferta(r, 1)
                        .nReng = arrOferta(r, 2)
                        .qOfert = arrOferta(r, 4)
                        .pUnit = arrOferta(r, 5)
                        .observacion = arrOferta(r, 7)
                        .nProv = p
                    End With
                    
                If arrOferta(r, 3) = Empty Then
                        oferta.prov = arrNombresP(p)
                Else
                        oferta.prov = arrNombresP(p) & " Alt. " & arrOferta(r, 3)
                End If
                    
                colOfertas.Add Item:=oferta
                Set oferta = Nothing
            End If
        Next r
        
        Set condicion = New clsCond
            
            With condicion
                .prov = arrNombresP(p)
                .mantOf = hojaProv.Range("i1").Value
                .formPago = hojaProv.Range("i2").Value
                .formEntrega = hojaProv.Range("i3").Value
            End With
            
            colCondiciones.Add Item:=condicion
        
        Set arrOferta = Nothing
        Set condicion = Nothing
        'hojaProv.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    Next p
    
''' Ordeno la coll de ofertas y consigo una coleccion ordenada de los renglones ofertados

Set colOfertas = ordenar(colOfertas)
Dim rengUnicos As Collection
Set rengUnicos = New Collection
    
For Each oferta In colOfertas
    Dim rengKey As String
    Dim reng As Integer
    Dim contador As Integer
    rengKey = "R" & oferta.nOrden
    reng = oferta.nOrden
    
    On Error Resume Next
''''''''''''''''''''''''''' Las voy agregando en orden
    If rengUnicos.Count = 0 Then
        rengUnicos.Add reng, rengKey
        GoTo nextoferta
    End If
    
    Dim ordReng As Variant
    contador = 1
    
    For Each ordReng In rengUnicos
        
        If reng < ordReng Then
            rengUnicos.Add reng, rengKey, contador
            GoTo nextoferta
        ElseIf (reng >= ordReng) And (rengUnicos.Count = contador) Then
            rengUnicos.Add reng, rengKey
        End If
        
        contador = contador + 1
        
    Next ordReng
    
    
nextoferta:
Next oferta
''''''''''''''''''''''''''''Tansformo en diccionario

Dim dicOfertas As Dictionary
Set dicOfertas = collectionToDictionary(colOfertas, rengUnicos)

'''''''''''''''''''''''''''''Consigo diccionario de lista de renglones ganados
Dim dicOfertasGanadas As Dictionary
Set dicOfertasGanadas = getRengByProv(dicOfertas, arrNombresP)


Dim dicDatosGrales As Dictionary
Set dicDatosGrales = New Dictionary

    dicDatosGrales.Add "tipoProc", strTipoProc
    dicDatosGrales.Add "numProc", strNumP
    dicDatosGrales.Add "anoProc", strAnoP
    dicDatosGrales.Add "cantReng", nReng
    dicDatosGrales.Add "cantProv", nProv
    dicDatosGrales.Add "objProc", strObjcont
    dicDatosGrales.Add "categoriaProc", Range("catProc").Value2
    dicDatosGrales.Add "organismoProc", Range("orgProc").Value2
    dicDatosGrales.Add "presupProc", Range("presupProc").Value2

Dim dicCondiciones As Dictionary
Set dicCondiciones = conditionCollectionToDictionary(colCondiciones)


'Debug.Print JsonConverter.ConvertToJson(dicCondiciones, Whitespace:=2)


''''''''''''''''''''''''''' consigo reng desiertos
ReDim Preserve arrDetalleRenglones(1 To UBound(arrDetalleRenglones), 1 To 5)

Dim rDesiertos As New Collection
Dim renglon As Variant

For r = 1 To nReng
    For Each renglon In rengUnicos
        If renglon = r Then
            GoTo nextR
        End If
    Next renglon
    rDesiertos.Add Item:=r
    'Agrego el dato de dsierto al arr para hacer el json
    arrDetalleRenglones(r, 5) = "Desierto"
nextR:
Next r

''''''''''''''''''''''''''''''Consigo el dic para mandar a google
Dim dicPaqueteFinal As Dictionary
Set dicPaqueteFinal = getJSONtoSendGoogle(dicDatosGrales, arrDetalleRenglones, dicOfertas, dicCondiciones)

Call sendPostGoogle(dicPaqueteFinal)



'''''''''''''''''''''''''''''' inserto Hoja de cuadros

Dim hojaCuadros As Worksheet
Set hojaCuadros = createCuadrosSh(strTipoProc, strNumP, strAnoP)
hojaCuadros.Visible = xlSheetVisible
hojaCuadros.Activate


'''''''''''''''''''''''''''''' Empiezo a meter cuadros por cada reng

Dim renglonOfertado As Variant


For Each renglonOfertado In rengUnicos

    Dim cuadroOF As Range
    Set cuadroOF = pasteCuadro(renglonOfertado, arrDetalleRenglones, hojaCuadros, _
                                strTipoProc, strNumP, strAnoP, strObjcont)
    
    
    Dim contadorOfxReng As Integer
        contadorOfxReng = 1
    
    Dim ofertRow As Range
    Dim rgOfert As Range
    Dim cantOfertas As Integer
        cantOfertas = colOfertas.Count
    
    Dim i As Integer
    'aca tengo que ir sacando mientras encuentro para hacer mas veloz
    For i = cantOfertas To 1 Step -1
                
         If colOfertas(i).nOrden = renglonOfertado Then
         
            If contadorOfxReng > 1 Then
                Set ofertRow = insOfReng(cuadroOF)
            Else
                Set ofertRow = cuadroOF.Cells(8, 1).EntireRow
            End If
                Set rgOfert = ofertRow.Range(Cells(1, 2), Cells(1, 8))
                    With rgOfert
                        .Cells(1) = renglonOfertado
                        .Cells(2) = colOfertas(i).nProv
                        .Cells(3) = colOfertas(i).prov
                        .Cells(4) = colOfertas(i).qOfert
                        .Cells(5) = colOfertas(i).pUnit
                        .Cells(6).FormulaR1C1 = "=r[0]c[-2] * r[0]c[-1]"
                        .Cells(7) = colOfertas(i).observacion
                        With .Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ThemeColor = 1
                            .TintAndShade = -0.349986266670736
                            .Weight = xlHairline
                        End With
'                        With .Borders(xlEdgeBottom)
'                            .LineStyle = xlContinuous
'                            .ThemeColor = 1
'                            .TintAndShade = -0.349986266670736
'                            .Weight = xlHairline
'                        End With

                        If arrDetalleRenglones(renglonOfertado, 4) < colOfertas(i).qOfert Then
                            With .Font
                                .Color = -16776961
                                .TintAndShade = 0
                            End With
                            With .Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = 13753087
                            End With
                        Else
                            With .Font
                                .ColorIndex = xlAutomatic
                                .TintAndShade = 0
                            End With
                            With .Interior
                                .Pattern = xlNone
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        End If
                    End With
                
            
            contadorOfxReng = contadorOfxReng + 1
         End If

        
    Next i
    
    contadorOfxReng = 1

Next renglonOfertado
''-+-++-+-+-+--------------+-+-++ Inserto cuadro de MejorEs Ofertas--------------------------------------
'Dim cuadroGanados As Range
'Dim rowGanadas As Range
'Set cuadroGanados = pasteGanados(hojaCuadros, strTipoProc, strNumP, strAnoP)
'
' For r = dicOfertasGanadas.Count To 1 Step -1
'        If r = dicOfertasGanadas.Count Then
'            Set ofertRow = cuadroGanados.Cells(6, 1).EntireRow
'        Else
'            Set ofertRow = insRengGanadores(cuadroGanados)
'        End If
'
'        Set rowGanadas = ofertRow.Range(Cells(1, 2), Cells(1, 8))
'
'
'            With rowGanadas
'                .Cells(1, 1) = r
'                .Cells(1, 3) = dicOfertasGanadas(r)("nombre")
'                .Cells(1, 4) = dicOfertasGanadas(r)("renglonesGanados")
'                .Cells(1, 7) = dicOfertasGanadas(r)("totalPesosGanados")
'
'                With .Borders(xlEdgeTop)
'                    .LineStyle = xlContinuous
'                    If r <> rDesiertos.Count - 1 Then
'                        .ThemeColor = 1
'                        .TintAndShade = -0.349986266670736
'                        .Weight = xlHairline
'                    End If
'                End With
'            End With
'        Set rowGanadas = Nothing
'    Next r



''--------------------------- inserto cuadro de desiertos y sus renglones--------------------------------
If rDesiertos.Count <> 0 Then

    Dim cuadroDesiertos As Range
    Dim rowDesierto As Range
    Set cuadroDesiertos = pasteDesiertos(hojaCuadros, strTipoProc, strNumP, strAnoP)
    
    For r = rDesiertos.Count To 1 Step -1
        If r = rDesiertos.Count Then
            Set ofertRow = cuadroDesiertos.Cells(7, 1).EntireRow
        Else
            Set ofertRow = insOfRengDesierto(cuadroDesiertos)
        End If
        
        Set rowDesierto = ofertRow.Range(Cells(1, 2), Cells(1, 8))
       

            With rowDesierto
                .Cells(1, 1) = arrDetalleRenglones(rDesiertos(r), 2)
                .Cells(1, 3) = arrDetalleRenglones(rDesiertos(r), 3)
                .Cells(1, 7) = arrDetalleRenglones(rDesiertos(r), 4)
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    If r <> rDesiertos.Count - 1 Then
                        .ThemeColor = 1
                        .TintAndShade = -0.349986266670736
                        .Weight = xlHairline
                    End If
                End With
            End With
        Set rowDesierto = Nothing
    Next r

End If
'------------------------------------------inserto condiciones------------------------------------------

Dim rgCondiciones As Range
Set rgCondiciones = pasteCondiciones(hojaCuadros, strTipoProc, strNumP, strAnoP)
Dim rowCondicion As Range



Dim justCondicion As Range
For r = 0 To colCondiciones.Count - 1
        If r = 0 Then
            Set rowCondicion = rgCondiciones.Cells(6, 1).EntireRow
        Else
            Set rowCondicion = insCondicion(rgCondiciones)
        End If
        
        Set justCondicion = rowCondicion.Range(Cells(1, 2), Cells(1, 7))
        
          With justCondicion
            .Cells(1, 1) = colCondiciones.Count - r
            .Cells(1, 3) = colCondiciones(colCondiciones.Count - r).prov
            .Cells(1, 4) = colCondiciones(colCondiciones.Count - r).mantOf
            .Cells(1, 5) = colCondiciones(colCondiciones.Count - r).formPago
            .Cells(1, 6) = colCondiciones(colCondiciones.Count - r).formEntrega
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                If r <> colCondiciones.Count - 1 Then
                    .ThemeColor = 1
                    .TintAndShade = -0.349986266670736
                    .Weight = xlHairline
                End If
            End With
        End With
                
Next r

Dim strFillTipoCont As String
If strTipoProc = "L.P." Then
    strFillTipoCont = "Licitación Pública"
End If
If strTipoProc = "C.A." Then
    strFillTipoCont = "Contratación Abreviada"
End If
If strTipoProc = "A.S." Then
    strFillTipoCont = "Adjudicación Simple"
End If


Dim header As String
header = strFillTipoCont & " " & "Nº" & strNumP & "/" & strAnoP

With hojaCuadros.PageSetup
    .PrintArea = hojaCuadros.UsedRange.Address
    .CenterHeader = "" & Chr(10) & "&B&14" & header & Chr(10) & strObjcont
    .FitToPagesWide = 1
    .FitToPagesTall = 0
    .LeftMargin = Application.InchesToPoints(0.7)
    .RightMargin = Application.InchesToPoints(0.7)
    .TopMargin = Application.InchesToPoints(0.85)
    .BottomMargin = Application.InchesToPoints(0.75)
    .HeaderMargin = Application.InchesToPoints(0.3)
    .FooterMargin = Application.InchesToPoints(0.3)
End With

generarCuadros = hojaCuadros


End Function








