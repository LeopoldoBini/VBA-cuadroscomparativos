Attribute VB_Name = "yAuxPaste"
Option Explicit

Public Function createCuadrosSh(tipoP, numP, anoP) As Worksheet

    Dim hojaCuadros As Worksheet
    
    modFinal.Copy before:=tableroProv
    
    Set hojaCuadros = Worksheets("mf (2)")
    
    ActiveWindow.DisplayGridlines = False
    
    Dim nombreHoja As String

    nombreHoja = "Cuadro " & tipoP & " " & numP _
                & "-" & Replace(anoP, "20", "")
    
    
        With hojaCuadros
            .Name = nombreHoja
'            .Range("A:C").ColumnWidth = 2.25
'            .Range("D:D").ColumnWidth = 38
'            .Range("E:F").ColumnWidth = 16
'            .Range("G:G").ColumnWidth = 20
'            .Range("H:H").ColumnWidth = 30
            With .Tab
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.499984740745262
            End With
        End With
    
    Set createCuadrosSh = hojaCuadros

End Function

Public Function pasteCuadro(nOrd, arrDetalle As Variant, cuadroSh As Worksheet, _
                            tipoP, numP, anoP, objCont) As Range
        
    Dim insertrow As Long
    insertrow = cuadroSh.UsedRange.Rows(cuadroSh.UsedRange.Rows.Count).Row + 1
    
    If insertrow = 2 Then
        insertrow = 1
    End If
     
    Dim titulo, detalle As String
    Dim cantSol As Long
    Dim rengPliego As Variant

    titulo = tipoP & " " & numP & "/" & Replace(anoP, "20", "")
    
    detalle = arrDetalle(nOrd, 3)
    cantSol = arrDetalle(nOrd, 4)
    rengPliego = arrDetalle(nOrd, 2)
                
    Range("modeloCuadro").Copy cuadroSh.Cells(insertrow, 1)
    
    Dim insertedCuadro As Range
    Dim rModelo, cModelo As Integer
        rModelo = Range("modeloCuadro").Rows.Count
        cModelo = Range("modeloCuadro").Columns.Count
        
    Set insertedCuadro = cuadroSh.Range(cuadroSh.Cells(insertrow, 1), _
                                        cuadroSh.Cells(insertrow + rModelo, cModelo))
    
    insertedCuadro.Cells(5, 1).RowHeight = 30
    If Len(detalle) < 60 Then
        insertedCuadro.Cells(6, 1).RowHeight = 25
    End If
    
    insertedCuadro.Cells(7, 1).RowHeight = 20
    insertedCuadro.Cells(3, 4) = titulo
    insertedCuadro.Cells(4, 4) = objCont
    insertedCuadro.Cells(6, 4) = rengPliego
    insertedCuadro.Cells(6, 5) = detalle
    insertedCuadro.Cells(6, 8) = cantSol
    
    Set pasteCuadro = insertedCuadro
    
End Function

Public Function insOfReng(cuadro As Range) As Range
'    Dim cuadro As Range
'    Set cuadro = Hoja2.Range("a1:h20")

    Dim insertOfert As Range
    Set insertOfert = cuadro.Cells(8, 1).EntireRow
    With insertOfert
        .Insert Shift:=xlDown, Copyorigin:=xlFormatFromRightOrBelow
    End With
    
    Set insOfReng = cuadro.Cells(8, 1).EntireRow
    
End Function
Public Function pasteGanados(cuadroSh As Worksheet, tipoP, numP, anoP) As Range
    Dim insertrow As Long
    insertrow = cuadroSh.UsedRange.Rows(cuadroSh.UsedRange.Rows.Count).Row + 2
    
    Dim insertedCuadro As Range
    Dim rModelo, cModelo As Integer
        rModelo = Range("modeloTOTALES").Rows.Count
        cModelo = Range("modeloTOTALES").Columns.Count
    
    Range("modeloTOTALES").Copy cuadroSh.Cells(insertrow, 1)
    
    Set insertedCuadro = cuadroSh.Range(cuadroSh.Cells(insertrow, 1), _
                                        cuadroSh.Cells(insertrow + rModelo, cModelo))
    Dim titulo As String
    titulo = tipoP & " " & numP & "/" & Replace(anoP, "20", "")
    
    With insertedCuadro
        .Cells(3, 4) = titulo
    End With
                            
    
    Set pasteGanados = insertedCuadro
End Function


Public Function pasteDesiertos(cuadroSh As Worksheet, tipoP, numP, anoP) As Range
    
    Dim insertrow As Long
    insertrow = cuadroSh.UsedRange.Rows(cuadroSh.UsedRange.Rows.Count).Row + 1
    
    Dim insertedCuadro As Range
    Dim rModelo, cModelo As Integer
        rModelo = Range("modeloDesiertos").Rows.Count
        cModelo = Range("modeloDesiertos").Columns.Count
    
    Range("modeloDesiertos").Copy cuadroSh.Cells(insertrow, 1)
    
    Set insertedCuadro = cuadroSh.Range(cuadroSh.Cells(insertrow, 1), _
                                        cuadroSh.Cells(insertrow + rModelo, cModelo))
    Dim titulo As String
    titulo = tipoP & " " & numP & "/" & Replace(anoP, "20", "")
    
    With insertedCuadro
        .Cells(4, 4) = titulo
    End With
                            
    
    Set pasteDesiertos = insertedCuadro
End Function

Public Function insRengGanadores(cuadro As Range) As Range
    Dim insertGanador As Range
    Set insertGanador = cuadro.Cells(6, 1).EntireRow
    With insertGanador
        .Insert Shift:=xlDown, Copyorigin:=xlFormatFromRightOrBelow
    End With

    Set insertGanador = cuadro.Cells(6, 1).EntireRow
    Set insRengGanadores = insertGanador
    
    
End Function


Public Function insOfRengDesierto(cuadro As Range) As Range

    Dim insertOfert As Range
    Set insertOfert = cuadro.Cells(7, 1).EntireRow
    With insertOfert
        .Insert Shift:=xlDown, Copyorigin:=xlFormatFromRightOrBelow
    End With
    
    Set insertOfert = cuadro.Cells(7, 1).EntireRow
    Set insOfRengDesierto = insertOfert
    
End Function

Public Function pasteCondiciones(cuadroSh As Worksheet, tipoP, numP, anoP) As Range
    
    Dim insertrow As Long
    insertrow = cuadroSh.UsedRange.Rows(cuadroSh.UsedRange.Rows.Count).Row + 1
    
    Dim insertedCuadro As Range
    Dim rModelo, cModelo As Integer
        rModelo = Range("modeloCond").Rows.Count
        cModelo = Range("modeloCond").Columns.Count
        
    
    Range("modeloCond").Copy cuadroSh.Cells(insertrow, 1)
     
    Set insertedCuadro = cuadroSh.Range(cuadroSh.Cells(insertrow, 1), _
                                        cuadroSh.Cells(insertrow + rModelo, cModelo))

    Dim titulo As String
    titulo = tipoP & " " & numP & "/" & Replace(anoP, "20", "")
    
    With insertedCuadro
        .Cells(3, 4) = titulo
    End With
                                        
    Set pasteCondiciones = insertedCuadro
    
End Function
Public Function insCondicion(cuadro As Range) As Range

    Dim insertOfert As Range
    Set insertOfert = cuadro.Cells(6, 1).EntireRow
    With insertOfert
        .Insert Shift:=xlDown, Copyorigin:=xlFormatFromRightOrBelow
    End With
    
    Set insCondicion = cuadro.Cells(6, 1).EntireRow
    
    
End Function
' insertar renglones en las paginas de ofertas de los proveedores
Public Function insertOfferRow(offerSh As Worksheet) As Range

    Dim insertedOfertRow As Range
    Set insertedOfertRow = offerSh.Cells(5, 1).EntireRow
    With insertedOfertRow
        .Insert Shift:=xlDown, Copyorigin:=xlFormatFromRightOrBelow
    End With
    
    Set insertedOfertRow = offerSh.Range(Cells(5, 1), Cells(5, 9))
    With insertedOfertRow.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.25
    End With
    With insertedOfertRow.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.25
    End With
    With insertedOfertRow.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.2
    End With

    Set insertOfferRow = insertedOfertRow

End Function


