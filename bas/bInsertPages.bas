Attribute VB_Name = "bInsertPages"
Option Explicit
Option Base 1


Public Sub insertProvPages()
Application.ScreenUpdating = False
Dim nProv As Integer
Dim nReng As Integer

    nProv = Range("cantProv")
    nReng = Range("cantReng")

Dim loTablaProv As ListObject
Dim rgNombresProv As Range
Dim arrNombres As Variant

    Set loTablaProv = tableroProv.ListObjects("tablaProveedores")
    
    Set rgNombresProv = loTablaProv.DataBodyRange.Columns(2)
    
    If rgNombresProv.Rows.Count > 1 Then
        arrNombres = WorksheetFunction.Transpose(rgNombresProv.Value2)
    Else
        arrNombres = Array(rgNombresProv.Value2)
    End If
    
               
Dim loTablaReng As ListObject
Dim rgRenglonesPedidos As Range
Dim arrRenglonesPedidos As Variant
    
    Set loTablaReng = tableroProv.ListObjects("tablaRenglones")
    Set rgRenglonesPedidos = loTablaReng.DataBodyRange
    arrRenglonesPedidos = rgRenglonesPedidos.Value2


               
Dim page As Integer
Dim shDeOferta As Worksheet

    For page = nProv To 1 Step -1
          
        modOferta.Copy After:=tableroProv
        
        Set shDeOferta = Worksheets("mo (2)")
            shDeOferta.Visible = xlSheetVisible
            shDeOferta.Activate
            ActiveWindow.DisplayHeadings = False

        With shDeOferta
            .Name = page & " - " & Left(arrNombres(page), 15) & ".."
            .Cells(1, 1) = page
            .Cells(1, 2) = arrNombres(page)
        End With
        
        
        Dim r As Long
        'Cargando los detalles de la sol. p cada pag
        For r = nReng To 1 Step -1
            Dim rgRowTemplate As Range
            If r = nReng Then
                Set rgRowTemplate = shDeOferta.Range("a5:i5")
            Else
                Set rgRowTemplate = insertOfferRow(shDeOferta)
            End If
            
                With rgRowTemplate
                    .Cells(1, 1) = arrRenglonesPedidos(r, 1)
                    .Cells(1, 2) = arrRenglonesPedidos(r, 2)
                    .Cells(1, 6).FormulaR1C1 = "=r[0]c[-2] * r[0]c[-1]"
                    .Cells(1, 8) = arrRenglonesPedidos(r, 4)
                    .Cells(1, 9) = arrRenglonesPedidos(r, 3)
                End With
        Next r
        
        ActiveWindow.DisplayGridlines = False
        'shDeOferta.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    Next page

Application.ScreenUpdating = True

End Sub


