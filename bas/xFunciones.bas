Attribute VB_Name = "xFunciones"
Option Explicit
Public Sub segunPliego()
    Dim ws As Worksheet, rg As Range
    Set ws = ActiveSheet
    Set rg = ws.Range("i1:i3")
    
    Dim ans
    
    ans = MsgBox("Segun Pliego??", VbMsgBoxStyle.vbYesNo, "Llenamos Segun Pliego todo?")
    If ans = vbYes Then
        rg.Value = "Seg�n Pliego"
    End If

End Sub

Public Function getJSONtoSendGoogle(DatosGenerales As Dictionary, arrDetalle, _
                                    DicOfertasRecibidas As Dictionary, _
                                    DicCondicionOfertas As Dictionary)

Dim Paquete As Dictionary
Set Paquete = New Dictionary



Paquete.Add "generalesContratacion", DatosGenerales
Paquete.Add "detallesContratacion", arrDetalle
Paquete.Add "ofertasRecibidas", DicOfertasRecibidas
Paquete.Add "condicionesOfertas", DicCondicionOfertas


Set getJSONtoSendGoogle = Paquete

End Function



Public Sub sendPostGoogle(dicPaquete As Dictionary)
Dim strJson As String
strJson = JsonConverter.ConvertToJson(dicPaquete, 2)


Dim Req As New MSXML2.XMLHTTP60

Dim reqURL As String

reqURL = "URLXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"




Req.Open "POST", reqURL, False
Req.setRequestHeader "Content-type", "application/json"
Req.send strJson

If Req.Status <> 200 Then

    MsgBox Req.Status & " - " & Req.statusText
    Exit Sub

Else
    
    'MsgBox Req.Status & " - Datos Subidos con exito"
    
End If

End Sub

Public Function conditionCollectionToDictionary(condCol As Collection) As Dictionary
Dim setCondition

Dim outerDic As Dictionary
Dim innerDic As Dictionary
Set outerDic = New Dictionary



For Each setCondition In condCol
    Set innerDic = New Dictionary
        
        innerDic.Add "formaEntrega", setCondition.formEntrega
        innerDic.Add "formaPago", setCondition.formPago
        innerDic.Add "mantenimientoOferta", setCondition.mantOf
        innerDic.Add "nombreProveedor", setCondition.prov
        
        outerDic.Add outerDic.Count + 1, innerDic
    

Next setCondition

Set conditionCollectionToDictionary = outerDic

End Function



Public Function collectionToDictionary(Coll As Collection, renglonesOfertados) As Dictionary

 
Dim FullofDictionary As Dictionary
Dim dicRenglon As Dictionary
Dim ofDict As Dictionary

Set FullofDictionary = New Dictionary
Dim rng

For Each rng In renglonesOfertados
    Set dicRenglon = New Dictionary
    FullofDictionary.Add rng, dicRenglon
Next rng

 
 Dim of
 
    For Each of In Coll

        Set ofDict = New Dictionary
    
        ofDict.Add "prov", of.prov
        ofDict.Add "nProv", of.nProv
        ofDict.Add "nOrden", of.nOrden
        ofDict.Add "nReng", of.nReng
        ofDict.Add "nAlt", of.nAlt
        ofDict.Add "qOfert", of.qOfert
        ofDict.Add "pUnit", of.pUnit
        ofDict.Add "observacion", of.observacion
        ofDict.Add "numProv", of.numProv

        
        FullofDictionary(of.nOrden).Add FullofDictionary(of.nOrden).Count + 1, ofDict

    Next of

' Dim innerKey As Variant
' Dim midKey As Variant
' Dim outerKey As Variant
'
'    For Each outerKey In FullofDictionary.Keys
'        Debug.Print "  Renglon N�:"; outerKey
'
'            For Each midKey In FullofDictionary(outerKey).Keys
'                Debug.Print "    Orden Merito:"; midKey
'
'                For Each innerKey In FullofDictionary(outerKey)(midKey)
'
'                    Debug.Print "      " & innerKey; ": "; FullofDictionary(outerKey)(midKey)(innerKey)
'                Next innerKey
'
'                 Debug.Print "   ----------------"
'            Next midKey
'
'        'PrintDictionary outer(outerKey)
'
'   Debug.Print "===================="
'    Next outerKey
        
Set collectionToDictionary = FullofDictionary

End Function

Public Function getRengByProv(fulOfDic As Dictionary, arrProv) As Dictionary


' -------------Creo el diccionario para cada proveedor

 Dim dicRenglonesPorProveedor As Dictionary
 Dim renglonesDeProv As Dictionary
 Dim Renglones As Collection
 Dim totalPesos As Double
 
 Dim nombreProv
 

 Set dicRenglonesPorProveedor = New Dictionary
 For Each nombreProv In arrProv
    Set renglonesDeProv = New Dictionary
    Dim strRengList As String
    
        renglonesDeProv.Add "nombre", nombreProv
        renglonesDeProv.Add "renglonesGanados", strRengList
        renglonesDeProv.Add "totalPesosGanados", totalPesos
        
    
    dicRenglonesPorProveedor.Add dicRenglonesPorProveedor.Count + 1, renglonesDeProv
 
 Next nombreProv
 
'-+-+-+-+-------------Recorro El diccionario de Ofeertas, y voy rellenando el dic de cada prov, totalizando la oferta
 
 
 
 Dim propOfertaKey As Variant
 Dim ordenMeritoKey As Variant

 Dim currWinner
 Dim currWinnerRengList
 Dim lng As Integer
    
 Dim renglonKey
    
    For Each renglonKey In fulOfDic.Keys
    
        Set currWinner = dicRenglonesPorProveedor(fulOfDic(renglonKey)(1)("nProv"))
        
        lng = Len(currWinner("renglonesGanados"))
        
        If lng = 0 Then
        
            currWinner("renglonesGanados") = renglonKey
             
        Else
        
            currWinner("renglonesGanados") = currWinner("renglonesGanados") & ", " & renglonKey
        
        End If
        
            
            currWinner("totalPesosGanados") = currWinner("totalPesosGanados") + _
                                    (fulOfDic(renglonKey)(1)("pUnit") * fulOfDic(renglonKey)(1)("qOfert"))
            
    Next renglonKey
                


' Dim midKey As Variant
' Dim outerKey As Variant
'
'    For Each outerKey In dicRenglonesPorProveedor.Keys
'        Debug.Print "  Proveedor N�:"; outerKey
'
'            For Each midKey In dicRenglonesPorProveedor(outerKey).Keys
'                Debug.Print "      "; midKey; ": "; dicRenglonesPorProveedor(outerKey)(midKey)
'            Next midKey
'
'   Debug.Print "===================="
'    Next outerKey
        
        
    Set getRengByProv = dicRenglonesPorProveedor


End Function


Public Sub NestedDictionaryExample()
    
    Dim outer As Dictionary
    Dim inner As Dictionary
    
    Set outer = New Dictionary
    
    Dim i As Long
    For i = 1 To 10
        Set inner = New Dictionary
        inner.Add 10 * i, "Value of inner dictionary ..."
        inner.Add 100 * i, "Another value of inner dictionary ..."
        inner.Add 1000 * i, "Third value of inner dictionary ..."
        outer.Add i, inner
    Next i
    
    Dim innerKey As Variant
    Dim outerKey As Variant
    
    For Each outerKey In outer.Keys
        Debug.Print "Outer key:"; outerKey
        Debug.Print "Inner key: value"
        
        'PrintDictionary outer(outerKey)
        For Each innerKey In outer(outerKey)
            Debug.Print innerKey; ": "; outer(outerKey)(innerKey)
        Next innerKey
        Debug.Print "----------------"
        
    Next outerKey
    
End Sub
 
Public Sub PrintDictionary(myDict As Object)
    
    Dim key     As Variant
    For Each key In myDict.Keys
        Debug.Print key; "-->"; myDict(key)
    Next key
    
End Sub

Public Function ordenar(deOrderedColl As Collection)

    Dim orderedColl As New Collection
    
    Dim oferta As Variant
    
    For Each oferta In deOrderedColl
        Dim contador As Integer
        contador = 1

        If orderedColl.Count = 0 Then
            orderedColl.Add oferta
            GoTo nextoferta
        End If
        
        Dim ordOferta As Variant

        For Each ordOferta In orderedColl
            
            If oferta.pUnit < ordOferta.pUnit Then
                orderedColl.Add Item:=oferta, before:=contador
                GoTo nextoferta
            ElseIf (oferta.pUnit >= ordOferta.pUnit) And (contador = orderedColl.Count) Then
                orderedColl.Add Item:=oferta
            End If
                
            contador = contador + 1
            
        Next ordOferta
            
nextoferta:
    Next oferta
    
   Set ordenar = orderedColl

End Function
Sub inputAlt()
   Dim mensaje As String
   mensaje = InputBox("N� de Orden:", "Ingresar Alternativa")
    
    If mensaje = "" Then
        Exit Sub
    End If
 
    If IsNumeric(mensaje) Then
    ActiveSheet.Unprotect
    Call insertAlternative(CInt(mensaje))
    Else
        MsgBox "Tiene que ser un Numero"
        Call inputAlt
    End If
    'ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub


Sub insertAlternative(nOrdenToInsert As Long)


    Dim lastrow As Long
        lastrow = Range("a4").CurrentRegion.Rows.Count
        
    Dim rgCurrentRows As Range
    Dim arrCurr As Variant
    Set rgCurrentRows = Range("a5:i" & lastrow)
        arrCurr = rgCurrentRows.Value2
     
    Dim r As Integer
        
    Dim firstAppearance As Long
    
    For r = 1 To rgCurrentRows.Rows.Count
        If arrCurr(r, 1) = nOrdenToInsert Then
            firstAppearance = r
            Exit For
        End If
    Next r
    
    If firstAppearance = 0 Then
        MsgBox "Numero de Orden indicado FUERA DE RANGO. Prest� atenci�n!"
        Exit Sub
    End If
    Dim counter As Integer
    counter = 0
        
    For r = firstAppearance To rgCurrentRows.Rows.Count
        If arrCurr(r, 1) = nOrdenToInsert Then
            counter = counter + 1
        Else
            Exit For
        End If
    Next r
    
    Dim maxAlt As Integer
    maxAlt = 0

    For r = 0 To counter - 1
        If IsNumeric(arrCurr(firstAppearance + r, 3)) Then
            If arrCurr(firstAppearance + r, 3) > maxAlt Then
                maxAlt = arrCurr(firstAppearance + r, 3)
            End If
        End If
    Next r

    Dim lastAppearance As Long
    lastAppearance = firstAppearance + counter - 1
    Dim insertOfertRow As Range
    Set insertOfertRow = rgCurrentRows.Cells(lastAppearance + 1, 1).EntireRow
    With insertOfertRow
        If lastAppearance = rgCurrentRows.Rows.Count Then
            .Insert Shift:=xlDown, Copyorigin:=xlFormatFromLeftOrAbove
        Else
            .Insert Shift:=xlDown, Copyorigin:=xlFormatFromRightOrBelow
        End If
    End With
    
    Set insertOfertRow = rgCurrentRows.Range(Cells(firstAppearance + counter, 1), _
                                            Cells(firstAppearance + counter, 9))
    
    
    With insertOfertRow
        .Cells(1, 1) = nOrdenToInsert
        .Cells(1, 2) = arrCurr(firstAppearance, 2)
        .Cells(1, 3) = maxAlt + 1
        .Cells(1, 6).FormulaR1C1 = "=r[0]c[-2] * r[0]c[-1]"
        .Cells(1, 8) = arrCurr(firstAppearance, 8)
        .Cells(1, 9) = arrCurr(firstAppearance, 9)
    End With
    
    With insertOfertRow.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.25
    End With
    With insertOfertRow.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.25
    End With
    With insertOfertRow.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.2
    End With


End Sub

Public Function procDataValidator()
Dim areValid As Boolean
    areValid = True
    
Dim data As New Collection

    data.Add Range("cantProv")
    data.Add Range("cantReng")
    data.Add Range("objetoProc")
    data.Add Range("tipoProc")
    data.Add Range("numProc")
    data.Add Range("anoProc")
    data.Add Range("presupProc")
    data.Add Range("orgProc")
    data.Add Range("catProc")

Dim celda As Range

    For Each celda In data
        If celda.Value2 = "" Then
            areValid = False
            celda.Activate
            MsgBox "Falta indicar " & celda.Offset(-1, 0).Value2
            GoTo final
        End If
    Next celda
final:
   procDataValidator = areValid

End Function

Public Function tablesDataValidator()
Dim areValid As Boolean
    areValid = True

Dim loProv As ListObject, loReng As ListObject
Dim arrProv, arrReng

Set loProv = tableroProv.ListObjects("tablaProveedores")
Set loReng = tableroProv.ListObjects("tablaRenglones")
    arrProv = loProv.DataBodyRange.Value2
    arrReng = loReng.DataBodyRange.Value2

Dim r As Long
Dim c As Long

    For r = 1 To loProv.DataBodyRange.Rows.Count
        If arrProv(r, 2) = "" Then
            areValid = False
            loProv.DataBodyRange.Cells(r, 2).Activate
            MsgBox "Falta indicar el nombre del Proveedor N�" & loProv.DataBodyRange.Cells(r, 2).Offset(0, -1).Value2
            GoTo final
        End If
    Next r
    
    For r = 1 To loReng.DataBodyRange.Rows.Count
        For c = 2 To 4
            If arrReng(r, c) = "" Then
                areValid = False
                loReng.DataBodyRange.Cells(r, c).Activate
                MsgBox "Falta indicar dato para Renglon N� de Orden  " & loReng.DataBodyRange.Cells(r, c).Offset(0, -(c - 1)).Value2
                GoTo final
            End If
        Next c
    Next r
final:
tablesDataValidator = areValid
End Function

Public Function offerShValidator(rgOferta As Range) As Boolean

Dim isValid As Boolean
    isValid = False
Dim arr
    arr = rgOferta.Value2

Dim r As Long
For r = 1 To rgOferta.Rows.Count
    If arr(r, 4) <> Empty Then
        If arr(r, 5) <> Empty Then
            isValid = True
        Else
            MsgBox "Falta indicar el PRECIO UNITARIO de : " & rgOferta.Cells(r, 9), vbCritical, rgOferta.Parent.Name
            rgOferta.Cells(r, 5).Activate
            isValid = False
            GoTo final
        End If
    End If
    If arr(r, 5) <> Empty Then
        If arr(r, 4) <> Empty Then
            isValid = True
        Else
            MsgBox "Falta indicar la CANTIDAD OFRECIDA de : " & rgOferta.Cells(r, 9), vbCritical, rgOferta.Parent.Name
            rgOferta.Cells(r, 4).Activate
            isValid = False
            GoTo final
        End If
    End If
        
Next r

If isValid = False Then
    MsgBox "Faltan cargar las ofertas de : " & rgOferta.Parent.Range("b1").Value2, vbCritical, rgOferta.Parent.Name
    rgOferta.Parent.Activate
    GoTo final
End If


Dim rgCondiciones As Range
Set rgCondiciones = rgOferta.Parent.Range("i1:i3")

Dim celda As Range
For Each celda In rgCondiciones
    If celda = Empty Then
        isValid = False
        MsgBox "Falta cargar " & celda.Offset(0, -1), vbCritical, rgOferta.Parent.Name
        GoTo final
    End If
Next celda


final:
offerShValidator = isValid

End Function
'-----------------------BORRAR HOJAS DEMAS--------------------------------------------------
Public Sub deleteSheets()
Application.DisplayAlerts = False 'switching off the alert button
Dim sh As Worksheet

    For Each sh In ThisWorkbook.Worksheets
       
        If sh.CodeName = "modCuadro" _
        Or sh.CodeName = "modFinal" _
        Or sh.CodeName = "modOferta" _
        Or sh.CodeName = "tableroProv" _
        Or sh.CodeName = "apuntes" Then
            GoTo nextSh
        End If
            
        sh.Delete
nextSh:
    Next sh
Application.DisplayAlerts = True 'switching off the alert button
End Sub
'------------------------Borrar datos del tablero----------------------------------------
Public Sub deleteContents()
Dim loProv As ListObject
Dim loReng As ListObject

Set loProv = tableroProv.ListObjects("tablaProveedores")
Set loReng = tableroProv.ListObjects("tablaRenglones")

Range("objetoProc").ClearContents
Range("tipoProc").ClearContents
Range("numProc").ClearContents
Range("anoProc").ClearContents
Range("cantReng").ClearContents
Range("cantProv").ClearContents
Range("presupProc").ClearContents
Range("orgProc").ClearContents
Range("catProc").ClearContents

On Error Resume Next
loProv.DataBodyRange.Delete
loReng.DataBodyRange.Delete

apuntes.UsedRange.EntireColumn.Delete Shift:=xlToLeft

End Sub

Public Function getPath()
Dim currentPath As String
Dim stobjetoProc As String, fileName As String
    

    
getPath = fileName
End Function



