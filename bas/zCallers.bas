Attribute VB_Name = "zCallers"
Option Explicit
Public Sub datosDelProc()
Dim msg As String
msg = "Procedimiento iniciado, No se puede volver a insertar hoja de ofertas sin Reiniciar el proceso"
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
        If celda.Value2 <> "" Then
            celda.Activate
            MsgBox msg, vbCritical
            Exit Sub
        End If
    Next celda


Dim rgProv As Range, rgReng As Range


Set rgProv = tableroProv.ListObjects("tablaProveedores").DataBodyRange
Set rgReng = tableroProv.ListObjects("tablaRenglones").DataBodyRange

If Not rgProv Is Nothing Then
For Each celda In rgProv
    If celda.Value2 <> "" Then
        celda.Activate
        MsgBox msg, vbCritical
        Exit Sub
    End If
Next celda
End If

If Not rgReng Is Nothing Then
For Each celda In rgReng
    If celda.Value2 <> "" Then
        celda.Activate
        MsgBox msg, vbCritical
        Exit Sub
    End If
Next celda
End If

If Worksheets.Count > 5 Then
    MsgBox msg, vbCritical
    Exit Sub
End If

ActiveSheet.Unprotect
formCantProv.Show
'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
End Sub


Public Sub insertarHojasDeOfertas()

ActiveWorkbook.Unprotect
If Worksheets.Count > 5 Then
    MsgBox "Procedimiento iniciado, No se puede volver a insertar hoja de ofertas sin Reiniciar el proceso", vbCritical
    'ActiveWorkbook.Protect Structure:=True, Windows:=False
    Exit Sub
End If

Dim conditionsAreOk As Boolean
    
conditionsAreOk = procDataValidator()
    
If conditionsAreOk Then
    conditionsAreOk = tablesDataValidator()
End If

If conditionsAreOk Then
    Call insertProvPages
End If

'ActiveWorkbook.Protect Structure:=True, Windows:=False
End Sub


Public Sub hacerElFuckingCuadro()
Application.ScreenUpdating = False
ActiveWorkbook.Unprotect

Dim shCuadroName As String, currentPath As String, stobjetoProc As String, _
    fileFullName As String, strTipoProc As String, strNumP As Variant, _
    strAnoP As String

Dim wb As Workbook
Dim ws As Worksheet

    strTipoProc = Range("tipoProc").Value
    strNumP = Range("numProc").Value2
    strAnoP = Range("anoProc").Value2


stobjetoProc = Range("objetoProc").Value
currentPath = ThisWorkbook.Path
        
shCuadroName = "Cuadro " & strTipoProc & " " & strNumP _
                & "-" & Replace(strAnoP, "20", "")

    Dim shCuadros As Worksheet
    Set shCuadros = generarCuadros()
   
Dim found As Boolean
For Each ws In Worksheets
    If ws.Name = shCuadroName Then
      ws.Move
      found = True
      Exit For
    End If
Next ws

If found = False Then
    'MsgBox "algo anda mal"
    Exit Sub
End If
    
For Each wb In Workbooks
    If wb.Worksheets.Count = 1 And wb.Worksheets(1).Name = shCuadroName Then
        wb.Unprotect
        fileFullName = currentPath & Application.PathSeparator & shCuadroName & " " _
            & stobjetoProc & ".xlsx"
            
        wb.SaveAs fileName:=fileFullName
        Shell "Explorer.exe /select, " & fileFullName, vbNormalFocus
        Exit For
    End If
    
    
    
Next wb
    


   ' shCuadros.Move
    '
   ' ActiveWorkbook.SaveAs (fileName)
    
   ' Shell "Explorer.exe /select, " & fileName, vbNormalFocus
'ActiveWorkbook.Protect Structure:=True, Windows:=False
Application.ScreenUpdating = True

End Sub


Public Sub reDo()
    tableroProv.Unprotect
    ThisWorkbook.Unprotect
    Application.ScreenUpdating = False
    
    Dim pregunta As String
    pregunta = MsgBox("Seguro que quiere borrar todo y volver a empezar???", vbYesNo + vbQuestion, "Hacer Cuadro Nuevo")
    
    If pregunta = vbYes Then
        Call deleteContents
        Call deleteSheets
    End If
    
    'ThisWorkbook.Protect Structure:=True, Windows:=False
    'tableroProv.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Application.ScreenUpdating = False
End Sub
