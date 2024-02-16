Attribute VB_Name = "aPromptForm"
Option Explicit

'Usa los datos del form
Public Sub StepOne(nProv As Integer, nReng As Integer, tProc As String, nProc As Variant, aProc As Integer, objetoProc As String, soliProc As String, catProc As String)

    Range("tipoProc") = tProc
    Range("numProc") = nProc
    Range("anoProc") = aProc
    Range("cantReng") = nReng
    Range("cantProv") = nProv
    Range("objetoProc") = objetoProc
    Range("catProc") = catProc
    Range("orgProc") = soliProc
    
    
Dim prov As Integer
Dim reng As Integer
       
    Dim tablaProv, tablaReng As ListObject
    Dim lRow As Range
    Set tablaProv = tableroProv.ListObjects("tablaProveedores")
    Set tablaReng = tableroProv.ListObjects("tablaRenglones")

    
    For prov = 1 To nProv
    
        tablaProv.ListRows.Add
        Set lRow = tablaProv.ListRows(prov).Range
        lRow.Cells(1, 1) = prov
        
    Next prov

    For reng = 1 To nReng
    
        tablaReng.ListRows.Add
        Set lRow = tablaReng.ListRows(reng).Range
        lRow.Cells(1, 1) = reng
    
    Next reng

    
End Sub
