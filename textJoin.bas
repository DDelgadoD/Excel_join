Attribute VB_Name = "Módulo1"
Function CTEXTJOIN(delimitador As String, ignoraVacios As Boolean, rango As Range) As String
Dim compiled As String
For Each cell In rango
    If ignoraVacios And IsEmpty(cell.Value) Then
    'nothing
    Else
    compiled = compiled + IIf(compiled = "", "", delimitador) + CStr(cell.Value)
    End If
Next
CTEXTJOIN = compiled
End Function
