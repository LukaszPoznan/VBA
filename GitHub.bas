Attribute VB_Name = "GitHub"
Function labAdd(rowLabel As String, arr As Variant) As Integer

'===========================================
'= RETURNS POSITION OF A STRING IN AN ARRAY=
'= useful when a cell being referenced is  =
'= dynamic
'===========================================
    
    Dim i As Integer
    'DEFAULT RETURN VALUE IF VALUE NOT FOUND IN ARRAY
    labAdd = -1
    
    For i = LBound(arr) To UBound(arr)
        If StrComp(rowLabel, arr(i), vbTextCompare) = 0 Then
            labAdd = i
            Exit For
        End If
    Next
End Function

