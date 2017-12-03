'============================================================
'= LABEL ADDRESS - RETURNS POSITION OF A STRING IN AN ARRAY =
'= useful when a cell being referenced is dynamic           =
'============================================================
Attribute VB_Name = "GitHub_functions"
Function labAdd(rowLabel As String, arr As Variant) As Integer
    
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

'===============================
'= RETUR A DATA TYPE IN A CELL =
'===============================

Function CellType(pRange As Range)
    
Application.Volatile
Set pRange = pRange.Range("A1")
Select Case True
    Case VBA.IsEmpty(pRange): CellType = "Blank"
    Case Application.IsText(pRange): CellType = "Text"
    Case Application.IsLogical(pRange): CellType = "Logical"
    Case Application.IsErr(pRange): CellType = "Error"
    Case VBA.IsDate(pRange): CellType = "Date"
    Case VBA.InStr(1, pRange.Text, ":") <> 0: CellType = "Time"
    Case VBA.IsNumeric(pRange): CellType = "Value"
End Select
End Function

'=================================================
'= RETURNS A URL OF A FIRST HYPERLINK IN A RANGE =
'=================================================

Function GetURL(rng As Range) As String
   
    On Error Resume Next
    GetURL = rng.Hyperlinks(1).address
End Function
