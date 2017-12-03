Attribute VB_Name = "GitHub_candidatesForm"
'=============================================================================================================================================
'= FORMATS DATA FROM AN EMAIL FORM INTO AN ACCESS FORM                                                                                       =
'= SELECTED OPERATIONS:                                                                                                                      =
'= - creates arrays with header cells (trimmed) for easy reference using the custom Label Address function (please see the Fubctions module) =
'= - creates buttons with the copy function                                                                                                  =
'= - cleans and trims labels (using split and trim functions)                                                                                =
'= - determines end of data range and number of actual records                                                                               =
'= - loops and if statements                                                                                                                 =
'= - using string, int, and boolean variables                                                                                                =
'= - adding comments                                                                                                                         =
'= - basic operations such as deleting, moving, hiding, and inserting columns                                                                =
'=============================================================================================================================================

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
Sub copRef()
    Dim arrCopy(1 To 80) As String
        For i = 1 To 80
            arrCopy(i) = Cells(i, 1)
        Next
        copRowLoc = labAdd("NOTES IN BOX (Y/N)", arrCopy) + 3
    Range(Cells(copRowLoc, 1), Cells(copRowLoc, labAdd("NOTES IN BOX (Y/N)", arrCopy))).Select
    Selection.Copy
End Sub
Sub copAct()
    Dim arrAct(1 To 80) As String
        For i = 1 To 80
            arrAct(i) = Cells(i, 1)
        Next
        copStaLoc = labAdd("CANDIDATE ID - ACTIVITIES", arrAct) + 1
        copEndLoc = labAdd("CANDIDATE ID - FEE INFORMATION", arrAct) - 2
    Range(Cells(copStaLoc, 1), Cells(copEndLoc, 3)).Select
    Selection.Copy
End Sub
Sub copFee()
    Dim arrFee(1 To 32) As String
        For i = 1 To 32
            arrFee(i) = Cells(i, 7)
        Next
    Dim arr1stCol(1 To 80) As String
        For i = 1 To 80
            arr1stCol(i) = Cells(i, 1)
        Next
        copRowLoc = labAdd("CANDIDATE ID", arr1stCol) + 1
    Range(Cells(copRowLoc, 1), Cells(copRowLoc, labAdd("CHARGE CODE", arrFee))).Select
    Selection.Copy
End Sub
Sub buttons()
'ADDING COPYING BUTTONS
    Dim w As Worksheet
    Dim b As Button
    
    Set w = ActiveSheet
    Set b = w.buttons.Add(5, 5, 80, 18, 75)
    b.OnAction = "cCopy1"
    b.Characters.Text = "Copy Referral Source"
        Cells(1, 7) = "testUSUN TO"
End Sub
Sub candidatesFormPro()
'SIMPLIFYING LABELS
    'SPLITTING LABELS
        For i = 1 To 80
            Range(Cells(i, 3), Cells(i, 3)) = Split(Cells(i, 1), "(")
        Next
    'OVERWRITING ORIGINAL WITH SPLITTED VALUES
        Range("C1:C80").Select
        Selection.Cut
        Range("A1:A80").Select
        ActiveSheet.Paste
'TRIMMING before mapping
    Dim oCell As Range
    Dim func As WorksheetFunction
    
    Set func = Application.WorksheetFunction
    Set trimRange = Range("A1:A80")
    
    For Each oCell In trimRange
        oCell = func.Trim(oCell)
    Next
'MAPPING
    Dim arrRaw(1 To 80) As String
    For i = 1 To 80
        arrRaw(i) = Cells(i, 1)
    Next
'ADDING MISSING ROWS
    'JOB ID
        Cells(1, 2) = UCase(Cells(1, 1))
        Cells(1, 1) = "CANDIDATE ID - REFERRAL SOURCE"
        Cells(2, 1).Select
        Selection.EntireRow.Insert
    'STATUS
        Cells(2, 1) = "STATUS"
            Cells(2, 2) = "Active"
        'MAPPING (required to do after every row insert)
            Dim arrRaw2(1 To 80) As String
                For i = 1 To 80
                    arrRaw2(i) = Cells(i, 1)
                Next
    'SPECIAL CIRCUMSTANCES
        If labAdd("SPECIAL CIRCUMSTANCES", arrRaw2) = -1 Then
            Cells(labAdd("REFERRAL SOURCE", arrRaw2), 1).Select
            Selection.EntireRow.Insert
            Cells(labAdd("REFERRAL SOURCE", arrRaw2), 1) = "SPECIAL CIRCUMSTANCES"
            Cells(labAdd("REFERRAL SOURCE", arrRaw2), 2) = 0
        End If
        'MAPPING (required to do after every row insert)
            Dim arrRaw3(1 To 80) As String
            For i = 1 To 80
                arrRaw3(i) = Cells(i, 1)
            Next
    'PE CHECKBOX
            Cells(labAdd("PE FIRM", arrRaw2), 1).Select
            Selection.EntireRow.Insert
            Cells(labAdd("PE FIRM", arrRaw2), 1) = "PE"
            If IsEmpty(Cells(labAdd("PE FIRM", arrRaw2) + 1, 2)) = False Then
                Cells(labAdd("PE FIRM", arrRaw2), 2) = "TRUE"
            Else
                Cells(labAdd("PE FIRM", arrRaw2), 2) = "FALSE"
            End If
        'MAPPING (required to do after every row insert) IT'S NOT NECESSARY BUT IT'S A REMNANT FROM PREVIOUS VERSION (DELETING IT WOULD REQUIRE CHANGING ARRAY REFERENCES)
            Dim arrRaw4(1 To 80) As String
            For i = 1 To 80
                arrRaw4(i) = Cells(i, 1)
            Next
'ADJUSTING CLIENT STATUS
    Dim cliStaFilled As Boolean
    If LCase(Cells(labAdd("CLIENT STATUS", arrRaw6), 2)) Like "*active*" Then
        cliStaFilled = True
        Cells(labAdd("CLIENT STATUS", arrRaw6), 2) = "Active"
    ElseIf LCase(Cells(labAdd("CLIENT STATUS", arrRaw6), 2)) Like "*never*" Then
        cliStaFilled = True
        Cells(labAdd("CLIENT STATUS", arrRaw6), 2) = "Never"
    ElseIf LCase(Cells(labAdd("CLIENT STATUS", arrRaw6), 2)) Like "*past*" Then
        cliStaFilled = True
        Cells(labAdd("CLIENT STATUS", arrRaw6), 2) = "Past"
    ElseIf LCase(Cells(labAdd("CLIENT STATUS", arrRaw6), 2)) Like "*y*" Then
        cliStaFilled = True
        Cells(labAdd("CLIENT STATUS", arrRaw6), 2) = "Active"
    ElseIf LCase(Cells(labAdd("CLIENT STATUS", arrRaw6), 2)) Like "*n*" Then
        cliStaFilled = True
        Cells(labAdd("CLIENT STATUS", arrRaw6), 2) = "Never"
    Else
        Cells(labAdd("CLIENT STATUS", arrRaw6), 2) = Cells(labAdd("CLIENT STATUS", arrRaw6), 2)
    End If
'ADJUSTING PRIORITY
    If LCase(Cells(labAdd("PRIORITY", arrRaw6), 2)) Like "*1*" Then
        Cells(labAdd("PRIORITY", arrRaw6), 2) = "1"
    End If
    If LCase(Cells(labAdd("PRIORITY", arrRaw6), 2)) Like "*2*" Then
        Cells(labAdd("PRIORITY", arrRaw6), 2) = "2"
    End If
    If LCase(Cells(labAdd("PRIORITY", arrRaw6), 2)) Like "*3*" Then
        Cells(labAdd("PRIORITY", arrRaw6), 2) = "3"
    End If
'ADJUSTING TITLE LEVEL
    If LCase(Cells(labAdd("Title Level", arrRaw6), 2)) Like "*board*" Then
        Cells(labAdd("Title Level", arrRaw6), 2) = "Board/Advisor"
    End If
    If LCase(Cells(labAdd("Title Level", arrRaw6), 2)) Like "*advisor*" Then
        Cells(labAdd("Title Level", arrRaw6), 2) = "Board/Advisor"
    End If
    If LCase(Cells(labAdd("Title Level", arrRaw6), 2)) Like "*ceo*" Then
        Cells(labAdd("Title Level", arrRaw6), 2) = "CEO"
    End If
    If LCase(Cells(labAdd("Title Level", arrRaw6), 2)) Like "*cxo*" Then
        Cells(labAdd("Title Level", arrRaw6), 2) = "CXO"
    End If
        If LCase(Cells(labAdd("Title Level", arrRaw6), 2)) Like "*n-1*" Then
        Cells(labAdd("Title Level", arrRaw6), 2) = "N-1"
    End If
    If LCase(Cells(labAdd("Title Level", arrRaw6), 2)) Like "*n-2*" Then
        Cells(labAdd("Title Level", arrRaw6), 2) = "N-2"
    End If
    If LCase(Cells(labAdd("Title Level", arrRaw6), 2)) Like "*other*" Then
        Cells(labAdd("Title Level", arrRaw6), 2) = "Other"
    End If
'ADJUSTING SPECIAL CIRCUMSTANCES
    If LCase(Cells(labAdd("SPECIAL CIRCUMSTANCES", arrRaw6), 2)) Like "*y*" Then
        Cells(labAdd("SPECIAL CIRCUMSTANCES", arrRaw6), 2) = 1
    ElseIf Cells(labAdd("SPECIAL CIRCUMSTANCES", arrRaw6), 2) Like "*1*" Then
        Cells(labAdd("SPECIAL CIRCUMSTANCES", arrRaw6), 2) = 1
    ElseIf Cells(labAdd("SPECIAL CIRCUMSTANCES", arrRaw6), 2) = 1 Then
        Cells(labAdd("SPECIAL CIRCUMSTANCES", arrRaw6), 2) = 1
    Else
        Cells(labAdd("SPECIAL CIRCUMSTANCES", arrRaw6), 2) = 0
    End If
'ADJUSTING EXEC SEARCH FIRMS
    If LCase(Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2)) Like "*ez*" Then
        Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2) = "EZ"
    ElseIf LCase(Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2)) Like "*h&s*" Then
        Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2) = "H&S"
    ElseIf LCase(Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2)) Like "*h and*" Then
        Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2) = "H&S"
    ElseIf LCase(Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2)) Like "*kf*" Then
        Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2) = "KF"
    ElseIf LCase(Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2)) Like "*rr*" Then
        Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2) = "RR"
    ElseIf LCase(Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2)) Like "*ss*" Then
        Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2) = "SS"
    ElseIf LCase(Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2)) = "" Then
        Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2) = ""
    ElseIf LCase(Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2)) = " " Then
        Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2) = ""
    Else
        Cells(labAdd("EXEC SEARCH FIRM", arrRaw6), 2) = "Other"
    End If
'NAMES SWAPPING
        Dim str As String
            'copying referring partner value to the variable to use it in the fee section later on
                str = Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 2)
        Dim nameSwap As Boolean
        If Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 2) <> "" Then
            Range(Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 9), Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 13)) = Split(Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 2))
            'TURNING N/A's TO BLANKS
                For i = 9 To 13
                    Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), i) = WorksheetFunction.IfError(Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), i), "")
                Next
            'IF NAME HAS TWO PARTS, DO THE NAME SWAP
                If Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 11) = "" Then
                    nameSwap = True
                    Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 2) = Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 10) & ", " & Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 9)
                End If
        End If
'DELETING NAME-SWAP DRAF AREA
        Range(Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 9), Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 13)).Select
        Selection.Clear
'ENGAGED PARTNER - ADDING FMNOs OF FREQUENT PARTNERS
        If souPar = True Then
            Cells(labAdd("PRIMARY ENGAGED PARTNER", arrRaw6), 2) = Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 2)
        End If
    'NAMES SWAPPING
        If souPar = False Then
            Dim str2 As String
            Dim nameSwap2 As Boolean
            engParLoc = labAdd("PRIMARY ENGAGED PARTNER", arrRaw6)
            If IsEmpty(Cells(engParLoc, 2)) = False Then
                Range(Cells(engParLoc, 9), Cells(engParLoc, 13)) = Split(Cells(engParLoc, 2))
                'TURNING N/A's TO BLANKS
                    For i = 9 To 13
                        Cells(engParLoc, i) = WorksheetFunction.IfError(Cells(engParLoc, i), "")
                    Next
                'IF NAME HAS TWO PARTS, DO THE NAME SWAP
                    If Cells(engParLoc, 11) = "" Then
                        nameSwap2 = True
                        Cells(engParLoc, 2) = Cells(engParLoc, 10) & ", " & Cells(engParLoc, 9)
                    End If
            End If
        End If
'ADJUSTING PRIMARY INDUSTRY
    If LCase(Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2)) Like "*advertising*" Then
        Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2) = "Advertising"
    End If
    If LCase(Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2)) Like "*aero*" Then
        Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2) = "Aerospace"
    End If
    If LCase(Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2)) Like "*agri*" Then
        Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2) = "Agriculture"
    End If
    If LCase(Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2)) Like "*asset*" Then
        Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2) = "Asset management"
    End If
    If LCase(Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2)) Like "*auto*" Then
        Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2) = "Automotive"
    End If
    If LCase(Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2)) Like "*banking*" Then
        Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2) = "Banking"
    End If
    If LCase(Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2)) Like "*bus*" Or LCase(Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2)) Like "*consumer s*" Then
        Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2) = "Business and consumer services"
    End If
    If LCase(Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2)) Like "*chem*" Then
        Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2) = "Chemicals"
    End If
    If LCase(Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2)) Like "*clean*" Then
        Cells(labAdd("PRIMARY INDUSTRY", arrRaw6), 2) = "Clean Tech"
    End If
'ADJUSTING PRIMARY FUNCTION
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*admin*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Administration"
    End If
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*board*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Board Member"
    End If
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*com*" Or LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*publish*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Comms & Publishing"
    End If
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*consul*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Consulting"
    End If
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*business dev*" Or LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*bus dev*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Corp & Bus Dev"
    End If
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*edu*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Education"
    End If
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*engine*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Engineering"
    End If
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*entre*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Entrepreneurs"
    End If
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*exec*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Executive leadership"
    End If
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*finance*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Finance"
    End If
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*financi*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Financial services-related"
    End If
    If LCase(Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2)) Like "*govern*" Then
        Cells(labAdd("PRIMARY FUNCTION", arrRaw6), 2) = "Government"
    End If
'COLORING REFERRAL SOURCE SECTION
    Range(Cells(1, 1), Cells(1, 2)).Select
    Selection.Style = "Accent5"
        'LABELS
            Range(Cells(2, 1), Cells(labAdd("MANAGER", arrRaw6), 1)).Select
            Selection.Style = "40% - Accent5"
        'VALUES
            Range(Cells(2, 2), Cells(labAdd("MANAGER", arrRaw6), 2)).Select
            Selection.Style = "20% - Accent5"
'GRABBING THE ACTIVITIES
    actToDatLoc = labAdd("6. ACTIVITY TO DATE", arrRaw6)
    feeLoc = labAdd("7. FEE INFORMATION", arrRaw6)
    'beginning of the activities
        For i = actToDatLoc To feeLoc - 1
            If IsEmpty(Cells(i, 2)) = False Then
                actStaLoc = i
                Exit For
            End If
        Next
    'end of the activities
        For i = actStaLoc To feeLoc - 1
            If IsEmpty(Cells(i, 2)) = True Then
                actEndLoc = i - 1
                Exit For
            End If
        Next
'ADJUSTING ACTIVITIES
    Dim R2J As Boolean
    For i = actStaLoc To actEndLoc
        If LCase(Cells(i, 2)) Like "*int*" And LCase(Cells(i, 2)) Like "*call*" Or LCase(Cells(i, 2)) Like "*intake*" Then
            Cells(i, 2) = "Intake Call"
        ElseIf LCase(Cells(i, 2)) Like "*advice*" Then
            Cells(i, 2) = "Advice on Resume"
        ElseIf LCase(Cells(i, 2)) Like "*party*" Or LCase(Cells(i, 2)) Like "*expert*" Then
            Cells(i, 2) = "3rd Party Resume Expert"
        End If
    Next
    
'FEE
    'ADDING MISSING ROWS
        Cells(labAdd("7. FEE INFORMATION", arrRaw6), 1).Select
        Selection.EntireRow.Insert
        Selection.EntireRow.Insert
        Selection.EntireRow.Insert
        Selection.EntireRow.Insert
        Selection.EntireRow.Insert
        Selection.EntireRow.Insert
            'I would normally do the mapping after every raw insert, but here it would require changing names of all the subsequent arrays so noth worth it
            'JOB ID
                Cells(labAdd("7. FEE INFORMATION", arrRaw6), 1).Select
                Selection.value = "CANDIDATE ID"
                Cells(labAdd("7. FEE INFORMATION", arrRaw6), 2) = Cells(1, 2)
            'FEE POTENTIAL
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 1, 1).Select
                Selection.value = "FEE POTENTIAL"
            'FEE RATIONALE
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 2, 1).Select
                Selection.value = "FEE RATIONALE"
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 2, 2) = Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 7, 2)
            'FEE ASK
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 3, 1).Select
                Selection.value = "FEE ASK"
                If IsEmpty(Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 7, 2)) = True Then
                    Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 3, 2) = "FALSE"
                Else
                    Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 3, 2) = "TRUE" 'here you need to set an alert message in case there's a comment, for instance
                End If
            'CCODE RECEIVED
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 4, 1).Select
                Selection.value = "CCODE RECEIVED"
                If IsEmpty(Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 7, 2)) = True Then
                    Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 4, 2) = "FALSE"
                Else
                    Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 4, 2) = "TRUE" 'here you need to set an alert message in case there's a comment, for instance
                End If
            'DECLINE REASON
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 5, 1).Select
                Selection.value = "DECLINE REASON"
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 5, 2) = Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 8, 2)
            'AMOUNT
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 6, 1).Select
                Selection.value = "AMOUNT"
            'CCODE (it should be at the end but I'm copying now so that the partner field does not overwrite it)
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 8, 1).Select
                Selection.value = "CHARGE CODE"
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 8, 2) = Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 7, 2)
            'REFERRING/APPROVING PARTNER
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 7, 1).Select
                Selection.value = "REFERRING/APPROVING PARTNER"
                Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 7, 2) = str
                'Cells(labAdd("7. FEE INFORMATION", arrRaw6) + 7, 2) = Cells(labAdd("PRIMARY REFEREE'S NAME", arrRaw6), 2)
        'MAPPING (required to do after every row insert)
            Dim arrRaw7(1 To 80) As String
            For i = 1 To 80
                arrRaw7(i) = Cells(i, 1)
            Next
        
        'MAPPING (required to do after every row insert)
            Dim arrRaw8(1 To 80) As String
            For i = 1 To 80
                arrRaw8(i) = Cells(i, 1)
            Next
    'ADJUSTING FEE ASK / COMMENTS (often, there is a comment in the T/F fee ask field; below checks for the comment, throws it to the comment field, and leaves T/F)
        Dim feeAskComment As String
        Dim feeAskCommentBool As Boolean
        If Len(Cells(labAdd("FEE ASK", arrRaw8), 2)) > 3 And InStr(1, (Cells(labAdd("FEE ASK", arrRaw8), 2)), "Y") = 1 Then '2nd condition checks if the string starts with y or n
            feeAskCommentBool = True
            Range(Cells(labAdd("FEE ASK", arrRaw8), 3), Cells(labAdd("FEE ASK", arrRaw8), 15)) = Split(Cells(labAdd("FEE ASK", arrRaw8), 2), " ")
                'TURNING N/A's TO BLANKS
                    For i = 3 To 15
                        Cells(labAdd("FEE ASK", arrRaw8), i) = WorksheetFunction.IfError(Cells(labAdd("FEE ASK", arrRaw8), i), "")
                    Next
                'THROWING SPLITTED CELLS INTO THE COMMENT VARIABLE (would use the len function, but not sure how to simply prevent throwing cut yes; this solution seems safer, though longer)
                    For i = 3 To 15
                        feeAskComment = feeAskComment + Cells(labAdd("FEE ASK", arrRaw8), i) + " "
                    Next
                'CLEARING THE SPLITTING DRAFT AREA
                    Range(Cells(labAdd("FEE ASK", arrRaw8), 3), Cells(labAdd("FEE ASK", arrRaw8), 15)).Clear
            Cells(labAdd("FEE ASK", arrRaw8), 2) = "Y"
            Cells(labAdd("COMMENTS", arrRaw8), 2) = feeAskComment
        ElseIf Len(Cells(labAdd("FEE ASK", arrRaw8), 2)) > 3 And InStr(1, (Cells(labAdd("FEE ASK", arrRaw8), 2)), "n") = 1 Then '2nd condition checks if the string starts with y or n
            feeAskCommentBool = True
            Range(Cells(labAdd("FEE ASK", arrRaw8), 3), Cells(labAdd("FEE ASK", arrRaw8), 15)) = Split(Cells(labAdd("FEE ASK", arrRaw8), 2), " ")
                'TURNING N/A's TO BLANKS
                    For i = 3 To 15
                        Cells(labAdd("FEE ASK", arrRaw8), i) = WorksheetFunction.IfError(Cells(labAdd("FEE ASK", arrRaw8), i), "")
                    Next
                'THROWING SPLITTED CELLS INTO THE COMMENT VARIABLE (would use the len function, but not sure how to simply prevent throwing cut yes; this solution seems safer, though longer)
                    For i = 5 To 15
                        feeAskComment = feeAskComment + Cells(labAdd("FEE ASK", arrRaw8), i) + " "
                    Next
                'CLEARING THE SPLITTING DRAFT AREA
                    Range(Cells(labAdd("FEE ASK", arrRaw8), 3), Cells(labAdd("FEE ASK", arrRaw8), 15)).Clear
            Cells(labAdd("FEE ASK", arrRaw8), 2) = "N"
            Cells(labAdd("COMMENTS", arrRaw8), 2) = feeAskComment
        End If
        
        'Cells(labAdd("FEE ASK", arrRaw8), 3) = Len(Cells(labAdd("FEE ASK", arrRaw8), 2))
        
        
    'ADJUSTING FEE RATIONALE
        If Cells(labAdd("FEE RATIONALE", arrRaw8), 2) = "" Or LCase(Cells(labAdd("FEE RATIONALE", arrRaw8), 2)) Like "*log*" Then
            Cells(labAdd("FEE RATIONALE", arrRaw8), 2) = "No - Log Only"
        ElseIf LCase(Cells(labAdd("FEE RATIONALE", arrRaw8), 2)) Like "*internal*" Then
            Cells(labAdd("FEE RATIONALE", arrRaw8), 2) = "No - Internal"
        ElseIf LCase(Cells(labAdd("FEE RATIONALE", arrRaw8), 2)) Like "*exec*" Then
            Cells(labAdd("FEE RATIONALE", arrRaw8), 2) = "No - Executive Search"
        ElseIf LCase(Cells(labAdd("FEE RATIONALE", arrRaw8), 2)) Like "*other*" Then
            Cells(labAdd("FEE RATIONALE", arrRaw8), 2) = "No - Other"
        End If
            'CHECKING FEE POTENTIAL
                If Cells(labAdd("FEE RATIONALE", arrRaw8), 2) Like "*Yes*" Then
                    Cells(labAdd("FEE POTENTIAL", arrRaw7), 2) = "TRUE"
                Else
                    Cells(labAdd("FEE POTENTIAL", arrRaw7), 2) = "FALSE"
                End If
                
    'ADJUSTING FEE ASK
        If LCase(Cells(labAdd("FEE ASK", arrRaw8), 2)) = "yes" Or LCase(Cells(labAdd("FEE ASK", arrRaw8), 2)) = "y" Then
            Cells(labAdd("FEE ASK", arrRaw8), 2) = 1
        ElseIf LCase(Cells(labAdd("FEE ASK", arrRaw8), 2)) = "no" Or LCase(Cells(labAdd("FEE ASK", arrRaw8), 2)) = "n" Then
            Cells(labAdd("FEE ASK", arrRaw8), 2) = 0
        Else
            Cells(labAdd("FEE ASK", arrRaw8), 2) = Cells(labAdd("FEE ASK", arrRaw8), 2)
        End If
    'ADJUSTING DECLINE REASON
        If LCase(Cells(labAdd("DECLINE REASON", arrRaw8), 2)) Like "*reply*" Then
            Cells(labAdd("DECLINE REASON", arrRaw8), 2) = "No reply"
        ElseIf LCase(Cells(labAdd("DECLINE REASON", arrRaw8), 2)) Like "*oppose*" Then
            Cells(labAdd("DECLINE REASON", arrRaw8), 2) = "Opposed to fees"
        ElseIf LCase(Cells(labAdd("DECLINE REASON", arrRaw8), 2)) Like "*insufficient*" Then
            Cells(labAdd("DECLINE REASON", arrRaw8), 2) = "Insufficient matches"
        ElseIf LCase(Cells(labAdd("DECLINE REASON", arrRaw8), 2)) Like "*wip*" Or LCase(Cells(labAdd("DECLINE REASON", arrRaw8), 2)) Like "*red*" Then
            Cells(labAdd("DECLINE REASON", arrRaw8), 2) = "WIP in the red"
        ElseIf LCase(Cells(labAdd("DECLINE REASON", arrRaw8), 2)) Like "*hold*" Then
            Cells(labAdd("DECLINE REASON", arrRaw8), 2) = "On hold"
        ElseIf LCase(Cells(labAdd("DECLINE REASON", arrRaw8), 2)) Like "*volume*" Or LCase(Cells(labAdd("DECLINE REASON", arrRaw8), 2)) Like "*discount*" Then
            Cells(labAdd("DECLINE REASON", arrRaw8), 2) = "Volume Discount"
        Else
            Cells(labAdd("DECLINE REASON", arrRaw8), 2) = ""
        End If
'changing Ys to TRUEs etc.
    For i = 1 To 100
        If LCase(Cells(i, 2)) = "y" Then
            Cells(i, 2) = "TRUE"
        ElseIf LCase(Cells(i, 2)) = "n" Then
            Cells(i, 2) = "FALSE"
        End If
    Next
'FORMATTING COLUMNS' WIDTH
    Range("A:A").Select
    Selection.ColumnWidth = 35
    Range("B:AK").Select
    Selection.ColumnWidth = 15
    Selection.HorizontalAlignment = xlLeft

'PASTING THE SECTIONS TO HORIZONTAL POSITIONS
    'ACTIVITIES
        Range(Cells(labAdd("6. ACTIVITY TO DATE", arrRaw8), 1), Cells(actEndLoc, 2)).Select
        Selection.Cut
        Cells(1, 4).Select
        ActiveSheet.Paste
        'MAPPING ACTIVITIES (needed for conditional on 'resume to job'/'resume share')
            Dim arrActLab(1 To 25) As String
            For i = 1 To 25
                arrActLab(i) = Cells(i, 4)
            Next
    'FEE
        Range(Cells(labAdd("CANDIDATE ID", arrRaw8), 1), Cells(labAdd("CHARGE CODE", arrRaw8), 2)).Select
        Selection.Cut
        Cells(1, 7).Select
        ActiveSheet.Paste

'PREVENTING CELLS' CONTENT TO RUN OVER TO EMPTY CELLS
    ManLoc = labAdd(" MANAGER", arrRaw8)
    For i = 1 To ManLoc
        Cells(i, 3) = " "
    Next
    For i = 1 To ManLoc
        Cells(i, 6) = " "
    Next
    For i = 1 To ManLoc
        Cells(i, 9) = " "
    Next
'PASTING TRANSPOSE
    'REFERRAL SOURCE
        Range(Cells(1, 1), Cells(ManLoc, 2)).Select
        Selection.Copy
        Range("A" & ManLoc + 2).Select
        Selection.PasteSpecial Paste:=xlPasteAll, Transpose:=True
        'CORRECTION OF COLORING
            'LABELS
                Range(Cells(ManLoc + 2, 1), Cells(ManLoc + 2, ManLoc)).Select
                Selection.Style = "Accent5"
            'VALUES
                Range(Cells(ManLoc + 3, 1), Cells(ManLoc + 3, ManLoc)).Select
                Selection.Style = "20% - Accent5"
        'SHADOWING CELLS
            'client status
            Cells(ManLoc + 3, 5).Select
            Selection.Style = "Input"
            'ref source
            Cells(ManLoc + 3, 12).Select
            Selection.Style = "Input"
            'priRef to engParFMNO
            Range(Cells(ManLoc + 3, 16), Cells(ManLoc + 3, 20)).Select
            Selection.Style = "Input"
            'target region
            Cells(ManLoc + 3, 26).Select
            Selection.Style = "Input"
            ' MANAGER
            Cells(ManLoc + 3, ManLoc).Select
            Selection.Style = "Input"
        'COMMENTS / ALERTS
            'CLIENT STATUS
            cliStaLoc = labAdd("CLIENT STATUS", arrRaw7)
                'ALERT
                    If cliStaFilled = False Then
                        Cells(ManLoc + 4, cliStaLoc) = "Check comment"
                        Cells(ManLoc + 4, cliStaLoc).Select
                        Selection.Style = "Bad"
                'COMMENTS
                        Cells(ManLoc + 3, cliStaLoc).Select
                        Selection.AddComment
                        Selection.Comment.Visible = False
                        Selection.Comment.Text Text:="Viable options are 'Active', 'Past', 'Never'. If unclear, please check our systems."
                    End If
            'PRIMARY REFEREE'S NAME
                If nameSwap = False And ppb = False Then
                    'ALERT
                        Cells(ManLoc + 4, labAdd("PRIMARY REFEREE'S NAME", arrRaw7)) = "Check comment"
                        Cells(ManLoc + 4, labAdd("PRIMARY REFEREE'S NAME", arrRaw7)).Select
                        Selection.Style = "Bad"
                    'COMMENTS
                        Cells(ManLoc + 3, labAdd("PRIMARY REFEREE'S NAME", arrRaw7)).Select
                        Selection.AddComment
                        Selection.Comment.Visible = False
                        Selection.Comment.Text Text:="LastName, FirstName format required."
                End If
            'PRIMARY REFEREE'S EMAIL
                priRefEmaLoc = labAdd("PRIMARY REFEREE'S EMAIL ADDRESS", arrRaw7)
                '
                    If nameSwap = False And is = True And IsEmpty(Cells(ManLoc + 3, priRefEmaLoc)) = True Or IsNull(Cells(ManLoc + 3, priRefEmaLoc)) = True Then
                        'ALERT
                            Cells(ManLoc + 4, priRefEmaLoc) = "Email needed"
                            Cells(ManLoc + 4, priRefEmaLoc).Select
                            Selection.Style = "Bad"
                    End If
                'NON-
                    If nameSwap = False And is = False And ppb = False Then
                        'ALERT
                            Cells(ManLoc + 4, priRefEmaLoc) = "Check comment"
                            Cells(ManLoc + 4, priRefEmaLoc).Select
                            Selection.Style = "Bad"
                        'COMMENTS
                            Cells(ManLoc + 3, priRefEmaLoc).Select
                            Selection.AddComment
                            Selection.Comment.Visible = False
                            Selection.Comment.Text Text:="If available, enter email."
                    End If
            'TARGET REGION
                If tarReg = True Then
                    'ALERT
                        Cells(ManLoc + 4, tarRegLoc) = "Check comment"
                        Cells(ManLoc + 4, tarRegLoc).Select
                        Selection.Style = "Bad"
                    'COMMENTS
                        Cells(ManLoc + 3, tarRegLoc).Select
                        Selection.AddComment
                        Selection.Comment.Visible = False
                        Selection.Comment.Text Text:="Please do manually." & Chr(10) & "This is is a multi-select" & Chr(10) & "checkbox field in the db."
                End If
            ' MANAGER
                If Cells(ManLoc + 3, ManLoc) = True Then
                    'ALERT
                        Cells(ManLoc + 4, ManLoc) = "Check comment"
                        Cells(ManLoc + 4, ManLoc).Select
                        Selection.Style = "Bad"
                    'COMMENTS
                        Cells(ManLoc + 3, ManLoc).Select
                        Selection.AddComment
                        Selection.Comment.Visible = False
                        Selection.Comment.Text Text:="Please double check." & Chr(10) & "Sometimes managers put" & Chr(10) & "Y automatically, even when" & Chr(10) & "there are no call notes" & Chr(10) & "in the submission email."
                        'Selection.Comment.Shape.TextFrame.AutoSize = True
                End If
        'AUTO-SIZING COMMENTS
            Dim XComment As Comment
            For Each XComment In Application.ActiveSheet.Comments
                XComment.Shape.TextFrame.AutoSize = True
            Next
        'ALERTS COUNTER
            Dim alertsRef As Integer
            alertsRef = 0
            For i = 3 To ManLoc
                If Cells(ManLoc + 4, i) <> "" Then
                    alertsRef = alertsRef + 1
                End If
            Next
            Cells(ManLoc + 4, 2) = "# of alerts: " & alertsRef
            Cells(ManLoc + 4, 2).Select
            If alertsRef = 0 Then
                Selection.Style = "Good"
            Else
                Selection.Style = "Bad"
            End If
        'BUTTON
            Dim w As Worksheet
            Dim b1 As Button
            Dim t1 As Range
            
            Set t1 = ActiveSheet.Range(Cells(ManLoc + 4, 1), Cells(ManLoc + 4, 1))
            Set w = ActiveSheet
            Set b1 = w.buttons.Add(t1.Left, t1.Top, t1.Width, t1.Height)
            b1.OnAction = "copyRef"
            b1.Characters.Text = "Copy Referral Source"
    'ACTIVITIES
        'NEED TO MAP AGAIN
        'POPULATING THE ARRAY
            Dim arrAct(9 To 29) As String 'can't use variable in arrays - only constant expression
            For i = 9 To 29
                arrAct(i) = Cells(i, 5)
            Next
        'DETERMINING # OF ACTIVITIES (it'll be needed for making paste-ready rows of activities)
            actHorStaLoc = actStaLoc - labAdd("6. ACTIVITY TO DATE", arrRaw8) + 1
            actHorEndLoc = actEndLoc - labAdd("6. ACTIVITY TO DATE", arrRaw8) + 1
            numOfAct = actHorEndLoc - actHorStaLoc + 1
        'COLORING ACTIVITIES
            'HEADERS
                Range(Cells(1, 4), Cells(1, 5)).Select
                Selection.Style = "Accent5"
            'LABELS
                Range(Cells(2, 4), Cells(actHorEndLoc, 4)).Select
                Selection.Style = "40% - Accent5"
            'VALUES
                Range(Cells(2, 5), Cells(actHorEndLoc, 5)).Select
                Selection.Style = "20% - Accent5"
                
        'CREATING LABELS
            Cells(ManLoc + 5, 1) = "CANDIDATE ID - ACTIVITIES"
            Cells(ManLoc + 5, 2) = "DATE"
            Cells(ManLoc + 5, 3) = "ACTIVITY"
        'PASTING ACTIVITIES
            'we're starting counting from 0 because if we used 1, LBound(arrAct) would start from a second activity
            For i = 0 To numOfAct - 1 '-1 beacause we're starting counting from 0
                Cells(ManLoc + 6 + i, 1) = Cells(1, 2)
                Cells(ManLoc + 6 + i, 2) = DateTime.Date
                Cells(ManLoc + 6 + i, 3) = Cells(actHorStaLoc + i, 5)
            Next
        'CORRECTION OF COLORING
            'LABELS
                Range(Cells(ManLoc + 5, 1), Cells(ManLoc + 5, 3)).Select
                Selection.Style = "Accent5"
            'VALUES
                Range(Cells(ManLoc + 5 + 1, 1), Cells(ManLoc + 5 + numOfAct, 3)).Select
                Selection.Style = "20% - Accent5"
        'PREVENTING CELLS' VALUES TO SPILL TO NEIGHBOURING CELLS
            For i = ManLoc + 5 + 1 To ManLoc + 5 + 1 + numOfAct
                Cells(i, 4) = " "
            Next
        'ALERTS/COMMENTS
            'CHECKING IF THERE ARE SOME UNUSUAL VALUES
                For i = ManLoc + 6 To ManLoc + 6 + numOfAct - 1
                    If Cells(i, 3) <> "Intake Call" Then
                        If Cells(i, 3) <> "Advice on Resume" Then
                            If Cells(i, 3) <> "3rd Party Resume Expert" Then
                                If Cells(i, 3) <> "IMT Search" Then
                                    If Cells(i, 3) <> " Networking List" Then
                                        If Cells(i, 3) <> "BoardEx Search" Then
                                            If Cells(i, 3) <> "Target Co Email" Then
                                                If Cells(i, 3) <> "Partner Email" Then
                                                    If Cells(i, 3) <> "DCS Intro" Then
                                                        If Cells(i, 3) <> "Head Hunter Intro" Then
                                                            If Cells(i, 3) <> "Network Intro" Then
                                                                If Cells(i, 3) <> " Coaching" Then
                                                                    If Cells(i, 3) <> " Phone Confirmation" Then
                                                                        If Cells(i, 3) <> " Interview" Then
                                                                            If Cells(i, 3) <> "Job Idea" Then
                                                                                If Cells(i, 3) <> "Resume to Job" Then
                                                                                    If Cells(i, 3) <> "Screen Interview" Then
                                                                                        If Cells(i, 3) <> "Follow up engagement" Then
                                                                                            If Cells(i, 3) <> "1 Rd Interview" Then
                                                                                                If Cells(i, 3) <> "2 Rd Interview" Then
                                                                                                    If Cells(i, 3) <> "3 Rd Interview" Then
                                                                                                        If Cells(i, 3) <> "Offer" Then
                                                                                                            If Cells(i, 3) <> "Accept" Then
                                                                                                                If Cells(i, 3) <> "Decline" Then
                                                                                                                    If Cells(i, 3) <> "Closed" Then
                                                                                                                        Cells(i, 4) = "Please check"
                                                                                                                        Cells(i, 4).Select
                                                                                                                        Selection.Style = "Bad"
                                                                                                                    End If
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                'RESUME TO JOB COMMENTS AND ALERTS
                If R2J = True Then
                    R2JLoc = labAdd("Resume to Job", arrAct) - labAdd(" Partner outreach", arrActLab)
                    Cells(ManLoc + 6 + R2JLoc, 3).AddComment
                    Cells(ManLoc + 6 + R2JLoc, 3).Comment.Visible = False
                    Cells(ManLoc + 6 + R2JLoc, 3).Comment.Text Text:="Please check if" & Chr(10) & "the original value doens't" & Chr(10) & "contain names or IDs." & Chr(10) & "-->" & R2JValue
                    Cells(ManLoc + 6 + R2JLoc, 3).Select
                    Selection.Comment.Shape.TextFrame.AutoSize = True
                    Cells(ManLoc + 6 + R2JLoc, 4).Select
                    Selection.value = "Check comment"
                    Selection.Style = "Bad"
                End If
            'ALERTS COUNTER
                Dim alertsAct As Integer
                alertsAct = 0
                For i = ManLoc + 6 To ManLoc + 6 + numOfAct - 1
                    If Cells(i, 4) = "Please check" Or Cells(i, 4) = "Check comment" Then
                        alertsAct = alertsAct + 1
                    End If
                Next
                Cells(ManLoc + 5 + numOfAct + 1, 2) = "# of alerts: " & alertsAct
                Cells(ManLoc + 5 + numOfAct + 1, 2).Select
                If alertsAct = 0 Then
                    Selection.Style = "Good"
                Else
                    Selection.Style = "Bad"
                End If
        'BUTTON
            Dim b2 As Button
            Dim t2 As Range
            
            Set t2 = ActiveSheet.Range(Cells(ManLoc + 5 + numOfAct + 1, 1), Cells(ManLoc + 5 + numOfAct + 1, 1))
            Set w = ActiveSheet
            Set b2 = w.buttons.Add(t2.Left, t2.Top, t2.Width, t2.Height)
            b2.OnAction = "copyAct"
            b2.Characters.Text = "Copy Activities"
    'PASTING FEE
        'POPULATING THE ARRAY
            Dim arrFee(1 To 32) As String
            For i = 1 To 32
                arrFee(i) = Cells(i, 7)
            Next
        'COLORING FEE
            'HEADERS
                Range(Cells(1, 7), Cells(1, 8)).Select
                Selection.Style = "Accent5"
            'LABELS
                Range(Cells(2, 7), Cells(labAdd("CHARGE CODE", arrFee), 7)).Select
                Selection.Style = "40% - Accent5"
            'VALUES
                Range(Cells(2, 8), Cells(labAdd("CHARGE CODE", arrFee), 8)).Select
                Selection.Style = "20% - Accent5"
        'PASTING TRANSPOSE
            Range(Cells(LBound(arrFee), 7), Cells(labAdd("CHARGE CODE", arrFee), 8)).Select
            Selection.Copy
            Range("A" & ManLoc + 5 + numOfAct + 2).Select '(+ 5) below referral source
            Selection.PasteSpecial Paste:=xlPasteAll, Transpose:=True
            'CORRECTION OF COLORING
                'LABELS
                    Range(Cells(ManLoc + 5 + numOfAct + 2, 1), Cells(ManLoc + 5 + numOfAct + 2, labAdd("CHARGE CODE", arrFee))).Select
                    Selection.Style = "Accent5"
                'VALUES
                    Range(Cells(ManLoc + 5 + numOfAct + 3, 1), Cells(ManLoc + 5 + numOfAct + 3, labAdd("CHARGE CODE", arrFee))).Select
                    Selection.Style = "20% - Accent5"
                'PREVENTING CELLS VALUES FROM SPILLING TO NEIGHBOURING CELLS
                    Dim endOfFee As Integer
                    endOfFee = labAdd("CHARGE CODE", arrFee)
                    Cells(ManLoc + 5 + numOfAct + 3, endOfFee + 1) = " "
        'SHADOWING CELLS
            'locating
                'ccode
                Cells(ManLoc + 5 + numOfAct + 3, 9).Select
                Selection.Style = "Input"
        'COMMENTS/ALERTS
            feeTraValLoc = ManLoc + 5 + numOfAct + 3
            'CHARGE CODE
                If chaCod = True Then
                    'ALERT
                    Cells(feeTraValLoc + 1, labAdd("CHARGE CODE", arrFee)) = "Pls double check"
                    Cells(feeTraValLoc + 1, labAdd("CHARGE CODE", arrFee)).Select
                    Selection.Style = "Bad"
                End If
            'ALERTS COUNTER
                Dim alertsFee As Integer
                alertsFee = 0
                For i = 3 To labAdd("CHARGE CODE", arrFee)
                    If Cells(feeTraValLoc + 1, i) <> "" Then
                        alertsFee = alertsFee + 1
                    End If
                Next
                Cells(feeTraValLoc + 1, 2) = "# of alerts: " & alertsFee
                Cells(feeTraValLoc + 1, 2).Select
                If alertsFee = 0 Then
                    Selection.Style = "Good"
                Else
                    Selection.Style = "Bad"
                End If
        'BUTTON
            Dim b3 As Button
            Dim t3 As Range
            
            Set t3 = ActiveSheet.Range(Cells(ManLoc + 5 + numOfAct + 4, 1), Cells(ManLoc + 5 + numOfAct + 4, 1))
            Set w = ActiveSheet
            Set b3 = w.buttons.Add(t3.Left, t3.Top, t3.Width, t3.Height)
            b3.OnAction = "copFee"
            b3.Characters.Text = "Copy Fee Information"
        'GETTING RID OF THE COPYING BROKEN LINE
        Application.CutCopyMode = False
        ActiveWorkbook.RefreshAll
End Sub


