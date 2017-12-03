Attribute VB_Name = "GitHub_IMT_n"
Sub databasePull()

'==============================================================================================================================================
'= MANIPULATES RAW DATA FROM MICROSOFT ACCESS. SANITIZED DUE TO CONFIDENTIALITY CONCERNS                                                      =
'= SELECTED LIST OF OPERATIONS:                                                                                                               =
'= - determines end of data range and number of actual records                                                                                =
'= - creating arrays with header cells (trimmed) for easy reference using the custom Label Address function (please see the Fubctions module) =
'= - loops and if statements                                                                                                                  =
'= - using string, int, and boolean variables                                                                                                 =
'= - sorting, filtering, adding comments                                                                                                      =
'= - basic operations such as deleting, moving, hiding, and inserting columns                                                                 =
'==============================================================================================================================================

'TRIMMING THE HEADER
    For i = 1 To 100
        Trim (Cells(6, i))
    Next
'LIST SIZE / ENDOFRANGE
    endOfRange = Range("G" & Rows.Count).End(xlUp).Row
    listSize = endOfRange - 6
    heaLoc = 6
'MOVING THE JOB START DATE COLUMN TO THE BEGINNING
    Dim arrHea13(1 To 100) As String
    
    For i = 1 To 100
        arrHea13(i) = Cells(heaLoc, i)
    Next
    
    Cells(1, labAdd("Start Date", arrHea13)).EntireColumn.Select
    Selection.Cut
    Cells(1, labAdd("PPB Status", arrHea13)).EntireColumn.Select
    Selection.Insert shift:=xlToRight
'MOVING AND FORMATING SPECIAL CIRCUMSTANCES
    'LOCATING SPEC CIRC
        For i = 1 To 100
            If LCase(Cells(heaLoc, i)) Like "*special circumstances*" Then
                speCirLoc = i
                Exit For
            End If
        Next
    'RENAMING
        Cells(heaLoc, speCirLoc) = "Special Circumstances"
    'COPYING SPE CIR
        Cells(heaLoc, speCirLoc).Select
        Selection.EntireColumn.Cut
        Columns("E:E").Select
        Selection.Insert shift:=xlToRight
    'UPDATING LOCATION
        For i = 1 To 100
            If LCase(Cells(heaLoc, i)) Like "*special circumstances*" Then
                speCirLoc = i
                Exit For
            End If
        Next
    'CHANGIN 1s TO Ys AND 0s TO BLANKS
        'THE ACTUAL CHANGE
            For i = 7 To endOfRange
                If Cells(i, speCirLoc) = 1 Then
                    Cells(i, speCirLoc) = "Y"
                Else
                    Cells(i, speCirLoc) = ""
                End If
            Next
'POPULATING THE ARRAY
    Dim arrHea0(1 To 100) As String
    
    heaLoc = 6
    For i = 1 To 100
        arrHea0(i) = Cells(heaLoc, i)
    Next

'MOVING 'CONFIDENTIAL COMPANY NAME'
    Cells(heaLoc, labAdd("Confidential Company Name", arrHea0)).Select
    Selection.EntireColumn.Cut
    Cells(heaLoc, labAdd("Online Profile Link", arrHea0)).Select
    Selection.EntireColumn.Insert shift:=xlToRight
Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    Range("B6").Select
    Selection.Copy
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("B1:B4").Select
    Range("B4").Activate
    Application.CutCopyMode = False
    Selection.Cut
    Range("A1").Select
    ActiveSheet.Paste
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "New Job"
    Range("B6").Select
    Selection.AutoFilter
    Rows("6:6").Select
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
'POPULATING THE ARRAY
    Dim arrHea(1 To 100) As String
    
    heaLoc = 6
    For i = 1 To 100
        arrHea(i) = Cells(heaLoc, i)
    Next
'DELETING DISFAVOUR
    Cells(1, labAdd("Disfavor", arrHea)).Select
    Selection.EntireColumn.Delete shift:=xlToLeft
'ADJUSTING TITLE AND COMPANY COLUMNS (CHANGING NAMES AND WIDHT)
    Cells(heaLoc, labAdd("TitleNew", arrHea)) = "Title"
    Cells(heaLoc, labAdd("TitleNew", arrHea)).Select
    Selection.ColumnWidth = 20
    Cells(heaLoc, labAdd("CompanyNew", arrHea)) = "Company"
    Cells(heaLoc, labAdd("CompanyNew", arrHea)).Select
    Selection.ColumnWidth = 20
    Cells(heaLoc, labAdd("Confidential Company Name", arrHea)).Select
    Selection.ColumnWidth = 20
'statusLoc
    For i = 1 To 100
        If LCase(Cells(heaLoc, i)) = "status" Then
            statusLoc = i
            Exit For
        End If
    Next
'confiLoc
    For i = 1 To 100
        If LCase(Cells(heaLoc, i)) = "confidential" Then
            confiLoc = i
            Exit For
        End If
    Next
    
'flaggingStartLoc
    flaggingStartLoc = labAdd("Last", arrHea)
'flaggingEndLoc
    For i = 1 To 100
        If LCase(Cells(heaLoc, i)) = "transition flag" Then
            flaggingEndLoc = i
            Exit For
        End If
    Next
    
'dataEndLoc
    For i = heaLoc To endOfRange
        If Cells(i, flaggingStartLoc) = "" Then
            dataEndLoc = i - 1
            Exit For
        End If
    Next
    
'Status
    For i = heaLoc + 1 To endOfRange
        If Cells(i, statusLoc) = "Active" Then
            Range(Cells(i, flaggingStartLoc), Cells(i, flaggingEndLoc)).Select
            With Selection.Font
                .Color = -4165632
                .TintAndShade = 0
            End With
        End If
    Next
'Confi
    For i = heaLoc + 1 To endOfRange
        If Cells(i, confiLoc) = "Y" Then
            Range(Cells(i, flaggingStartLoc), Cells(i, flaggingEndLoc)).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
        End If
    Next
'DELETING STATUS AND CONFIDENTIAL COLUMNS (WARNING: NOT DYNAMIC)
    Range("D:D").Select
    Selection.EntireColumn.Delete
    Range("H:H").Select
    Selection.EntireColumn.Delete

'INSERTING THE TWO INITIAL FLAGGING COLUMNS
    Range("D:D").Select
    Selection.Insert shift:=xlToRight
    Range("D:D").Select
    Selection.Insert shift:=xlToRight
    
'DATE COMPARISON
    'defining header's location
    For i = 1 To 10
        If LCase(Cells(i, 1)) Like "*new job*" Then
            heaLoc = i
        End If
    Next
    
    'defining dates' locations
    For i = 1 To 100
        If LCase(Cells(heaLoc, i)) Like "*end date*" Then
            endDateLoc = i
        End If
    Next
    For i = 1 To 100
        If LCase(Cells(heaLoc, i)) Like "*start date*" Then
            startDateLoc = i
        End If
    Next
    For i = 1 To 100
        If LCase(Cells(heaLoc, i)) Like "*opt in date*" Then
            optInDateLoc = i
        End If
    Next
            
    For i = endOfRange To 7 Step -1
        If Cells(i, endDateLoc) = "" Then
            If Cells(i, startDateLoc) <> "" Then
                If Cells(i, startDateLoc) > Cells(i, optInDateLoc) Then
                    Cells(i, 1) = "Y"
                End If
            End If
        End If
        If Cells(i, 2) = "Y" Then
            Cells(i, 1) = "Y"
        End If
    Next
'sorting by Last Name
    'lastNameColumnLoc
    For i = 1 To 100
        If LCase(Cells(heaLoc, i)) = "last" Then
            lastNameColumnLoc = i
            Exit For
        End If
    Next
'FILTERING
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range( _
        Cells(heaLoc, lastNameColumnLoc), Cells(endOfRange, lastNameColumnLoc)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'LAST, FIRST TO FIRST LAST
    'CHECKING IF THERE ALREADY IS FULL NAME COLUMN
        Dim fulNamBol As Boolean
        For i = 1 To 50
            If Cells(heaLoc, i) = "Full Name" Then
                fulNamBol = True
                fulNamLoc = i
                Exit For
            End If
        Next
        
    'CONDITIONALS
        If fulNamBol = True Then
            'DOING THE FIRST NAME LAST NAME
            For j = heaLoc + 1 To endOfRange
                Cells(j, fulNamLoc) = Cells(j, fulNamLoc - 1) & " " & Cells(j, fulNamLoc - 2)
            Next
        Else
            'LOCATING CANDIDATE TYPE AND INSERTING COLUMN FOR FULL NAME
            For k = 1 To 50
                If Cells(heaLoc, k) = "Candidate Type" Then
                    canTypLoc = k
                    Exit For
                End If
            Next
            'INSERTING COLUMN FOR FULL NAME
            Range(Cells(1, canTypLoc), Cells(1, canTypLoc)).Select
            Selection.EntireColumn.Insert
            'NAMING THE NEWLY INSERTED COLUM
            Cells(heaLoc, canTypLoc) = "Full Name"
            'DOING THE FIRST NAME LAST NAME
            For i = heaLoc + 1 To endOfRange
                Cells(i, canTypLoc) = Cells(i, canTypLoc - 1) & " " & Cells(i, canTypLoc - 2)
            Next
        End If
'PRIMARY FLAG 0 = FMR Y
    'POPULATING THE ARRAY (most of the macro still uses variables from loops)
        Dim arrHea2(1 To 100) As String 'on 20170725 the header had 65 colums (including 'conf comp nam' and 'special circ'
        
        'header in raw dataset sits in 6th row
        For i = 1 To 100
            arrHea2(i) = Cells(heaLoc, i)
        Next
        
    For i = heaLoc + 1 To endOfRange
        If Cells(i, labAdd("PRIMARY_FLAG_1", arrHea2)) = 0 Then
            Cells(i, labAdd("Former Role", arrHea2)) = "Y"
        End If
    Next
    'GETTING RID OF PRIMARY FLAG COLUMN
        Cells(1, labAdd("PRIMARY_FLAG_1", arrHea2)).Select
        Selection.EntireColumn.Delete

'ADDRESS COLUMS
    Dim arrHea3(1 To 100) As String
    Dim arrHea4(1 To 100) As String
    Dim arrHea5(1 To 100) As String
    Dim arrHea6(1 To 100) As String
    Dim arrHea7(1 To 100) As String
    Dim arrHea8(1 To 100) As String
    Dim arrHea9(1 To 100) As String
    Dim arrHea10(1 To 100) As String
    Dim arrHea11(1 To 100) As String
    
    For i = 1 To 100
        arrHea3(i) = Cells(heaLoc, i)
    Next
    
    Cells(heaLoc, labAdd("Willing to relocate", arrHea3)).Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Cells(heaLoc, labAdd("Willing to relocate", arrHea3) + 4).Select
    Selection.value = "Work Address at the End"
    Cells(heaLoc, labAdd("Willing to relocate", arrHea3) + 3).Select
    Selection.value = "Continent"
    Cells(heaLoc, labAdd("Willing to relocate", arrHea3) + 2).Select
    Selection.value = "Country"
    Cells(heaLoc, labAdd("Willing to relocate", arrHea3) + 1).Select
    Selection.value = "State"
    Cells(heaLoc, labAdd("Willing to relocate", arrHea3)).Select
    Selection.value = "City"
    
    For i = 1 To 100
        arrHea4(i) = Cells(heaLoc, i)
    Next
    
'ADDING A COMMENT TO HOME ADDRESS COLUMN
    Range(Cells(heaLoc, labAdd("Work Address at the End", arrHea4)), Cells(heaLoc, labAdd("Work Address at the End", arrHea4))).Select
    Selection.AddComment
    Selection.Comment.Visible = False
    Selection.Comment.Text Text:="If Y, there is also work address" & Chr(10) & "at the end of the spreadsheet"
    Selection.Comment.Shape.TextFrame.AutoSize = True

'TRIMMING THE ADDRESSES
    For i = heaLoc + 1 To endOfRange
        Cells(i, labAdd("Home City", arrHea4)) = Trim(Cells(i, labAdd("Home City", arrHea4)))
        Cells(i, labAdd("Home State", arrHea4)) = Trim(Cells(i, labAdd("Home State", arrHea4)))
        Cells(i, labAdd("Home Country", arrHea4)) = Trim(Cells(i, labAdd("Home Country", arrHea4)))
        Cells(i, labAdd("Home Continent", arrHea4)) = Trim(Cells(i, labAdd("Home Continent", arrHea4)))
        Cells(i, labAdd("Work City", arrHea4)) = Trim(Cells(i, labAdd("Work City", arrHea4)))
        Cells(i, labAdd("Work State", arrHea4)) = Trim(Cells(i, labAdd("Work State", arrHea4)))
        Cells(i, labAdd("Work Country", arrHea4)) = Trim(Cells(i, labAdd("Work Country", arrHea4)))
        Cells(i, labAdd("Work Continent", arrHea4)) = Trim(Cells(i, labAdd("Work Continent", arrHea4)))
    Next

'THE MAIN PART
    
    For i = heaLoc + 1 To endOfRange
        If IsEmpty(Cells(i, labAdd("Home City", arrHea4))) = True And IsEmpty(Cells(i, labAdd("Home State", arrHea4))) = True And IsEmpty(Cells(i, labAdd("Home Country", arrHea4))) = True And IsEmpty(Cells(i, labAdd("Home Continent", arrHea4))) = True Then
            Cells(i, labAdd("Work Address at the End", arrHea4)) = " " '!proceeding cell running over
            Cells(i, labAdd("Continent", arrHea4)) = Cells(i, labAdd("Work Continent", arrHea4))
                If IsEmpty(Cells(i, labAdd("Continent", arrHea4))) = True Then '!proceeding cell running over
                    Cells(i, labAdd("Continent", arrHea4)) = " "
                End If
            Cells(i, labAdd("Country", arrHea4)) = Cells(i, labAdd("Work Country", arrHea4))
                If IsEmpty(Cells(i, labAdd("Country", arrHea4))) = True Then '!proceeding cell running over
                    Cells(i, labAdd("Country", arrHea4)) = " "
                End If
            Cells(i, labAdd("State", arrHea4)) = Cells(i, labAdd("Work State", arrHea4))
                If IsEmpty(Cells(i, labAdd("State", arrHea4))) = True Then '!proceeding cell running over
                    Cells(i, labAdd("State", arrHea4)) = " "
                End If
            Cells(i, labAdd("City", arrHea4)) = Cells(i, labAdd("Work City", arrHea4))
                If IsEmpty(Cells(i, labAdd("City", arrHea4))) = True Then '!proceeding cell running over
                    Cells(i, labAdd("City", arrHea4)) = " "
                End If
        Else
            'checking if there is a work address as well
            If IsEmpty(Cells(i, labAdd("Work City", arrHea4))) = False Or IsEmpty(Cells(i, labAdd("Work State", arrHea4))) = False Or IsEmpty(Cells(i, labAdd("Work Country", arrHea4))) = False Or IsEmpty(Cells(i, labAdd("Work Continent", arrHea4))) = False Then
                Cells(i, labAdd("Work Address at the End", arrHea4)) = "Y"
            End If
            Cells(i, labAdd("Continent", arrHea4)) = Cells(i, labAdd("Home Continent", arrHea4))
                If IsEmpty(Cells(i, labAdd("Continent", arrHea4))) = True Then '!proceeding cell running over
                    Cells(i, labAdd("Continent", arrHea4)) = " "
                End If
            Cells(i, labAdd("Country", arrHea4)) = Cells(i, labAdd("Home Country", arrHea4))
                If IsEmpty(Cells(i, labAdd("Country", arrHea4))) = True Then '!proceeding cell running over
                    Cells(i, labAdd("Country", arrHea4)) = " "
                End If
            Cells(i, labAdd("State", arrHea4)) = Cells(i, labAdd("Home State", arrHea4))
                If IsEmpty(Cells(i, labAdd("State", arrHea4))) = True Then '!proceeding cell running over
                    Cells(i, labAdd("State", arrHea4)) = " "
                End If
            Cells(i, labAdd("City", arrHea4)) = Cells(i, labAdd("Home City", arrHea4))
                If IsEmpty(Cells(i, labAdd("City", arrHea4))) = True Then '!proceeding cell running over
                    Cells(i, labAdd("City", arrHea4)) = " "
                End If
        End If
    Next
    
'MOVING THE OLD ADDRESS COLUMNS TO THE END
    'can't figure out a simpler way to make sure that after switching columns order labAdd references the correct column than mapping the array every time, so here we go...
        'probably I could limit that to clusters of neighbouring columns but don't have time for that now
    Cells(1, labAdd("Work Continent", arrHea4)).EntireColumn.Select
    Selection.Cut
    Cells(1, labAdd("Transition Flag", arrHea4) + 1).EntireColumn.Select
    Selection.Insert shift:=xlToRight
    For i = 1 To 100
        arrHea5(i) = Cells(heaLoc, i)
    Next
    
    Cells(1, labAdd("Work Country", arrHea5)).EntireColumn.Select
    Selection.Cut
    Cells(1, labAdd("Transition Flag", arrHea5) + 1).EntireColumn.Select
    Selection.Insert shift:=xlToRight
    For i = 1 To 100
        arrHea6(i) = Cells(heaLoc, i)
    Next
    
    Cells(1, labAdd("Work State", arrHea6)).EntireColumn.Select
    Selection.Cut
    Cells(1, labAdd("Transition Flag", arrHea6) + 1).EntireColumn.Select
    Selection.Insert shift:=xlToRight
    For i = 1 To 100
        arrHea7(i) = Cells(heaLoc, i)
    Next
    
    Cells(1, labAdd("Work City", arrHea7)).EntireColumn.Select
    Selection.Cut
    Cells(1, labAdd("Transition Flag", arrHea7) + 1).EntireColumn.Select
    Selection.Insert shift:=xlToRight
    For i = 1 To 100
        arrHea8(i) = Cells(heaLoc, i)
    Next
    
    Cells(1, labAdd("Home Continent", arrHea8)).EntireColumn.Select
    Selection.Cut
    Cells(1, labAdd("Transition Flag", arrHea8) + 1).EntireColumn.Select
    Selection.Insert shift:=xlToRight
    For i = 1 To 100
        arrHea9(i) = Cells(heaLoc, i)
    Next
    
    Cells(1, labAdd("Home Country", arrHea9)).EntireColumn.Select
    Selection.Cut
    Cells(1, labAdd("Transition Flag", arrHea9) + 1).EntireColumn.Select
    Selection.Insert shift:=xlToRight
    For i = 1 To 100
        arrHea10(i) = Cells(heaLoc, i)
    Next
    
    Cells(1, labAdd("Home State", arrHea10)).EntireColumn.Select
    Selection.Cut
    Cells(1, labAdd("Transition Flag", arrHea10) + 1).EntireColumn.Select
    Selection.Insert shift:=xlToRight
    For i = 1 To 100
        arrHea11(i) = Cells(heaLoc, i)
    Next
    
    Cells(1, labAdd("Home City", arrHea11)).EntireColumn.Select
    Selection.Cut
    Cells(1, labAdd("Transition Flag", arrHea11) + 1).EntireColumn.Select
    Selection.Insert shift:=xlToRight
    'FILTERING ALL THE COLUMNS AGAIN
        Range(Cells(heaLoc, 1), Cells(heaLoc, 70)).Select
        Selection.AutoFilter
        Selection.AutoFilter
'CONDITIONAL FORMATTING FOR SPECIAL CIRCUMSTANCES
    speCirLoc = labAdd("Special Circumstances", arrHea11)
    Cells(1, speCirLoc).EntireColumn.Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Y"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16751204
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
'CONDITIONAL FORMATTING FOR FLAGGING COLUMNS
    flaggingColumnOneLoc = labAdd("PPB Status", arrHea11) + 1
    Cells(1, flaggingColumnOneLoc).EntireColumn.Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Y"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    flaggingColumnTwoLoc = flaggingColumnOneLoc + 1
    Cells(1, flaggingColumnTwoLoc).EntireColumn.Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Y"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
'CONDITIONAL FORMATTING FOR NEW JOB
    newJobLoc = labAdd("New Job", arrHea11)
    Cells(1, newJobLoc).EntireColumn.Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Y"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
'left align all
    Cells.Select
    Range("A1").Activate
    Selection.HorizontalAlignment = xlLeft
'THE COMMA FORMAT OF TENURE
    Dim arrHea14(1 To 100) As String
    For i = 1 To 100
        arrHea14(i) = Trim(Cells(heaLoc, i))
    Next
    Range(Cells(heaLoc + 1, labAdd("Tenure", arrHea14)), Cells(endOfRange, labAdd("Tenure", arrHea14))).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
End Sub


