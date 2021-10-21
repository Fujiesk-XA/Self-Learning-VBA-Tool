'Reformat SAP data information to SOP format for easy read and updates
Sub UPDATE1()

Application.ScreenUpdating = False
'Add new Sheets to raw data file
Dim y, x As Double
Dim e, checkFill As Boolean
Dim r, rr, rrr As Excel.Range
'Process Instruction and Process Parameters Variables
Dim longText As String, processStep As String, temp, letterCol As String
Dim lastRow As Double, processColumn As Integer, emptyColumn As Integer
Dim processPos As Integer, longPos As Integer, colonCount As Integer, semiCount As Integer, ProcessPosCol As Integer
Dim LArray() As String
Dim LText As String
Dim iii, iiii, iiiii, iiiiii As Integer
'Instructions and Parameters' variables
Dim sequenceNum As Integer, specCol As Integer, spec As String
Dim instOperationCol As Integer, parOperationCol As Integer, instOperation As String, parOperation As String
Dim instActionCol As Integer, parActionCol As Integer, instAction As String, parAction As String
Dim parameterCol As Integer, descriptionCol As Integer, parSequenceCol As Integer, instSequenceCol As Integer
Dim parameterDescription As String, instSequence As String, parSequence As String
Dim instructionRow As Integer
Dim found
'Used to determine dataentryform variables
Dim matchnumber As Integer
Dim planttxt As String

e = False

For y = 1 To Sheets.Count
    If Sheets(y).Name = "InstructionsAndParameters" Then
        e = True
    End If
Next

If e = False Then
Sheets.Add(, Sheets(y - 1)).Name = "InstructionsAndParameters"
End If

For y = 1 To 4

    If y = 1 Then
        Sheets("ProcessParameters").Select
    ElseIf y = 2 Then
        Sheets("ProcessInstructions").Select
    ElseIf y = 3 Then
        Sheets("WorkInstructions").Select
    ElseIf y = 4 Then
        Sheets("RecipeQuantities").Select
    End If
    
'FillinBlanks Macro

    checkFill = False
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set r = Range("A1:A" & lastRow)
    
    For Each Cell In r
        If Cell = "" Then
            checkFill = True
            Exit For
        End If
    Next

    If checkFill = True Then
        Cells(1, 1).Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$M" & lastRow).AutoFilter Field:=1, Criteria1:="="
        Range("A3:C3").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
        ActiveSheet.Range("$A$1:$M" & lastRow).AutoFilter Field:=1
        Columns("A:A").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.NumberFormat = "General"
        Columns("A:C").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=R[-1]C"
        Columns("A:C").Select
        Selection.copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("A1").Select
    End If
    
Next

'Format Recipe Quantities

    Sheets("RecipeQuantities").Select
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("A:A").Select
    Selection.copy
    Columns("BO:BO").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:= _
        False, Transpose:=False
    Range("BP2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/2"
    Range("BP2").Select
    Selection.AutoFill Destination:=Range("BP2:BP" & lastRow)
    Range("BP2:BP" & lastRow).Select
    Selection.copy
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.NumberFormat = "0"

    Columns("J:J").Select
    Selection.copy
    Columns("BQ:BQ").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:= _
        False, Transpose:=False
    Range("BR2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/2,"""")"
    Range("BR2").Select
    Selection.AutoFill Destination:=Range("BR2:BR" & lastRow)
    Range("BR2:BR" & lastRow).Select
    Selection.copy
    Range("J2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Columns("R:R").Select
    Selection.copy
    Columns("BS:BS").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:= _
        False, Transpose:=False
    Range("BT2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/2,"""")"
    Range("BT2").Select
    Selection.AutoFill Destination:=Range("BT2:BT" & lastRow)
    Range("BT2:BT" & lastRow).Select
    Selection.copy
    Range("R2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Formatting Process Parameters Spec

Sheets("ProcessParameters").Select
    
processColumn = Rows(1).Find(What:="User-Def. Text", LookIn:=xlValues).Column
y = Rows(1).Find(What:="Recipe Number", LookIn:=xlValues).Column
lastRow = Cells(Rows.Count, processColumn).End(xlUp).Row

If Cells(1, processColumn + 1).Value <> "Process Parameter Description" Then
emptyColumn = Rows(1).Find(What:="", LookIn:=xlValues).Column
Cells(1, processColumn + 1).Value = "Process Parameter Description"
Cells(1, processColumn + 2).Value = "Recipe-Stage-Operation-Action"
Cells(1, processColumn + 3).Value = "Recipe Version Used on floor"
Range(Cells(2, processColumn + 3), Cells(lastRow, processColumn + 3)).Select
Selection.FormulaR1C1 = "=IFERROR(VLOOKUP(RC" & y & ",WorkInstructions!C5,1,0),""Delete"")"
Cells(1, processColumn + 4).Value = "Process Step Description (LOOKUP)"

    For x = 2 To lastRow
        temp = Cells(x, processColumn).Value
        processPos = InStr(temp, "Process Parameter Description")
        ProcessPosCol = InStr(temp, ";")
        longPos = 0
        longPos = InStr(temp, "Process Parameter Description (French)")

        If processPos = 1 Then
        
            If longPos = 0 Then
                processStep = Split(temp, ":", 2)(1)
                Cells(x, processColumn + 1).Value = processStep
            End If
            
            If longPos = 1 Then
            
                If processPos = 0 Then
                    longText = Split(Mid(temp, longPos), "Process Parameter Description (French):")(1)
                    Cells(x, processColumn + 1).Value = longText
                Else
                    processStep = Split(Mid(temp, ProcessPosCol), ":")(1)
                    longText = Split(Mid(temp, longPos, ProcessPosCol - 1), ":", 2)(1)
                    If longText = "" Then
                    Cells(x, processColumn + 1).Value = processStep
                    Else
                    Cells(x, processColumn + 1).Value = processStep & " " & longText
                    End If
                End If
                
            End If
            
        End If
        
    Next

Cells.Select
Selection.AutoFilter
Selection.AutoFilter
y = processColumn + 3

letterCol = Chr(y + 64)
Set rr = Range("" & letterCol & "1:" & letterCol & "" & lastRow & "")

For Each Cell In rr
    If Cell = "Delete" Then
        ActiveSheet.Range("$A$1:$BQ" & lastRow & "").AutoFilter Field:=y, Criteria1:="Delete"
        Range("A2:BQ" & lastRow).SpecialCells(xlCellTypeVisible).Select
        Selection.EntireRow.Delete
        ActiveSheet.Range("$A$1:$BQ" & lastRow & "").AutoFilter Field:=y
        Range("$B$1:$BQ" & lastRow & "").Select
        Selection.NumberFormat = "General"
        Exit For
    End If
Next

    
End If

'Formatting Process Instructions Spec

Sheets("ProcessInstructions").Select

processColumn = Rows(1).Find(What:="User-Def. Text", LookIn:=xlValues).Column
y = Rows(1).Find(What:="Recipe Number", LookIn:=xlValues).Column
lastRow = Cells(Rows.Count, processColumn).End(xlUp).Row

If Cells(1, processColumn + 2).Value <> "Process Parameter Description" Then

emptyColumn = Rows(1).Find(What:="", LookIn:=xlValues).Column
Cells(1, processColumn + 2).Value = "Process Parameter Description"
Cells(1, processColumn + 1).Value = "Recipe-Stage-Operation-Action"
Cells(1, processColumn + 3).Value = "Recipe Version Used on floor"
Range(Cells(2, processColumn + 3), Cells(lastRow, processColumn + 3)).Select
Selection.FormulaR1C1 = "=IFERROR(VLOOKUP(RC" & y & ",WorkInstructions!C5,1,0),""Delete"")"

'Formatting WorkInstruction WorkSheet
For ii = 2 To lastRow
    ReDim LArray(0)
    LText = Cells(ii, processColumn).Text
    LArray() = Split(LText, ";")
    iiii = UBound(LArray, 1)
    
    For iii = 0 To iiii
        y = InStr(LArray(iii), "Process Step Descriptoin:")
        If y > 0 Then
            If InStr(LArray(0), "Long Text:") > 0 Then
                yy = 1
                Exit For
            End If
        End If
    Next
    
    If yy = 1 Then
        Cells(ii, processColumn + 2).Value = Split(LArray(iii), ":", 2)(1) + " " + Split(LArray(0), ":", 2)(1)
    Else
        On Error Resume Next
        Cells(ii, processColumn + 2).Value = Split(LArray(0), ":", 2)(1)
    End If
        
    y = 0
    yy = 0

    
Next

'    For x = 2 To lastRow
 '       temp = Cells(x, processColumn).Value
  '      processPos = InStr(temp, "Process Step Descriptoin:")
   '     longPos = InStr(temp, "Long Text")

    '    If processPos = 1 Then
     '       processStep = Split(temp, ":", 2)(1)
      '      Cells(x, processColumn + 2).Value = processStep
       ' Else
        '    If longPos = 1 Then
         '       If processPos = 0 Then
          '          longText = Split(Mid(temp, longPos), "Long Text:")(1)
           '         Cells(x, processColumn + 2).Value = longText
            '    Else
             '       processStep = Split(Mid(temp, processPos - 1), ":")(1)
              '      longText = Split(Mid(temp, longPos, processPos - longPos - 1), ":", 2)(1)
               '     Cells(x, processColumn + 2).Value = processStep & " " & longText
                'End If
            'End If
        'End If
   ' Next
    
Cells.Select
Selection.AutoFilter
Selection.AutoFilter
y = processColumn + 3

letterCol = Chr(y + 64)
Set rrr = Range("" & letterCol & "1:" & letterCol & "" & lastRow & "")

For Each Cell In rrr
    If Cell = "Delete" Then
        ActiveSheet.Range("$A$1:$BQ" & lastRow & "").AutoFilter Field:=y, Criteria1:="Delete"
        Range("A2:BQ" & lastRow).SpecialCells(xlCellTypeVisible).Select
        Selection.EntireRow.Delete
        ActiveSheet.Range("$A$1:$BQ" & lastRow & "").AutoFilter Field:=y
        Range("$B$1:$BQ" & lastRow & "").Select
        Selection.NumberFormat = "General"
        Exit For
    End If
Next

End If
    
'Formatting Process Instructions and Parameters to basic data entry file

Sheets("InstructionsAndParameters").Select

If Cells(1, 1).Value = "" Then
    Sheets("ProcessInstructions").Select
    ActiveSheet.UsedRange.copy
    Sheets("InstructionsAndParameters").Select
    ActiveSheet.Select
    Cells(1, 1).Select
    ActiveSheet.Paste
    Columns("J:P").EntireColumn.Insert
    lastRow = Sheets("InstructionsAndParameters").Cells(Rows.Count, 1).End(xlUp).Row
    Range("Q1:Q" & lastRow).Value = Range("V1:V" & lastRow).Value
    Columns("R:W").Delete
    Range("J1:O1").Value = Sheets("ProcessParameters").Range("K1:P1").Value
    Range("P1").Value = "Process Parameter Description"

    parOperationCol = Sheets("ProcessParameters").Rows(1).Find(What:="Operation", LookIn:=xlValues).Column
    parActionCol = Sheets("ProcessParameters").Rows(1).Find(What:="Action", LookIn:=xlValues).Column
    parSequenceCol = Sheets("ProcessParameters").Rows(1).Find(What:="Seq.", LookIn:=xlValues).Column
    
    instOperationCol = Rows(1).Find(What:="Operation", LookIn:=xlValues).Column
    instActionCol = Rows(1).Find(What:="Action", LookIn:=xlValues).Column
    instSequenceCol = Rows(1).Find(What:="Process Parameter", LookIn:=xlValues).Column 'This is use to check if there is a value in the process paramter column
    
    parameterCol = Sheets("ProcessParameters").Rows(1).Find(What:="Process Parameter", LookIn:=xlValues).Column
    descriptionCol = Sheets("ProcessParameters").Rows(1).Find(What:="Process Parameter Description", LookIn:=xlValues).Column
    
    specColumn = Sheets("ProcessParameters").Rows(1).Find(What:="Spec", LookIn:=xlValues).Column
    lastRow = Sheets("ProcessParameters").Cells(Rows.Count, specColumn).End(xlUp).Row
    
    For x = 2 To lastRow
        spec = Sheets("ProcessParameters").Cells(x, specColumn).Value
        Set found = Columns(1).Find(What:=spec, LookIn:=xlValues)
        
        If found Is Nothing Then
            
        Else
            instructionRow = Columns(1).Find(What:=spec, LookIn:=xlValues).Row
            parOperation = Sheets("ProcessParameters").Cells(x, parOperationCol).Value
            parAction = Sheets("ProcessParameters").Cells(x, parActionCol).Value
            parSequence = Sheets("ProcessParameters").Cells(x, parSequenceCol).Value
            
            Do While spec = Cells(instructionRow, 1)
                
                instOperation = Cells(instructionRow, instOperationCol).Value
                instAction = Cells(instructionRow, instActionCol).Value
                instSequence = Cells(instructionRow, instSequenceCol).Value
                
                If instOperation = parOperation Then
                    If instAction = parAction Then
                        If instSequence = "" Then
                            instructionRow = instructionRow + 1
                            Rows(instructionRow).Insert
                            Range("A" & instructionRow & ":I" & instructionRow).Value = Range("A" & instructionRow - 1 & ":I" & instructionRow - 1).Value
                            Cells(instructionRow, "Q").Value = Cells(instructionRow - 1, "Q").Value
                            parameterDescription = Sheets("ProcessParameters").Cells(x, descriptionCol).Value
                            Range("J" & instructionRow & ":O" & instructionRow).Value = Sheets("ProcessParameters").Range("K" & x & ":P" & x).Value
                            Cells(instructionRow, "P").Value = parameterDescription
                            Rows(instructionRow - 1).EntireRow.Delete
                        Else
                            Do While instAction = parAction
                                instructionRow = instructionRow + 1
                                instAction = Cells(instructionRow, instActionCol).Value
                            Loop
                            Rows(instructionRow).Insert
                            Range("A" & instructionRow & ":I" & instructionRow).Value = Range("A" & instructionRow - 1 & ":I" & instructionRow - 1).Value
                            Cells(instructionRow, "Q").Value = Cells(instructionRow - 1, "Q").Value
                            parameterDescription = Sheets("ProcessParameters").Cells(x, descriptionCol).Value
                            Range("J" & instructionRow & ":O" & instructionRow).Value = Sheets("ProcessParameters").Range("K" & x & ":P" & x).Value
                            Cells(instructionRow, "P").Value = parameterDescription
                        End If
                        
                    End If
                End If
                
                instructionRow = instructionRow + 1
                
            Loop
        End If
    Next

    Columns(1).NumberFormat = "0"
    
End If

'Copying data from process Instruction and parameters to Data entry form

Sheets("DataEntryForm").Select
lastRow = Sheets("DataEntryForm").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:A" & lastRow + 1).EntireRow.Delete
Range("A2").Select
lastRow = Sheets("InstructionsAndParameters").Cells(Rows.Count, 1).End(xlUp).Row
Range("A1:A" & lastRow).Value = Range("A1:A" & lastRow).Value

Sheets("DataEntryForm").Range("A2:A" & lastRow).Value = Sheets("InstructionsAndParameters").Range("A2:A" & lastRow).Value
Sheets("DataEntryForm").Range("N2:N" & lastRow).Value = Sheets("InstructionsAndParameters").Range("B2:B" & lastRow).Value
Sheets("DataEntryForm").Range("O2:O" & lastRow).Value = Sheets("InstructionsAndParameters").Range("E2:E" & lastRow).Value
Sheets("DataEntryForm").Range("P2:P" & lastRow).Value = Sheets("InstructionsAndParameters").Range("F2:F" & lastRow).Value
Sheets("DataEntryForm").Range("R2:R" & lastRow).Value = Sheets("InstructionsAndParameters").Range("G2:G" & lastRow).Value
Sheets("DataEntryForm").Range("S2:S" & lastRow).Value = Sheets("InstructionsAndParameters").Range("H2:H" & lastRow).Value
Sheets("DataEntryForm").Range("T2:T" & lastRow).Value = Sheets("InstructionsAndParameters").Range("I2:I" & lastRow).Value
Sheets("DataEntryForm").Range("V2:V" & lastRow).Value = Sheets("InstructionsAndParameters").Range("Q2:Q" & lastRow).Value
Sheets("DataEntryForm").Range("W2:W" & lastRow).Value = Sheets("InstructionsAndParameters").Range("J2:J" & lastRow).Value
Sheets("DataEntryForm").Range("X2:X" & lastRow).Value = Sheets("InstructionsAndParameters").Range("P2:P" & lastRow).Value
Sheets("DataEntryForm").Range("Z2:Z" & lastRow).Value = Sheets("InstructionsAndParameters").Range("K2:K" & lastRow).Value
Sheets("DataEntryForm").Range("AA2:AA" & lastRow).Value = Sheets("InstructionsAndParameters").Range("L2:L" & lastRow).Value
Sheets("DataEntryForm").Range("AB2:AB" & lastRow).Value = Sheets("InstructionsAndParameters").Range("M2:M" & lastRow).Value
Sheets("DataEntryForm").Range("AC2:AC" & lastRow).Value = Sheets("InstructionsAndParameters").Range("N2:N" & lastRow).Value

'Determine which formula best fit the respected column of SAP mat, Proc Stages and Sku Number... Use file MDSS and respective plant for reference

Sheets("DataEntryForm").Select

planttxt = Sheets("Notes").Cells(6, 10).Value
lastRow = Sheets("DataEntryForm").Cells(Rows.Count, 1).End(xlUp).Row

    Cells(2, 2).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1]," & planttxt & "!C[-1]:C[4],2,FALSE),""No Spec Ref"")"
    Cells(2, 4).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],SpecStatus!C[-3]:C[2],6,FALSE),""Please Update Data File/SAP"")"
    Cells(2, 5).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-4],RecipeQuantities!C[-4]:C[13],18,FALSE),""Please Check formula/Please Enter Spec #"")"
    Cells(2, 6).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1],MDSS!C[-5]:C,6,FALSE),""No SAP #"")"
    Cells(2, 7).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-2],MDSS!C[-6]:C[-4],3,FALSE),RC[7])"
    Cells(2, 8).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],MDSS!C[-7]:C[-6],2,FALSE),""No SKU#"")"
    Cells(2, 9).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-4],MDSS!C[-8]:C[-4],4,FALSE),""No SAP #/No Product Cat."")"
    Cells(2, 10).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-5],MDSS!C[-9]:C[-5],5,FALSE),""No SAP #/No Product Sub Cat."")"
    Cells(2, 11).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-10]," & planttxt & "!C[-10]:C[-5],5,FALSE),""No Spec Ref"")"
    Cells(2, 12).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-11]," & planttxt & "!C[-11]:C[-6],6,FALSE),""No Spec Ref."")"

Range("B2:L2").Select
Selection.AutoFill Destination:=Range("B2:L" & lastRow)
Cells(2, 2).Select

'Changing E formula for blanks

    'Cells.Select
    'Selection.AutoFilter
    'ActiveSheet.Range("$A$1:$AD" & lastRow).AutoFilter Field:=5, Criteria1:="="
    'Range("E2:E" & lastRow).Select
    'Selection.ClearContents
    'Cells.Select
    'Selection.AutoFilter
    'Columns("E:E").Select
    'Selection.SpecialCells(xlCellTypeBlanks).Select
    'Selection.FormulaR1C1 = "=VLOOKUP(RC[-4]," & planttxt & "!C1:C3,3,FALSE)"

'Sorting from FG to WIP
    Cells.Select
    ActiveWorkbook.Worksheets("DataEntryForm").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DataEntryForm").Sort.SortFields.Add Key:=Range( _
        "A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("DataEntryForm").Sort
        .SetRange Range("A1:AI" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Formatting Cells With line
For x = 2 To lastRow
    If Cells(x, 1).Value <> Cells(x - 1, 1).Value Then
        Cells(x, 1).EntireRow.Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    End If
Next

    
Application.ScreenUpdating = True
Sheets("Notes").Select
MsgBox ("Your File has finished running")



End Sub



