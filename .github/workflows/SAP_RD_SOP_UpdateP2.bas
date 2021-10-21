'Update SOP file with Bill Of Material formulation to analysis Standard Operating Procedure further
Sub UPDATE2()

Dim xWIColNum, xWIColNumUserTxt, emptyColumn As Integer
Dim lastRow, x As Double
Dim letterCol, letterColTxt As String
Dim r As Excel.Range
Dim checkFill As Boolean
Dim LArray() As String
Dim LText As String
Dim ii, iii, iiii As Integer
Dim y As Integer
'Instructions and Parameters
Dim sequenceNum As Integer, specCol As Integer, spec As String
Dim instOperationCol As Integer, parOperationCol As Integer, instOperation As String, parOperation As String
Dim instActionCol As Integer, parActionCol As Integer, instAction As String, parAction As String
Dim parameterCol As Integer, descriptionCol As Integer, parSequenceCol As Integer, instSequenceCol As Integer
Dim parameterDescription As String, instSequence As String, parSequence As String
Dim instructionRow As Double
Dim found
'Input data into data entry Form
Dim matchnumber As Integer
Dim planttxt As String
'Updating Formulation Legend
Dim formulaRecipeSpecRow, xx As Double
Dim formulaRecipeSpecCount, recipetextCol As Integer
Dim recipecol, recipecol2 As String

Application.ScreenUpdating = False

If Cells(19, 2).Value = 0 Then
'Formatting Work Instructions
Sheets("WorkInstructions").Select

If Cells(1, 8).Value <> "Recipe Item" Then
Columns("H:K").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, 8).Value = "Recipe Item"
Cells(1, 9).Value = "Action No."
Cells(1, 10).Value = "Operation No."
Cells(1, 11).Value = "FormulaItem"
End If

xWIColNum = Rows(1).Find(What:="Work Instruction Step", LookIn:=xlValues).Column
xWIColNumUserTxt = Rows(1).Find(What:="User-Def. Text", LookIn:=xlValues).Column

letterCol = Chr(xWIColNum + 64)
letterColTxt = Chr(xWIColNumUserTxt + 64)
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
checkFill = False


Set r = Range("" & letterCol & "1:" & letterCol & "" & lastRow)
    For Each Cell In r
        If Cell = "" Then
            checkFill = True
            Exit For
        End If
    Next
    
If checkFill = True Then
    Cells.Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$BQ" & lastRow).AutoFilter Field:=xWIColNum, Criteria1:="="
    Columns("" & letterCol & "").Select
    Selection.ClearContents
    Cells(1, xWIColNum).FormulaR1C1 = "Work Instruction Step"
    Selection.AutoFilter
    Selection.NumberFormat = "General"
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("" & letterCol & "").Select
    Selection.copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    For ii = 2 To lastRow
        On Error Resume Next
        Cells(ii, 7).Value = Split(Cells(ii, 6).Value, ":", 2)(1)
    Next
    Cells.Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$BQ" & lastRow).AutoFilter Field:=7, Criteria1:="="
    Columns("G").Select
    Selection.ClearContents
    Cells(1, 7).FormulaR1C1 = "Work Instruction Step"
    Selection.AutoFilter
    Selection.NumberFormat = "General"
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("G").Select
    Selection.copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
End If

'Formatting WorkInstruction WorkSheet
For ii = 2 To lastRow
    ReDim LArray(0)
    LText = Cells(ii, xWIColNumUserTxt).Text
    LArray() = Split(LText, ";")
    iiii = UBound(LArray, 1)
    
    For iii = 0 To iiii
        y = InStr(LArray(iii), "Formula Item:")
        If y > 0 Then
            Cells(ii, 8).Value = Split(LArray(iii), ":", 2)(1)
            If Cells(ii, 8).Value <> "" Then
                Cells(ii, 9).Value = Cells(ii, 6).Value
                Cells(ii, 10).Value = Cells(ii, 7).Value
                Cells(ii, 11).Value = "Formula Item"
            End If
        End If
        y = 0
    Next
Next

'Deleting Blank Row
Columns("H").Select
Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Select
Selection.Delete

'Adding data to Instruction and Parameters file
    Sheets("InstructionsAndParameters").Select
    
    parOperationCol = Sheets("WorkInstructions").Rows(1).Find(What:="Operation", LookIn:=xlValues).Column
    parActionCol = Sheets("WorkInstructions").Rows(1).Find(What:="Action", LookIn:=xlValues).Column
    parSequenceCol = Sheets("WorkInstructions").Rows(1).Find(What:="Seq.", LookIn:=xlValues).Column
    
    instOperationCol = Rows(1).Find(What:="Operation", LookIn:=xlValues).Column
    instActionCol = Rows(1).Find(What:="Action", LookIn:=xlValues).Column
    instSequenceCol = Rows(1).Find(What:="Process Parameter", LookIn:=xlValues).Column 'This is use to check if there is a value in the process paramter column
    
    parameterCol = Sheets("WorkInstructions").Rows(1).Find(What:="Recipe item", LookIn:=xlValues).Column
    descriptionCol = Sheets("WorkInstructions").Rows(1).Find(What:="Recipe item", LookIn:=xlValues).Column
    
    specColumn = Sheets("WorkInstructions").Rows(1).Find(What:="Spec", LookIn:=xlValues).Column
    lastRow = Sheets("WorkInstructions").Cells(Rows.Count, specColumn).End(xlUp).Row
    
    For x = 2 To lastRow
        spec = Sheets("WorkInstructions").Cells(x, specColumn).Value
        Set found = Columns(1).Find(What:=spec, LookIn:=xlValues)
        
        If found Is Nothing Then
            
        Else
            instructionRow = Columns(1).Find(What:=spec, LookIn:=xlValues).Row
            parOperation = Sheets("WorkInstructions").Cells(x, parOperationCol).Value
            parAction = Sheets("WorkInstructions").Cells(x, parActionCol).Value
            parSequence = Sheets("WorkInstructions").Cells(x, parSequenceCol).Value
            
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
                            parameterDescription = Sheets("WorkInstructions").Cells(x, descriptionCol).Value
                            Range("J" & instructionRow & ":O" & instructionRow).Value = Sheets("WorkInstructions").Range("K" & x & ":P" & x).Value
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
                            parameterDescription = Sheets("WorkInstructions").Cells(x, descriptionCol).Value
                            Range("J" & instructionRow & ":O" & instructionRow).Value = Sheets("WorkInstructions").Range("K" & x & ":P" & x).Value
                            Cells(instructionRow, "P").Value = parameterDescription
                        End If
                        
                    End If
                End If
                
                instructionRow = instructionRow + 1
                
            Loop
        End If
    Next
    Columns(1).NumberFormat = "0"


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

'Updating Formulation Legend & Format Recipe Quantities
Sheets("RecipeQuantities").Select

recipetextCol = Rows(1).Find(What:="User-Def. Text", LookIn:=xlValues).Column
recipecol = Chr(recipetextCol + 64)
lastRow = Sheets("RecipeQuantities").Cells(Rows.Count, 1).End(xlUp).Row

For ii = 2 To lastRow
    ReDim LArray(0)
    LText = Cells(ii, recipetextCol).Text
    LArray() = Split(LText, ";")
    iiii = UBound(LArray, 1)
    
    For iii = 0 To iiii
        y = InStr(LArray(iii), "Item Description:")
        If y > 0 Then
            Cells(ii, 73).Value = Split(LArray(iii), ":", 2)(1)
        End If
        y = 0
    Next
Next


'Updating DataEntry Form

x = 2

Do While x < lastRow

    formulaRecipeSpecRow = Application.WorksheetFunction.Match(Sheets("RecipeQuantities").Cells(x, 1).Value, Sheets("DataEntryForm").Range("A:A"), 0)
    formulaRecipeSpecCount = Application.WorksheetFunction.CountIf(Sheets("RecipeQuantities").Range("A:A"), Sheets("RecipeQuantities").Cells(x, 1))
    
    For xx = 1 To formulaRecipeSpecCount - 1
        Sheets("DataEntryForm").Cells(formulaRecipeSpecRow + xx - 1, 32).Value = Sheets("RecipeQuantities").Cells(xx + x, 73).Value 'Change Column
        Sheets("DataEntryForm").Cells(formulaRecipeSpecRow + xx - 1, 33).Value = Sheets("RecipeQuantities").Cells(xx + x, 18).Value
        Sheets("DataEntryForm").Cells(formulaRecipeSpecRow + xx - 1, 34).Value = Sheets("RecipeQuantities").Cells(xx + x, 11).Value
        Sheets("DataEntryForm").Cells(formulaRecipeSpecRow + xx - 1, 35).Value = Sheets("RecipeQuantities").Cells(xx + x, 15).Value
    Next
    
    x = x + formulaRecipeSpecCount
    
Loop

End If

Application.ScreenUpdating = True
Sheets("Notes").Select
Cells(19, 2).Value = 1

MsgBox ("Your file has finished running.")

End Sub
