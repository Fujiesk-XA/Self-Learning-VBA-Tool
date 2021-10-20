'This program is substitue for BOM by level for all plants rather than single Ladder for BOM
'Any recipe greater than 400000000000 is formulation recipe an consider to have a subset
'Repeated Recipe cause this to loop, will continue till the end of excel formula, 
'The format is based on the comment and the colouring of the cell.
'The next iteration is the begining of the comment and the colouring represent which preceding specification is coming from

Sub MaterialNumberExplosionPRog()

Dim lastRow, lastRow2, lastRow3, i, ii, lastRow5 As Double
Dim convertTonumber As Double
Dim x, processColumn, y As Integer
Dim columnletter, columnletter2 As String
Dim executeprogram As Boolean
Dim r As Range
Dim CheckFill As Boolean

Application.ScreenUpdating = False

Sheets("Data Table").Select

executeprogram = True
'Format Entered Data
lastRow2 = Cells(Rows.Count, 1).End(xlUp).Row

x = 1

    If lastRow2 = 4 Then
        lastRow2 = Cells(Rows.Count, 2).End(xlUp).Row
        x = 2
        
        If lastRow2 < 4 Then
            executeprogram = False
            msgBox ("PLEASE ENTER MAT# or SPEC #")
        End If
        
    End If
    
If lastRow2 > 4 Then
    For i = 5 To lastRow2
    
        convertTonumber = Cells(i, x).Value
    
        convertTonumber = convertTonumber * 2
    
        Cells(i, x).Value = convertTonumber / 2

    Next
End If

'Clear Contents, Fill and Comments
    Range("D8").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Delete Shift:=xlToLeft
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.ClearComments
    Selection.NumberFormat = "0"

If executeprogram = True Then

'Formatting Recipe QUANTITIES
Sheets("RecipeQuantities").Select
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
processColumn = Rows(1).Find(what:="User-Def. Text", LookIn:=xlValues).Column
Cells(1, processColumn + 2).FormulaR1C1 = "Division"
columnletter = Chr(processColumn + 1 + 64)
Columns("" & columnletter & ":" & columnletter & "").ClearContents

'Columns A Formatting

CheckFill = False
Set r = Sheets("RecipeQuantities").Range("A1:A" & lastRow)

For Each cell In r
    If cell = "" Then
        CheckFill = True
        Exit For
    End If
Next
        'Columns A Formatting
        If CheckFill = True Then
        lastRow3 = Cells(Rows.Count, 6).End(xlUp).Row
        Cells(1, 1).Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$S" & lastRow3).AutoFilter Field:=1, Criteria1:="="
        Range("A3:C3").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
        ActiveSheet.Range("$A$1:$S" & lastRow3).AutoFilter Field:=1
        Columns("A:A").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.NumberFormat = "General"
        Columns("A:C").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=R[-1]C"
        Columns("A:C").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("A1").Select
        End If
    
'Columns A,J &K Spec&Mat Number Formatting

For i = 1 To 4
    If i = 1 Then
        y = 1
    ElseIf i = 2 Then
        y = 10
    ElseIf i = 3 Then
        y = 11
    ElseIf i = 4 Then
        y = 18
    End If
    
    columnletter2 = Chr(64 + y)
    columnletter = Chr(processColumn + 1 + 64)
    Columns("" & columnletter & ":" & columnletter & "").ClearContents
    Columns("" & columnletter2 & ":" & columnletter2 & "").Select
    Selection.Copy
    columnletter = Chr(processColumn + 1 + 64)
    Columns("" & columnletter & ":" & columnletter & "").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:= _
        False, Transpose:=False
        
    Cells(1, processColumn + 2).FormulaR1C1 = "Division"
    Cells(1, processColumn + 2).Value = "Division"
    Cells(2, processColumn + 2).NumberFormat = "General"
    Cells(2, processColumn + 2).FormulaR1C1 = "=RC[-1]/2"
    Cells(2, processColumn + 2).Select
    columnletter = Chr(processColumn + 2 + 64)
    Selection.AutoFill Destination:=Range("" & columnletter & "2:" & columnletter & "" & lastRow)
    Range("" & columnletter & "2:" & columnletter & "" & lastRow).Select
    Selection.Copy
    Range("" & columnletter2 & "2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("" & columnletter2 & ":" & columnletter2 & "").Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "0"
Next
    
Cells.Select
ActiveWorkbook.Worksheets("RecipeQuantities").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("RecipeQuantities").Sort.SortFields.Add Key:=Range( _
    "A1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
    xlSortTextAsNumbers
With ActiveWorkbook.Worksheets("RecipeQuantities").Sort
    .SetRange Range("A2:BQ" & lastRow)
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
    
'Program Starts
Sheets("Data Table").Select
If x = 1 Then
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "=INDEX(RecipeQuantities!C[9]:C[16],MATCH(RC[-1],RecipeQuantities!C[16],0),1)"
    Selection.NumberFormat = "0"
    Selection.AutoFill Destination:=Range("B5:B" & lastRow2)
End If

Range("B5:B" & lastRow2).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("D8").Select
ActiveSheet.Paste


'Program Find

Dim specifcount, xx, xy, xyz, iii, iiii, iiiii, iiiiii, countcheck2, countcheck3, countcheck4, countcheck5, variable3count, variable4count, variable5count As Integer
Dim specRow, checkvalue, timesexp As Double
Dim nextLineCounter As Double
timesexp = Sheets("Data Table").Cells(1, 6).Value

For i = 8 To lastRow2 + 3

    xx = 5
    countcheck2 = 1
    z = 1
    xyz = 5

    For iiii = 1 To timesexp
    
        If xx <= 204 Then
            Cells(i, xx).AddComment
            Cells(i, xx).Comment.Visible = False
            Cells(i, xx).Comment.Text Text:="USER:" & Chr(10) & "Explosion " & iiii
        End If
        
        For iii = 1 To countcheck2
        
            If iii = 1 Then
                xyz = xx 'Reset each explostion counter
            End If
            
            specifcount = Application.WorksheetFunction.countif(Sheets("RecipeQuantities").Range("A:A"), Cells(i, xyz - countcheck2 + iii - 1).Value) 'Determine the Spec loop
            specRow = Application.WorksheetFunction.Match(Cells(i, xyz - countcheck2 - 1 + iii).Value, Sheets("RecipeQuantities").Range("A:A"), 0)    'Locate the Row of the Spec in Recipe quantities
            
            For ii = 1 To specifcount - 1
                
                If Sheets("RecipeQuantities").Cells(specRow + ii, 9).Value = "PRIMARY OUTPUT" Then
                    GoTo Nextiteration
                ElseIf Sheets("RecipeQuantities").Cells(specRow + ii, 9).Value = "SECONDARY OUTPUT" Then
                    GoTo Nextiteration
                End If
                
                On Error Resume Next
                If Sheets("RecipeQuantities").Cells(specRow + ii, 11).Value >= 400000000000# Then
                        Cells(i, xx).Value = Sheets("RecipeQuantities").Cells(specRow + ii, 11).Value
                        Cells(i, xx).Interior.ColorIndex = 32 + iii
                            If xx > 204 Then
                                Cells(i, xx).Interior.ColorIndex = 4
                            End If
                        xx = xx + 1
                        xy = xy + 1
                End If
Nextiteration:
            Next
            
        Next
        
        countcheck2 = xy 'Count total 4*spec in last run
        xy = 0
        
    Next

Next
    
    
End If
    
Sheets("Data Table").Select

Application.ScreenUpdating = True

End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////
            'Because of the limitation of the extract it can only be one layer at a time therefore does not show the complete set, by repeating copy the next level of the chain the extract can be run multiple time to find all extract
            'The next part is to determine which extract version needs to be extracted using vlookup, while the first layer is always the next recipe for the next layer so vlookup on the next sheet will identify the next set to extract
           Sub TableToColumn()
Dim lastcolumn, lastRow2, i, ii, x, xx, lastRow3 As Long

Sheets("Validation").Select
Cells.Select
Selection.ClearContents

lastRow2 = Sheets("Data Table").Cells(Rows.Count, 4).End(xlUp).Row
xx = 2
For i = 8 To lastRow2
    lastcolumn = Sheets("Data Table").Cells(i, Columns.Count).End(xlToLeft).Column
    For x = 5 To lastcolumn
        Sheets("Validation").Cells(xx, 1).Value = Sheets("Data Table").Cells(i, x).Value
        Sheets("Validation").Cells(xx, 2).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],RecipeQuantities!C[-1],1,FALSE),""ERROR"")"
        xx = xx + 1
    Next

Next

Sheets("Validation").Select
Cells(1, 1).Value = "DATA"
Cells(1, 2).Value = "VLOOKUP"
Cells(1, 4).Value = "ERROR"
lastRow3 = Sheets("Validation").Cells(Rows.Count, 1).End(xlUp).Row
ActiveSheet.Range("$A$1:$B$" & lastRow3).RemoveDuplicates Columns:=Array(1, 2), _
        Header:=xlYes
Range("A:B").Select
Selection.NumberFormat = "0"
Cells(1, 1).Select

lastRow3 = Sheets("Validation").Cells(Rows.Count, 1).End(xlUp).Row
ii = 2
For i = 1 To lastRow3
    If Cells(i, 2).Value = "ERROR" Then
        Cells(ii, 4).Value = Cells(i, 1).Value
        ii = ii + 1
    End If
Next

    Columns("D:D").Select
    ActiveWorkbook.Worksheets("Validation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Validation").Sort.SortFields.Add Key:=Range("D1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Validation").Sort
        .SetRange Range("D2:D" & lastRow3)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
Sheets("Data Table").Select
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    'Highligh error
    
Sub FindMissingSpec()
Dim fnd As String, FirstFound As String
Dim FoundCell As Range, rng As Range
Dim myRange As Range, LastCell As Range
Dim i, lastRow4 As Long
lastRow4 = Sheets("Validation").Cells(Rows.Count, 4).End(xlUp).Row
'What value do you want to find (must be in string form)?

For i = 2 To lastRow4
Sheets("Validation").Select
  fnd = Cells(i, 4).Text
Sheets("Data Table").Select

Set myRange = ActiveSheet.UsedRange
Set LastCell = myRange.Cells(myRange.Cells.Count)
Set FoundCell = myRange.Find(what:=fnd, after:=LastCell)

'Test to see if anything was found
  If Not FoundCell Is Nothing Then
    FirstFound = FoundCell.Address
  Else
    GoTo NothingFound
  End If

Set rng = FoundCell

'Loop until cycled through all unique finds
  Do Until FoundCell Is Nothing
    'Find next cell with fnd value
      Set FoundCell = myRange.FindNext(after:=FoundCell)
    
    'Add found cell to rng range variable
      Set rng = Union(rng, FoundCell)
    
    'Test to see if cycled through to first found cell
      If FoundCell.Address = FirstFound Then Exit Do
      
  Loop

'Select Cells Containing Find Value
  rng.Select
  Selection.Interior.ColorIndex = 3
  
Next
Sheets("Data Table").Select
Exit Sub

'Error Handler
NothingFound:
  msgBox "No values were found in this worksheet"

End Sub


         
            
            
            
            
            
            
            
