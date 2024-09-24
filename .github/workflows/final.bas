'Section to extrapolate data based on 5 different line
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim introw As Integer
    Dim intBlanks As Integer
    Dim blnFound As Boolean
    Dim lngRef As Long
    Dim colSelected As Integer
    Dim COL_S_RECIPE, COL_S_DSC, COL_S_INTERVAL, COL_S_MIXSIZE As Integer
    Dim lineSelected As String
    introw = 3
    intBlanks = 0
    blnFound = False

    Select Case Target.Column
        Case 3, 16, 29, 42, 55:
            If (IsNumeric(Cells(Target.Row, Target.Column).Value) And _
                (Cells(Target.Row, Target.Column).Value <> "")) Then
                'populate from database and redirect to start time
                lngRef = CLng(Cells(Target.Row, Target.Column).Value)
                colSelected = Target.Column
                Select Case colSelected
                    Case 3:
                        COL_S_RECIPE = 4
                        COL_S_DSC = 5
                        COL_S_INTERVAL = 6
                        COL_S_MIXSIZE = 8
                        lineSelected = "SL01"
                    Case 16:
                        COL_S_RECIPE = 17
                        COL_S_DSC = 18
                        COL_S_INTERVAL = 19
                        COL_S_MIXSIZE = 21
                        lineSelected = "SL02"
                    Case 29:
                        COL_S_RECIPE = 30
                        COL_S_DSC = 31
                        COL_S_INTERVAL = 32
                        COL_S_MIXSIZE = 34
                        lineSelected = "SL03"
                    Case 42:
                        COL_S_RECIPE = 43
                        COL_S_DSC = 44
                        COL_S_INTERVAL = 45
                        COL_S_MIXSIZE = 47
                        lineSelected = "SL04"
                    Case 55:
                        COL_S_RECIPE = 56
                        COL_S_DSC = 57
                        COL_S_INTERVAL = 58
                        COL_S_MIXSIZE = 60
                        lineSelected = "SL05"
                End Select
                Do While intBlanks < 5
                    If (Worksheets("SKU_LIST_ALL_LINE").Cells(introw, 1).Value = "") Then
                        intBlanks = intBlanks + 1
                    ElseIf (CLng(Worksheets("SKU_LIST_ALL_LINE").Cells(introw, 1).Value) = lngRef) Then
                        If (Worksheets("SKU_LIST_ALL_LINE").Cells(introw, 4).Value = lineSelected) Then
                            With Worksheets("scheduler")
                                .Cells(Target.Row, COL_S_RECIPE) = Worksheets("SKU_LIST_ALL_LINE").Cells(introw, 3)
                                .Cells(Target.Row, COL_S_DSC) = Worksheets("SKU_LIST_ALL_LINE").Cells(introw, 2)
                                .Cells(Target.Row, COL_S_INTERVAL) = Worksheets("SKU_LIST_ALL_LINE").Cells(introw, 8)
                                .Cells(Target.Row, COL_S_MIXSIZE) = Worksheets("SKU_LIST_ALL_LINE").Cells(introw, 10)
                                .Cells(Target.Row, COL_S_INTERVAL + 1).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],DOUGH_DATA!C1:C4,4,FALSE),"""")"
                                .Cells(Target.Row, COL_S_MIXSIZE + 5) = Worksheets("SKU_LIST_ALL_LINE").Cells(introw, 7)
                                If Target.Row = 3 Then
                                    .Cells(Target.Row, COL_S_MIXSIZE + 3).FormulaR1C1 = "=R[-2]C[-4]+R[-2]C[-2]+RC[-4]/60/24"
                                    .Cells(Target.Row, COL_S_MIXSIZE + 4).FormulaR1C1 = "=+RC[-1]+(RC[-3]*RC[-6]/60/24)" '1*RC[-6]/60/24"
                                Else
                                    .Cells(Target.Row, COL_S_MIXSIZE + 3).FormulaR1C1 = "=+R[-1]C[1]+R[-1]C[-1]/60/24"
                                    .Cells(Target.Row, COL_S_MIXSIZE + 4).FormulaR1C1 = "=+RC[-1]+(RC[-3]*RC[-6]/60/24)" '+1*RC[-6]/60/24"
                                End If
                            End With
                            intBlanks = 0
                            blnFound = True
                        End If
                    Else
                        intBlanks = 0
                    End If
                    introw = introw + 1
                Loop

                If (blnFound = False) Then
                    Worksheets("scheduler").Cells(Target.Row, COL_S_DSC) = "NOT FOUND"
                End If
                'Range("" & GOTO2_COL & Target.Row & "").Select
            End If
        Case COL_S_MIXSIZE
            Range("" & GOTO2_COL & Target.Row & "").Select
     End Select
     
End Sub
'Copy to sheets
Sub reformatSch()
'Template to be copied over
Sheets("scheduler").Activate
Sheets("scheduler").Range("A3:BL999").ClearContents


    Sheets("ProductionPlannerScheduler").Activate
    Cells.Select
    Range("B1").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Dim LnOneColSt, LnTwoColSt, LnThreeColSt, LnFourColSt, LnFiveColSt As Integer
    Dim LnOneRng, LnTwoRng, LnThreeRng, LnFourRng, LnFiveRng As Range
    Dim LnOneSKUCol, LnTwoSKUCol, LnThreeSKUCol, LnFourSKUCol, LnFiveSKUCol As Integer
    Dim LnOneBatCol, LnTwoBatCol, LnThreeBatCol, LnFourBatCol, LnFiveBatCol As Integer
    Dim LnOneFGDCol, LnTwoFGDCol, LnThreeFGDCol, LnFourFGDCol, LnFiveFGDCol As Integer
    Dim LnOneHrsCol, LnTwoHrsCol, LnThreeHrsCol, LnFourHrsCol, LnFiveHrsCol As Integer
    Dim skuCol, batCol, FGDCol, HrsCol As Integer
    Dim outputCol As Integer
    
    Dim w1Row, w3Row, w5Row, w7Row, w9Row, w11Row, w13Row As Integer
    Dim i, ii, outputPrintLine As Integer
    Dim lineName As String
    
    
    Set found1 = Cells.Find(what:="SL01", LookIn:=xlValues)
    If found1 Is Nothing Then
        MsgBox ("There is no SL01 present in the copied Schedule Template")
        Exit Sub
    Else
        LnOneColSt = found1.Column
    End If
    Set found2 = Cells.Find(what:="SL02", LookIn:=xlValues)
    If found2 Is Nothing Then
        MsgBox ("There is no SL02 present in the copied Schedule Template")
        Exit Sub
    Else
        LnTwoColSt = found2.Column
    End If
    Set found3 = Cells.Find(what:="SL03", LookIn:=xlValues)
    If found3 Is Nothing Then
        MsgBox ("There is no SL03 present in the copied Schedule Template")
        Exit Sub
    Else
        LnThreeColSt = found3.Column
    End If
    Set found4 = Cells.Find(what:="SL04", LookIn:=xlValues)
    If found4 Is Nothing Then
        MsgBox ("There is no SL04 present in the copied Schedule Template")
        Exit Sub
    Else
        LnFourColSt = found4.Column
    End If
    Set found5 = Cells.Find(what:="SL05", LookIn:=xlValues)
    If found5 Is Nothing Then
        MsgBox ("There is no SL05 present in the copied Schedule Template")
        Exit Sub
    Else
        LnFiveColSt = found5.Column
    End If

    Set LnOneRng = Range(Columns(LnOneColSt), Columns(LnTwoColSt - 1))
    Set LnTwoRng = Range(Columns(LnTwoColSt), Columns(LnThreeColSt - 1))
    Set LnThreeRng = Range(Columns(LnThreeColSt), Columns(LnFourColSt - 1))
    Set LnFourRng = Range(Columns(LnFourColSt), Columns(LnFiveColSt - 1))
    Set LnFiveRng = Range(Columns(LnFiveColSt), Columns(LnFiveColSt + 10))
'Line One, Finding SKU,HRS, DOUGH and Batches Column
    Set foundDelta = LnOneRng.Find(what:="SKU", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no SKU column for Line One")
        Exit Sub
    Else
        LnOneSKUCol = foundDelta.Column
    End If
    Set foundDelta = LnOneRng.Find(what:="Batches", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no Batches column for Line One")
        Exit Sub
    Else
        LnOneBatCol = foundDelta.Column
    End If
    Set foundDelta = LnOneRng.Find(what:="FG Dough", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no FG Dough column for Line One")
        Exit Sub
    Else
        LnOneFGDCol = foundDelta.Column
    End If
    Set foundDelta = LnOneRng.Find(what:="Hours", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no Hours column for Line One")
        Exit Sub
    Else
        LnOneHrsCol = foundDelta.Column
    End If
'Line Two
    Set foundDelta = LnTwoRng.Find(what:="SKU", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no SKU column for Line Two")
        Exit Sub
    Else
        LnTwoSKUCol = foundDelta.Column
    End If
    Set foundDelta = LnTwoRng.Find(what:="Batches", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no Batches column for Line Two")
        Exit Sub
    Else
        LnTwoBatCol = foundDelta.Column
    End If
    Set foundDelta = LnTwoRng.Find(what:="FG Dough", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no FG Dough column for Line Two")
        Exit Sub
    Else
        LnTwoFGDCol = foundDelta.Column
    End If
    Set foundDelta = LnTwoRng.Find(what:="Hours", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no Hours column for Line Two")
        Exit Sub
    Else
        LnTwoHrsCol = foundDelta.Column
    End If
'Line Three
    Set foundDelta = LnThreeRng.Find(what:="SKU", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no SKU column for Line Three")
        Exit Sub
    Else
        LnThreeSKUCol = foundDelta.Column
    End If
    Set foundDelta = LnThreeRng.Find(what:="Batches", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no Batches column for Line Three")
        Exit Sub
    Else
        LnThreeBatCol = foundDelta.Column
    End If
    Set foundDelta = LnThreeRng.Find(what:="FG Dough", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no FG Dough column for Line Three")
        Exit Sub
    Else
        LnThreeFGDCol = foundDelta.Column
    End If
    Set foundDelta = LnThreeRng.Find(what:="Hours", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no Hours column for Line Three")
        Exit Sub
    Else
        LnThreeHrsCol = foundDelta.Column
    End If
'Line Four
    Set foundDelta = LnFourRng.Find(what:="SKU", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no SKU column for Line Four")
        Exit Sub
    Else
        LnFourSKUCol = foundDelta.Column
    End If
    Set foundDelta = LnFourRng.Find(what:="Batches", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no Batches column for Line Four")
        Exit Sub
    Else
        LnFourBatCol = foundDelta.Column
    End If
    Set foundDelta = LnFourRng.Find(what:="FG Dough", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no FG Dough column for Line Four")
        Exit Sub
    Else
        LnFourFGDCol = foundDelta.Column
    End If
    Set foundDelta = LnFourRng.Find(what:="Hours", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no Hours column for Line Four")
        Exit Sub
    Else
        LnFourHrsCol = foundDelta.Column
    End If
'Line Five
    Set foundDelta = LnFiveRng.Find(what:="SKU", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no SKU column for Line Five")
        Exit Sub
    Else
        LnFiveSKUCol = foundDelta.Column
    End If
    Set foundDelta = LnFiveRng.Find(what:="Batches", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no Batches column for Line Five")
        Exit Sub
    Else
        LnFiveBatCol = foundDelta.Column
    End If
    Set foundDelta = LnFiveRng.Find(what:="FG Dough", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no FG Dough column for Line Five")
        Exit Sub
    Else
        LnFiveFGDCol = foundDelta.Column
    End If
    Set foundDelta = LnFiveRng.Find(what:="Hours", LookIn:=xlValues)
    If foundDelta Is Nothing Then
        MsgBox ("There is no Hours column for Line Five")
        Exit Sub
    Else
        LnFiveHrsCol = foundDelta.Column
    End If
'Find W1, W3, W5, W7, W9, W11, W13
    Set foundBeta = Columns(LnOneSKUCol).Find(what:="W1", LookIn:=xlValues)
    If foundBeta Is Nothing Then
        MsgBox ("Please correct Sheets to include W1 Series")
        Exit Sub
    Else
        w1Row = foundBeta.Row
    End If
    
    Set foundBeta = Columns(LnOneSKUCol).Find(what:="W3", LookIn:=xlValues)
    If foundBeta Is Nothing Then
        MsgBox ("Please correct Sheets to include W3 Series")
        Exit Sub
    Else
        w3Row = foundBeta.Row
    End If
    
    Set foundBeta = Columns(LnOneSKUCol).Find(what:="W5", LookIn:=xlValues)
    If foundBeta Is Nothing Then
        MsgBox ("Please correct Sheets to include W5 Series")
        Exit Sub
    Else
        w5Row = foundBeta.Row
    End If

    Set foundBeta = Columns(LnOneSKUCol).Find(what:="W7", LookIn:=xlValues)
    If foundBeta Is Nothing Then
        MsgBox ("Please correct Sheets to include W7 Series")
        Exit Sub
    Else
        w7Row = foundBeta.Row
    End If

    Set foundBeta = Columns(LnOneSKUCol).Find(what:="W9", LookIn:=xlValues)
    If foundBeta Is Nothing Then
        MsgBox ("Please correct Sheets to include W9 Series")
        Exit Sub
    Else
        w9Row = foundBeta.Row
    End If
    
    Set foundBeta = Columns(LnOneSKUCol).Find(what:="W11", LookIn:=xlValues)
    If foundBeta Is Nothing Then
        MsgBox ("Please correct Sheets to include W11 Series")
        Exit Sub
    Else
        w11Row = foundBeta.Row
    End If
    
    Set foundBeta = Columns(LnOneSKUCol).Find(what:="W13", LookIn:=xlValues)
    If foundBeta Is Nothing Then
        MsgBox ("Please correct Sheets to include W13 Series")
        Exit Sub
    Else
        w13Row = foundBeta.Row
    End If
    
    outputPrintLine = 2
    
    'Beginning of Printout of the Schedule
    For i = 1 To 5 ' For Number of Lines
        Select Case i
            Case 1:
                lineName = "LINE 1"
                skuCol = LnOneSKUCol
                batCol = LnOneBatCol
                FGDCol = LnOneFGDCol
                HrsCol = LnOneHrsCol
                outputCol = 3
            Case 2:
                lineName = "LINE 2"
                skuCol = LnTwoSKUCol
                batCol = LnTwoBatCol
                FGDCol = LnTwoFGDCol
                HrsCol = LnTwoHrsCol
                outputCol = 16
            Case 3:
                lineName = "LINE 3"
                skuCol = LnThreeSKUCol
                batCol = LnThreeBatCol
                FGDCol = LnThreeFGDCol
                HrsCol = LnThreeHrsCol
                outputCol = 29
            Case 4:
                lineName = "LINE 4"
                skuCol = LnFourSKUCol
                batCol = LnFourBatCol
                FGDCol = LnFourFGDCol
                HrsCol = LnFourHrsCol
                outputCol = 42
            Case 5:
                lineName = "LINE 5"
                skuCol = LnFiveSKUCol
                batCol = LnFiveBatCol
                FGDCol = LnFiveFGDCol
                HrsCol = LnFiveHrsCol
                outputCol = 55
        End Select
        iii = 3
        Sheets("scheduler").Activate
        For ii = w1Row To (w13Row + 12)
            If Sheets("ProductionPlannerScheduler").Cells(ii, skuCol).Value > 7000000 Then
               If Sheets("ProductionPlannerScheduler").Cells(ii, skuCol).Value < 8999999 Then
                    Sheets("scheduler").Cells(iii, outputCol).Value = Sheets("ProductionPlannerScheduler").Cells(ii, skuCol).Value
                    Sheets("scheduler").Cells(iii, outputCol + 6).Value = Sheets("ProductionPlannerScheduler").Cells(ii, batCol).Value
                    Sheets("scheduler").Cells(iii, outputCol).Select
                    Selection.Value = Selection.FormulaR1C1
                    Sheets("scheduler").Cells(iii, outputCol + 4).Value = Sheets("scheduler").Cells(iii, outputCol + 4).Value
                    If iii = 3 Then
                        Sheets("scheduler").Cells(iii, outputCol + 8).FormulaR1C1 = "=R[-2]C[-4]+R[-2]C[-2]+RC[-4]/60/24"
                        Sheets("scheduler").Cells(iii, outputCol + 9).FormulaR1C1 = "=+RC[-1]+(RC[-3]*RC[-6]/60/24)" '1*RC[-6]/60/24"
                    Else
                        Sheets("scheduler").Cells(iii, outputCol + 8).FormulaR1C1 = "=+R[-1]C[1]+R[-1]C[-1]/60/24"
                        Sheets("scheduler").Cells(iii, outputCol + 9).FormulaR1C1 = "=+RC[-1]+(RC[-3]*RC[-6]/60/24)" '1*RC[-6]/60/24"
                    End If
                    
                    iii = iii + 1
               End If
            End If
        Next
    Next
    
    If Sheets("scheduler").Cells(2, 69).Value = "LINE 1" Then
        Sheets("scheduler").Range("N3:Z50").ClearContents
        Sheets("scheduler").Range("AA3:AM50").ClearContents
        Sheets("scheduler").Range("AN3:AZ50").ClearContents
        Sheets("scheduler").Range("BA3:BM50").ClearContents
    ElseIf Sheets("scheduler").Cells(2, 69).Value = "LINE 2" Then
        Sheets("scheduler").Range("A3:M50").ClearContents
        Sheets("scheduler").Range("AA3:AM50").ClearContents
        Sheets("scheduler").Range("AN3:AZ50").ClearContents
        Sheets("scheduler").Range("BA3:BM50").ClearContents
    ElseIf Sheets("scheduler").Cells(2, 69).Value = "LINE 3" Then
        Sheets("scheduler").Range("A3:M50").ClearContents
        Sheets("scheduler").Range("N3:Z50").ClearContents
        Sheets("scheduler").Range("AN3:AZ50").ClearContents
        Sheets("scheduler").Range("BA3:BM50").ClearContents
    ElseIf Sheets("scheduler").Cells(2, 69).Value = "LINE 4" Then
        Sheets("scheduler").Range("A3:M50").ClearContents
        Sheets("scheduler").Range("N3:Z50").ClearContents
        Sheets("scheduler").Range("AA3:AM50").ClearContents
        Sheets("scheduler").Range("BA3:BM50").ClearContents
    ElseIf Sheets("scheduler").Cells(2, 69).Value = "LINE 5" Then
        Sheets("scheduler").Range("A3:M50").ClearContents
        Sheets("scheduler").Range("N3:Z50").ClearContents
        Sheets("scheduler").Range("AA3:AM50").ClearContents
        Sheets("scheduler").Range("AN3:AZ50").ClearContents
    ElseIf Sheets("scheduler").Cells(2, 69).Value = "ALL" Then
        'Do Nothing
    End If
    
End Sub
'execute function
Sub executeCalc()

Dim startMixDateAndTime, workingDate As Date
Dim outputCol As Integer
Dim lineText As String
Dim starterArrayList() As String
Dim doughLastRow, intDoughSearch, lastRowLine As Integer
Dim countStarterDough As Integer
Dim intTotalMixes As Double
Dim doughsize As Double
Dim printOutLastRow As Double
Dim i, ii, iii, iiii As Integer

'Setup Starter list based on the reference Table
doughLastRow = Sheets("DOUGH_DATA").Cells(Rows.Count, 1).End(xlUp).Row
countStarterDough = Application.WorksheetFunction.CountIf(Sheets("DOUGH_DATA").Range("G:G"), "YES")
ReDim starterArrayList(countStarterDough)
For i = 2 To doughLastRow
    If Sheets("DOUGH_DATA").Cells(i, 7).Value = "YES" Then
        starterArrayList(i - 2) = Sheets("DOUGH_DATA").Cells(i, 1).Text
    End If
Next
Sheets("preferment_calc").Range("B2:BA15000").ClearContents
For i = 1 To 5

Application.ScreenUpdating = False

    Select Case i
        Case 1:
            outputCol = 3
            lineText = "SL01"
        Case 2:
            outputCol = 16
            lineText = "SL02"
        Case 3:
            outputCol = 29
            lineText = "SL03"
        Case 4:
            outputCol = 42
            lineText = "SL04"
        Case 5:
            outputCol = 55
            lineText = "SL05"
    End Select
    
    lastRowLine = Sheets("scheduler").Cells(Rows.Count, outputCol).End(xlUp).Row
    ' This adds the amount to the initial page
    If lastRowLine > 2 Then
        For ii = 3 To lastRowLine
            Set foundDough = Sheets("DOUGH_DATA").Cells.Find(what:=Sheets("scheduler").Cells(ii, outputCol + 1).Value, LookIn:=xlValues)
            If foundDough Is Nothing Then
                MsgBox ("There is no Dough for" & lineText & " " & Sheets("scheduler").Cells(ii, outputCol + 2).Value)
                Exit Sub
            Else
               intDoughSearch = foundDough.Row
            End If
            
            workingDate = Sheets("scheduler").Cells(ii, outputCol + 8).Value - Sheets("scheduler").Cells(ii, outputCol + 4).Value / 60 / 24
            intTotalMixes = Sheets("scheduler").Cells(ii, outputCol + 6).Value
            doughsize = Sheets("scheduler").Cells(ii, outputCol + 5).Value
            intervalTime = Sheets("scheduler").Cells(ii, outputCol + 3).Value
            
            If (add_preferment(workingDate, intDoughSearch, intTotalMixes, starterArrayList, doughsize, intervalTime, lineText) = False) Then
    
                MsgBox "Error adding preferments for " & Worksheets("DOUGH_DATA").Cells(intDoughSearch, 2) & _
                        " mix at " & workingDate
                GoTo ResExit
            End If
        Next
        
        Sheets("preferment_calc").Activate
        Columns("B:G").Select
        ActiveWorkbook.Worksheets("preferment_calc").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("preferment_calc").Sort.SortFields.Add2 Key:=Range( _
            "D2:D15000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("preferment_calc").Sort
            .SetRange Range("B1:G15000")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
Next

Dim prefermentLastRow, noRefeedlastRow, reFeedLastRow, starterDoughRowSearch, lastRowRevised As Integer
Dim checkStarter As Integer
Dim uboundStarter, counterRow1, counterRow2 As Integer
Dim noRefeedSourceCol, reFeedSourceCol, printOutCol As Integer
Dim sequenceDateCheck As Double
Dim usageTime, starterMaxSize, starterfermentationTime, starterRefeedSize, starterRefeedPercentage, starterColRef As Double
Dim twoDLArrayNoRefeed(1 To 30, 1 To 6) As Variant
Dim usageArray() As String
Dim counterArrayX As Integer
Dim savedFirst As Integer
Dim binCount As Integer
Dim remainValueTestCase, remainValueTestCaseCompare As Double
Dim minBatchSizeToMax As Double
Dim matchLineArray As Boolean
Dim yxyx As Integer
Dim yyyy As Integer

prefermentLastRow = Sheets("preferment_calc").Cells(Rows.Count, 2).End(xlUp).Row
uboundStarter = UBound(starterArrayList)
noRefeedSourceCol = 8
reFeedSourceCol = 18
printOutCol = 38
For i = 0 To uboundStarter - 1
    counterRow1 = 2
    counterRow2 = 2
    Set foundStarter = Sheets("DOUGH_DATA").Columns(1).Find(what:=starterArrayList(i), LookIn:=xlValues)
    If foundStarter Is Nothing Then
        MsgBox ("The Starter Dough is not setup properly")
        Exit Sub
    Else
       starterDoughRowSearch = foundStarter.Row
       usageTime = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 5).Value
       starterMaxSize = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 3).Value
       starterfermentationTime = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 4).Value
       If Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 9).Value = "" Then
            MsgBox ("Please Fill in Starter Dough Data Column number")
            Exit Sub
       End If
       starterColRef = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 9).Value
       starterRefeedPercentage = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, starterColRef).Value
       starterRefeedSize = starterMaxSize * starterRefeedPercentage
       minBatchSizeToMax = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 10).Value
    End If
    'Require Change for column updates
    For ii = 2 To prefermentLastRow
        If Sheets("preferment_calc").Cells(ii, 2).Text = starterArrayList(i) Then
            If Sheets("preferment_calc").Cells(ii, 5).Value = "NO" Then
                remainValueTestCaseCompare = 10
                Sheets("preferment_calc").Cells(counterRow1, 8).Value = Sheets("preferment_calc").Cells(ii, 2).Value
                Sheets("preferment_calc").Cells(counterRow1, 9).Value = Sheets("preferment_calc").Cells(ii, 3).Value * 1.01 'One percent Scrap
                Sheets("preferment_calc").Cells(counterRow1, 10).Value = Sheets("preferment_calc").Cells(ii, 4).Value
                Sheets("preferment_calc").Cells(counterRow1, 11).Value = Sheets("preferment_calc").Cells(ii, 7).Value
                Sheets("preferment_calc").Cells(counterRow1, 12).Value = Sheets("preferment_calc").Cells(ii, 6).Value
                Sheets("preferment_calc").Cells(counterRow1, 17).Value = Sheets("preferment_calc").Cells(ii, 3).Value
                'Bin Determination
                 For iii = 0 To 5
                    remainValueTestCase = Sheets("preferment_calc").Cells(counterRow1, noRefeedSourceCol + 9).Value Mod (15 - iii)
                    If remainValueTestCase < remainValueTestCaseCompare Then
                        binCount = Application.WorksheetFunction.RoundDown(Sheets("preferment_calc").Cells(counterRow1, noRefeedSourceCol + 9).Value / (15 - iii), 0)
                        If binCount = 0 Then
                            binCount = 1
                        End If
                        Sheets("preferment_calc").Cells(counterRow1, noRefeedSourceCol + 6).Value = Round(Sheets("preferment_calc").Cells(counterRow1, noRefeedSourceCol + 9).Value / binCount, 2)
                        remainValueTestCaseCompare = remainValueTestCase
                    End If
                Next
                'End of Bin determination
                Sheets("preferment_calc").Cells(counterRow1, 15).Value = Sheets("preferment_calc").Cells(ii, 4).Value
                counterRow1 = counterRow1 + 1
            ElseIf Sheets("preferment_calc").Cells(ii, 5).Value = "YES" Then
                Sheets("preferment_calc").Cells(counterRow2, 18).Value = Sheets("preferment_calc").Cells(ii, 2).Value
                Sheets("preferment_calc").Cells(counterRow2, 19).Value = Sheets("preferment_calc").Cells(ii, 3).Value
                Sheets("preferment_calc").Cells(counterRow2, 20).Value = Sheets("preferment_calc").Cells(ii, 4).Value
                Sheets("preferment_calc").Cells(counterRow2, 22).Value = Sheets("preferment_calc").Cells(ii, 7).Value
                Sheets("preferment_calc").Cells(counterRow2, 27).Value = Sheets("preferment_calc").Cells(ii, 6).Value
                Sheets("preferment_calc").Cells(counterRow2, 34).Value = counterRow2 - 1
                counterRow2 = counterRow2 + 1
            Else
                MsgBox ("The Starter Dough STARTER REFEED Column has not been properly fill in")
                Exit Sub
            End If
        End If
    Next
'For OWD, BIGA AND POOLISH AND ADDITIONAL STARTER WITH THE REFEED INDICATOR MARK as NO

    If counterRow1 > 2 Then
        printOutLastRow = Sheets("preferment_calc").Cells(Rows.Count, 38).End(xlUp).Row + 1
        noRefeedlastRow = Sheets("preferment_calc").Cells(Rows.Count, noRefeedSourceCol).End(xlUp).Row
        For iii = 2 To noRefeedlastRow
            savedFirst = iii
            counterArrayX = 2
            Erase twoDLArrayNoRefeed
            'Format Array 1st
            If Sheets("preferment_calc").Cells(iii, noRefeedSourceCol + 1).Value <> "" Then
                twoDLArrayNoRefeed(1, 1) = Sheets("preferment_calc").Cells(iii, noRefeedSourceCol + 4).Value
                twoDLArrayNoRefeed(1, 2) = Sheets("preferment_calc").Cells(iii, noRefeedSourceCol + 6).Value
                twoDLArrayNoRefeed(1, 3) = Round(Sheets("preferment_calc").Cells(iii, noRefeedSourceCol + 1).Value / Sheets("preferment_calc").Cells(iii, noRefeedSourceCol + 6).Value, 0)
                twoDLArrayNoRefeed(1, 4) = Sheets("preferment_calc").Cells(iii, noRefeedSourceCol + 9).Value
                twoDLArrayNoRefeed(1, 5) = Round(Sheets("preferment_calc").Cells(iii, noRefeedSourceCol + 1).Value / Sheets("preferment_calc").Cells(iii, noRefeedSourceCol + 6).Value, 0)
                twoDLArrayNoRefeed(1, 6) = Sheets("preferment_calc").Cells(iii, noRefeedSourceCol + 1).Value
            End If
            'End of first array list
            sequenceDateCheck = Sheets("preferment_calc").Cells(iii, noRefeedSourceCol + 2).Value
            With Sheets("preferment_calc")
                For iiii = iii + 1 To noRefeedlastRow
                    If .Cells(iiii, noRefeedSourceCol + 2).Value > (sequenceDateCheck + ((usageTime - 0.5) / 60 / 24)) Then 'Time Check for 2 hours Usage Time
                        Exit For
                    Else
                        If .Cells(iiii, noRefeedSourceCol + 1).Value = 0 Then
                            Exit For
                        ElseIf .Cells(iii, noRefeedSourceCol + 1).Value < starterMaxSize Then 'Till here
                            .Cells(iii, noRefeedSourceCol + 1).Value = .Cells(iii, noRefeedSourceCol + 1).Value + .Cells(iiii, noRefeedSourceCol + 1).Value
                            .Cells(iii, noRefeedSourceCol + 2).Value = .Cells(iii, noRefeedSourceCol + 2).Value & " | " & .Cells(iiii, noRefeedSourceCol + 2).Value
                            If .Cells(iii, noRefeedSourceCol + 1).Value <= starterMaxSize Then
                                'Format Array
                                
                                matchLineArray = False
                                For yxyx = 1 To 30
                                    If twoDLArrayNoRefeed(yxyx, 1) = .Cells(iiii, noRefeedSourceCol + 4).Value Then
                                        If twoDLArrayNoRefeed(yxyx, 2) = .Cells(iiii, noRefeedSourceCol + 6).Value Then
                                            If twoDLArrayNoRefeed(yxyx, 3) = Round(.Cells(iiii, noRefeedSourceCol + 1).Value / .Cells(iiii, noRefeedSourceCol + 6).Value, 0) Then
                                                If twoDLArrayNoRefeed(yxyx, 4) = .Cells(iiii, noRefeedSourceCol + 9).Value Then
                                                    twoDLArrayNoRefeed(yxyx, 5) = twoDLArrayNoRefeed(yxyx, 5) + Round(.Cells(iiii, noRefeedSourceCol + 1).Value / .Cells(iiii, noRefeedSourceCol + 6).Value, 0)
                                                    twoDLArrayNoRefeed(yxyx, 6) = twoDLArrayNoRefeed(yxyx, 6) + .Cells(iiii, noRefeedSourceCol + 1).Value
                                                    matchLineArray = True
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                                If matchLineArray = False Then
                                    twoDLArrayNoRefeed(counterArrayX, 1) = .Cells(iiii, noRefeedSourceCol + 4).Value
                                    twoDLArrayNoRefeed(counterArrayX, 2) = .Cells(iiii, noRefeedSourceCol + 6).Value
                                    twoDLArrayNoRefeed(counterArrayX, 3) = Round(.Cells(iiii, noRefeedSourceCol + 1).Value / .Cells(iiii, noRefeedSourceCol + 6).Value, 0)
                                    twoDLArrayNoRefeed(counterArrayX, 4) = .Cells(iiii, noRefeedSourceCol + 9).Value
                                    twoDLArrayNoRefeed(counterArrayX, 5) = Round(.Cells(iiii, noRefeedSourceCol + 1).Value / .Cells(iiii, noRefeedSourceCol + 6).Value, 0)
                                    twoDLArrayNoRefeed(counterArrayX, 6) = .Cells(iiii, noRefeedSourceCol + 1).Value
                                    counterArrayX = counterArrayX + 1
                                End If
                                
                                
                                'End format Array
                                .Cells(iiii, noRefeedSourceCol).Value = ""
                                .Cells(iiii, noRefeedSourceCol + 1).Value = ""
                                .Cells(iiii, noRefeedSourceCol + 2).Value = ""
                                .Cells(iiii, noRefeedSourceCol + 4).Value = ""
                                .Cells(iiii, noRefeedSourceCol + 6).Value = ""
                                .Cells(iiii, noRefeedSourceCol + 9).Value = ""
                                If .Cells(iiii, noRefeedSourceCol + 3).Value = "NEW" Then 'Added NEW to the split
                                    .Cells(iiii, noRefeedSourceCol + 3).Value = ""
                                    .Cells(iii, noRefeedSourceCol + 3).Value = "NEW"
                                    .Cells(iii, noRefeedSourceCol + 7).Value = .Cells(iiii, noRefeedSourceCol + 7).Value
                                    .Cells(iiii, noRefeedSourceCol + 7).Value = ""
                                Else
                                    .Cells(iiii, noRefeedSourceCol + 3).Value = ""
                                    .Cells(iiii, noRefeedSourceCol + 7).Value = ""
                                End If
                            ElseIf .Cells(iii, noRefeedSourceCol + 1).Value > starterMaxSize Then
                                'Format Array
                                matchLineArray = False
                                For yxyx = 1 To 30
                                    If twoDLArrayNoRefeed(yxyx, 1) = .Cells(iiii, noRefeedSourceCol + 4).Value Then
                                        If twoDLArrayNoRefeed(yxyx, 2) = .Cells(iiii, noRefeedSourceCol + 6).Value Then
                                            If twoDLArrayNoRefeed(yxyx, 3) = Round((.Cells(iiii, noRefeedSourceCol + 1).Value - (.Cells(iii, noRefeedSourceCol + 1).Value - starterMaxSize)) / .Cells(iiii, noRefeedSourceCol + 6).Value, 0) Then
                                                If twoDLArrayNoRefeed(yxyx, 4) = .Cells(iiii, noRefeedSourceCol + 9).Value Then
                                                    twoDLArrayNoRefeed(yxyx, 5) = twoDLArrayNoRefeed(yxyx, 5) + Round((.Cells(iiii, noRefeedSourceCol + 1).Value - (.Cells(iii, noRefeedSourceCol + 1).Value - starterMaxSize)) / .Cells(iiii, noRefeedSourceCol + 6).Value, 0)
                                                    twoDLArrayNoRefeed(yxyx, 6) = twoDLArrayNoRefeed(yxyx, 6) + (.Cells(iiii, noRefeedSourceCol + 1).Value - (.Cells(iii, noRefeedSourceCol + 1).Value - starterMaxSize))
                                                    matchLineArray = True
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                                If matchLineArray = False Then
                                    twoDLArrayNoRefeed(counterArrayX, 1) = .Cells(iiii, noRefeedSourceCol + 4).Value
                                    twoDLArrayNoRefeed(counterArrayX, 2) = .Cells(iiii, noRefeedSourceCol + 6).Value
                                    twoDLArrayNoRefeed(counterArrayX, 3) = Round((.Cells(iiii, noRefeedSourceCol + 1).Value - (.Cells(iii, noRefeedSourceCol + 1).Value - starterMaxSize)) / .Cells(iiii, noRefeedSourceCol + 6).Value, 0)
                                    twoDLArrayNoRefeed(counterArrayX, 4) = .Cells(iiii, noRefeedSourceCol + 9).Value
                                    twoDLArrayNoRefeed(counterArrayX, 5) = Round((.Cells(iiii, noRefeedSourceCol + 1).Value - (.Cells(iii, noRefeedSourceCol + 1).Value - starterMaxSize)) / .Cells(iiii, noRefeedSourceCol + 6).Value, 0)
                                    twoDLArrayNoRefeed(counterArrayX, 6) = (.Cells(iiii, noRefeedSourceCol + 1).Value - (.Cells(iii, noRefeedSourceCol + 1).Value - starterMaxSize))
                                    counterArrayX = counterArrayX + 1
                                End If
                                'End format Array
                                '.Cells(iii, noRefeedSourceCol + 2).Value = .Cells(iii, noRefeedSourceCol + 2).Value & " | " & .Cells(iiii, noRefeedSourceCol + 2).Value
                                .Cells(iiii, noRefeedSourceCol + 1).Value = .Cells(iii, noRefeedSourceCol + 1).Value - starterMaxSize
                                .Cells(iii, noRefeedSourceCol + 1).Value = starterMaxSize
                                iii = iiii - 1
                                Exit For
                            End If
                        Else
                            MsgBox ("ERROR, Please re-identify Constraints for Starter" & starterArrayList(i))
                            Exit For
                        End If
                    End If
                Next
            End With
            'Print Array
            If Sheets("preferment_calc").Cells(savedFirst, noRefeedSourceCol + 1).Value <> "" Then
                For yyyy = 1 To counterArrayX
                    Sheets("preferment_calc").Cells(savedFirst, noRefeedSourceCol + 8).Value = Sheets("preferment_calc").Cells(savedFirst, noRefeedSourceCol + 8).Value & twoDLArrayNoRefeed(yyyy, 1) & " | " & twoDLArrayNoRefeed(yyyy, 2) & " | " & twoDLArrayNoRefeed(yyyy, 3) & " | " & twoDLArrayNoRefeed(yyyy, 4) & " | " & twoDLArrayNoRefeed(yyyy, 5) & " | " & Round(twoDLArrayNoRefeed(yyyy, 6), 2) & Chr(10)
                Next
            End If
            'End Print Array

        Next
        'Remove Blanks Spot
        Sheets("preferment_calc").Activate
        Sheets("preferment_calc").Range("H2:Q14000").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.Delete Shift:=xlUp
        'Clear the auto Wrap Text for the spread sheet
        Sheets("preferment_calc").Range("L2:L14000").Select
        Selection.WrapText = False
        'Clear the auto Wrap Text for the spread sheet
        Sheets("preferment_calc").Range("P2:P14000").Select
        Selection.WrapText = False
        'Clear the auto Wrap Text for the spread sheet
        Sheets("preferment_calc").Range("Q2:Q14000").Select
        Selection.WrapText = False
        'Print Upper and Lower Range of Date and Time
        lastRowRevised = Sheets("preferment_calc").Cells(Rows.Count, noRefeedSourceCol).End(xlUp).Row
        For ii = 2 To lastRowRevised

            If IsDate(Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 2).Value) = False Then
                 usageArray = Split(Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 2).Text, " | ")
                Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 5).Value = usageArray(0)
            Else
                Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 5).Value = Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 2).Value
            End If
            If Sheets("preferment_calc").Cells(ii, noRefeedSourceCol).Value <> "" Then
                If Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 1).Value < minBatchSizeToMax Then
                'Round to the nearest 10
                    Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 1).Value = Application.WorksheetFunction.RoundUp(Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 1).Value / 10, 0) * 10
                Else
                    Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 1).Value = starterMaxSize
                End If
             End If
             If Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 1).Value < 40 Then
                Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 1).Value = 40
             End If
        Next
        
        noRefeedlastRow = Sheets("preferment_calc").Cells(Rows.Count, noRefeedSourceCol).End(xlUp).Row
        For ii = 2 To noRefeedlastRow
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol).Value = Sheets("preferment_calc").Cells(ii, noRefeedSourceCol).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 1).Value = Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 1).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 2).Value = Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 2).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 3).Value = Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 3).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 4).Value = Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 4).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 5).Value = Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 5).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 6).Value = Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 7).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 7).Value = Sheets("preferment_calc").Cells(ii, noRefeedSourceCol + 8).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 8).Value = "NA"
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 9).Value = "NA"
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 10).FormulaR1C1 = "=VLOOKUP(RC[-10],DOUGH_DATA!C[-47]:C[-46],2,FALSE)"
        Next

    
        Sheets("preferment_calc").Range("H2:Q14000").ClearContents

    End If
    
'End of Starters without Refeed

'For SourSeed, LiquidSour, and Light Rye Levain AND ADDITIONAL STARTER WITH THE REFEED INDICATOR MARK as YES
'Regular formula as the standard but with modify standard of 1e-8 less becomes zero, then start from FT build Group and Minimum limit for the next group
'Compared to the Maximum standard

Dim reFeedCheckBlanks As Boolean
Dim reFeedLoopTrue As Boolean
Dim reFeedLoopTwoTrue As Boolean
Dim reFeedFirstCheck As Boolean
Dim reFeedValueOne, reFeedValueTwo As Double
Dim reFeedFirstTestOne, reFeedFirstTestTwo As Double
Dim reFeedValueSaved As Double
Dim shortTestValue1, shortTestValue2 As Double
Dim totalValueChecks As Double
Dim counterCheck As Integer
Dim savedDateandTime As Date
Dim yy, yyy, iiiii As Integer
Dim reFeedCheck As Boolean

reFeedLoopTrue = False
reFeedCheckBlanks = True

'reFeedSourceCol

    If counterRow2 > 2 Then
    
    
'Added1 is set in stone unless less than 10 min, else create its own starting point
'Add the Division based amount to divid Bins based on its set differential limit
'Did not change from the initial amount, will change starting from here

'Round all FT and added1 to the nearest 10 and 40 to give an impression of a 0.5 % - 1.0 % scrap and 1% added on the
        reFeedLastRow = Sheets("preferment_calc").Cells(Rows.Count, reFeedSourceCol).End(xlUp).Row
        For ii = 2 To reFeedLastRow
            remainValueTestCaseCompare = 10
            For iii = 0 To 5
                    
                remainValueTestCase = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value Mod (15 - iii)
                If remainValueTestCase < remainValueTestCaseCompare Then
                    binCount = Application.WorksheetFunction.RoundDown(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value / (15 - iii), 0)
                    If binCount = 0 Then
                        binCount = 1
                    End If
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 13).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value / binCount, 2)
                    'Sheets("preferment_calc").Cells(counterRow1, 16).Value = Round(Sheets("preferment_calc").Cells(counterRow1, 9).Value / Sheets("preferment_calc").Cells(counterRow1, 14).Value, 0)
                    remainValueTestCaseCompare = remainValueTestCase
                End If

            Next
            If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 13).Value = "" Then
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 13).Value = "NA"
            End If
            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value * 1.005 'refeed scrap
        Next
    
    
        reFeedFirstCheck = True 'Added as of part 2
        Do While reFeedCheckBlanks = True
            If Application.WorksheetFunction.CountA(Sheets("preferment_calc").Range("T:T")) = Application.WorksheetFunction.CountA(Sheets("preferment_calc").Range("U:U")) Then 'Change made for
                reFeedCheckBlanks = False
            Else
                reFeedCheckBlanks = True
            End If
            reFeedLoopTrue = True
            reFeedLoopTwoTrue = True
            ii = 2
    '        ssFTCheckPointRow = 2
            Do While reFeedLoopTrue = True
                iiii = 3
                reFeedLoopTwoTrue = True
                If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value <> "FT" Then
                    If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value <> "added" Then 'Change here PT 2
                        If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value <> "added1" Then 'Change here PT2
                            Do While reFeedLoopTwoTrue = True
                                If ii = 2 Then
                                    'shortTestValue1 = Sheets("preferment_calc").Cells((ii), (reFeedSourceCol + 1)).Value * 0.14
                                    'If shortTestValue1 > 0.0005 Then  ' Test to make is faster 04/24/2024 BUT CALCULATION WILL BE 0.3% ERROR Less
                                        Sheets("preferment_calc").Range("R2:AH2").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove 'Need manual change modify to 04/23/2024
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 2).Value - (starterfermentationTime / 60 / 24)
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 1).Value * starterRefeedPercentage ' Manual change %
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol).Value
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 9).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 9).Value
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "FT"
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 4).Value = "NA" 'Added for part2 as well
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 13).Value = "NA"
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 16).Value 'added for sequence number
                                    'End If ' Test to make is faster 04/24/2024
                                    If reFeedFirstCheck = True Then
                                        Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 3).Value = "added" 'Change here modify to 11/27/2023
                                    Else
                                        Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 3).Value = "added1"
                                    End If
                                    ii = ii + 1
                                    reFeedLoopTwoTrue = False
                                ElseIf ii <> 2 Then
                                    If (Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value - (starterfermentationTime / 60 / 24)) >= Sheets("preferment_calc").Cells(iiii - 1, reFeedSourceCol + 2).Value Then 'need manual change
                                        If (Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value - (starterfermentationTime / 60 / 24)) < Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 2).Value Then ' need manual change
                                            'To determine added or added1 and NA
                                            If reFeedFirstCheck = True Then
                                                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "added" 'Change here modify to 11/27/2023
                                                If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 4).Value = "" Then
                                                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 4).Value = "NA" 'Added Here "" check to see if NA is required 1
                                                End If
                                            Else
                                                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "added1"
                                                If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 4).Value = "" Then
                                                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 4).Value = "NA" 'Added Here "" check to see if NA is required 2
                                                End If
                                            End If
                                            'End of Adding NA
                                            'New Addition to test amount before adding
                                            shortTestValue2 = Sheets("preferment_calc").Cells((ii), (reFeedSourceCol + 1)).Value * starterRefeedPercentage
                                            If shortTestValue2 > 0.005 Then ' Test to make is faster 04/24/2024 BUT CALCULATION WILL BE 0.3% ERROR Less
                                                Sheets("preferment_calc").Range("R" & iiii & ":AH" & iiii).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove 'Need manual change
                                                Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 1).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 1).Value * starterRefeedPercentage ' Manual change % Limit Change to optimize process to reduce the number of iterations
                                                Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 2).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 2).Value - (starterfermentationTime / 60 / 24) 'need manual change
                                                Sheets("preferment_calc").Cells(iiii, reFeedSourceCol).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol).Value
                                                Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 9).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 9).Value
                                                Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 16).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 16).Value 'added for sequence number
                                                Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 13).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 13).Value
                                                If Sheets("preferment_calc").Cells(iiii - 1, reFeedSourceCol + 3).Value = "FT" Then
                                                    Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 3).Value = "FT"
                                                    Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 4).Value = "NA"
                                                    Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 13).Value = "NA"
                                                    ii = ii + 1
                                                Else
                                                    If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 2).Value - (starterfermentationTime / 60 / 24) > Sheets("preferment_calc").Cells(iiii - 1, reFeedSourceCol + 2).Value Then 'need manual change
                                                        Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 3).Value = "FT"
                                                        Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 4).Value = "NA"
                                                        Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 13).Value = "NA"
                                                    End If
                                                End If
                                            End If ' Test to make is faster 04/24/2024
                                            reFeedLoopTwoTrue = False
                                        End If
                                    End If
                                    If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 4).Value = "" Then 'added 4
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 4).Value = "NA" 'Added 3
                                    End If
                                End If
                                iiii = iiii + 1
                            Loop
                        End If
                    End If
                End If
            ii = ii + 1
            
            If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value = "" Then
                reFeedLoopTrue = False
            End If
            
            Loop
            
            If Application.WorksheetFunction.CountA(Sheets("preferment_calc").Range("T:T")) = Application.WorksheetFunction.CountA(Sheets("preferment_calc").Range("U:U")) Then 'Need manual change
                reFeedCheckBlanks = False
            Else
                reFeedCheckBlanks = True
            End If
            'Added as of Part Two
            reFeedFirstCheck = False
        Loop

        reFeedLastRow = Sheets("preferment_calc").Cells(Rows.Count, reFeedSourceCol).End(xlUp).Row
        yy = 65
        iiiii = 2
        For ii = 2 To reFeedLastRow
            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 12).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value
        Next
'Stop here to see absolute calculation without rounding


'Beginning of Maxing added1 and FT to ensure capacity is setup for preceding requirement
        For ii = 2 To reFeedLastRow
            If IsDate(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value) Then
                savedDateandTime = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value
            ElseIf Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value = "" Then
                GoTo nextIteration
            End If
            For iii = ii + 1 To reFeedLastRow
                If Sheets("preferment_calc").Cells(ii, reFeedSourceCol).Value = "" Then
                    Exit For
                Else
                    If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "FT" Then
                        If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 2).Value > (savedDateandTime + ((usageTime - 0.5) / 60 / 24)) Then
                            Exit For
                        Else
                            If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 3).Value = "FT" Then
                                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value + Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value
                                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value & " | " & Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 2).Value
                                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Value & " | " & Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 16).Value
                                reFeedValueTwo = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 16).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 12).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 9).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 5).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 6).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 7).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 8).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 10).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 11).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 13).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 4).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 3).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 2).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value = ""
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol).Value = ""
                            End If
                        End If
                    End If
                End If
            Next

            If IsDate(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value) = False Then
                usageArray = Split(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Text, " | ")
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value = usageArray(0)
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 11).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value + ((usageTime - 0.5) / 60 / 24)
            Else
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 11).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value + ((usageTime - 0.5) / 60 / 24)
            End If

            If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "FT" Then
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Application.WorksheetFunction.RoundUp(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value, 0)
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 5).Value = 0
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value, 2)
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value / starterRefeedPercentage, 3)
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value = Chr(yy)
                yy = yy + 1
            ElseIf Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "added1" Then
                'Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Application.WorksheetFunction.RoundUp(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value / starterRefeedSize, 0) * starterRefeedSize
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Application.WorksheetFunction.RoundUp(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value, 2)
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 5).Value = 0
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value, 2)
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value / starterRefeedPercentage, 2)
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value = "NA"
            ElseIf Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "added" Then
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value, 2)
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 5).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value = 0
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value = 0
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value = "NA"
            End If
nextIteration:
        Next

        'Remove Blanks Spot
        Sheets("preferment_calc").Activate
        Sheets("preferment_calc").Range("R2:AH14000").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.Delete Shift:=xlUp
'Clear all the
'Beginning of added1 formula using FT as a guiding point to start grouping(A,B,C,D) both added1 and added
reFeedLastRow = Sheets("preferment_calc").Cells(Rows.Count, reFeedSourceCol).End(xlUp).Row

Dim seqNum() As String

        For ii = 2 To reFeedLastRow
            Erase seqNum()
            If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "FT" Then
                If IsNumeric(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Value) = False Then
                    seqNum = Split(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Text, " | ")
                    For iii = 2 To reFeedLastRow
                        For iiii = 0 To UBound(seqNum)
                            If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 16).Value = seqNum(iiii) Then
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 8).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value
                            End If
                        Next
                    Next
                Else
                    If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value <> "NA" Then
                        For iii = 2 To reFeedLastRow
                            If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 16).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Value Then
                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 8).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value
                            End If
                        Next
                    End If
                End If
            End If
        Next
        
        For ii = 2 To reFeedLastRow
            If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value = "NA" Then
                If Sheets("preferment_calc").Cells(ii - 1, reFeedSourceCol + 2).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value Then
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value = Sheets("preferment_calc").Cells(ii - 1, reFeedSourceCol + 8).Value
                End If
            End If
        Next
'Another Section to Group the added1 continuation FT only group to batch 167 for Sourseed because of roundin factor of 0.0005 decimal qty
            Dim added1Checkloop As Boolean
            Dim cel As Range
            added1Checkloop = True
            Do While (added1Checkloop = True)
                For ii = 2 To reFeedLastRow
                    If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value <> "FT" Then
                        If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value = "NA" Then
                            If Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 8).Value <> "NA" Then
                                If Sheets("preferment_calc").Cells(ii - 1, reFeedSourceCol + 8).Value <> "NA" Then
                                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value = Sheets("preferment_calc").Cells(ii - 1, reFeedSourceCol + 8).Value
                                End If
                            End If
                        ElseIf Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value <> "NA" Then
                            For iii = 2 To reFeedLastRow
                                If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 16).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Value Then
                                    Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 8).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value
                                End If
                            Next
                        End If
                    End If
                Next
                For Each cel In Sheets("preferment_calc").Range("Z2:Z14000")
                    If cel.Value = "NA" Then
                        added1Checkloop = True
                        Exit For
                    Else
                        added1Checkloop = False
                    End If
                Next
            Loop
            
            
            
            
            For ii = 2 To reFeedLastRow
                If IsDate(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value) Then
                    savedDateandTime = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value
                ElseIf Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value = "" Then
                    GoTo nextiterationtwo
                End If
                If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value = "NA" Then
                    GoTo nextiterationtwo
                End If
    
                For iii = ii + 1 To reFeedLastRow
                    If Sheets("preferment_calc").Cells(ii, reFeedSourceCol).Value = "" Then
                        Exit For
                    Else
                        If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "added1" Then
                            If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 10).Value > (savedDateandTime + ((usageTime - 0.5) / 60 / 24)) Then
                                Exit For
                            Else
                                If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 8).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value Then
                                    If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 3).Value = "added1" Then
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value + Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value & " | " & Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 2).Value
                                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Value & " | " & Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 16).Value
                                        reFeedValueTwo = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 16).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 12).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 9).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 5).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 6).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 7).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 8).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 10).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 11).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 13).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 4).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 3).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 2).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value = ""
                                        Sheets("preferment_calc").Cells(iii, reFeedSourceCol).Value = ""
                                    End If
                                Else
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
                If IsDate(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value) = False Then
                    usageArray = Split(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Text, " | ")
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value = usageArray(0)
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 11).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value + ((usageTime - 0.5) / 60 / 24)
                Else
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 11).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value + ((usageTime - 0.5) / 60 / 24)
                End If
    
                If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "added1" Then
                    'Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Application.WorksheetFunction.RoundUp(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value / starterRefeedSize, 0) * starterRefeedSize
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Application.WorksheetFunction.RoundUp(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value, 1)
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 5).Value = 0
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value, 2)
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value / starterRefeedPercentage, 2)
                ElseIf Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "added" Then
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value, 2)
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 5).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value = 0
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value = 0
                End If
nextiterationtwo:
            Next
    
            Sheets("preferment_calc").Activate
            Sheets("preferment_calc").Range("R2:AH14000").Select
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Selection.Delete Shift:=xlUp
            
            reFeedLastRow = Sheets("preferment_calc").Cells(Rows.Count, reFeedSourceCol).End(xlUp).Row
            
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
'Document finding: Since method in balancing added1 becoming increasing difficult with no methodology in sight for completion, rounding to the nearest whole is concluded, with rounding, just need a check to round of the rounding
'Finally combining Usage requirement for all the Requirement

        
        Sheets("preferment_calc").Range("AH2:AI15000").ClearContents
        
        'preDetermination
        
        If (preDetermination(reFeedSourceCol, starterMaxSize, starterRefeedPercentage, reFeedLastRow, usageTime)) = False Then
               MsgBox ("Error formatting Refeed2")
            GoTo ResExit
        End If

'Also check time + usage time +1 Split time based on expected maximum
        
        Sheets("preferment_calc").Range("AF2:AI15000").ClearContents
'End of balancing
'This is the part where it starts adding, located in Module 4
        reFeedLastRow = Sheets("preferment_calc").Cells(Rows.Count, reFeedSourceCol).End(xlUp).Row
    
        If (formatRefeedPart2(usageTime, starterMaxSize, starterfermentationTime, starterRefeedPercentage, starterRefeedSize, minBatchSizeToMax, reFeedLastRow, reFeedSourceCol)) = False Then
               MsgBox ("Error formatting Refeed2")
            GoTo ResExit
        End If


        For ii = 2 To reFeedLastRow
            If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value > 0 Then
                If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 14).Value = "" Then
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 14).Value = "NA"
                End If
            End If
        Next
        
        'Remove Blanks Spot, Columns needs to change as well
        Sheets("preferment_calc").Activate
        Sheets("preferment_calc").Range("R2:AH14000").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.Delete Shift:=xlUp
        Sheets("preferment_calc").Activate
        reFeedLastRow = Sheets("preferment_calc").Cells(Rows.Count, reFeedSourceCol).End(xlUp).Row
'Clean up some mess
           For ii = 2 To reFeedLastRow
                If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value < starterMaxSize Then
                    savedDateandTime = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value
                    For iii = ii + 1 To reFeedLastRow
                        If Sheets("preferment_calc").Cells(ii, reFeedSourceCol).Value = "" Then
                            Exit For
                        Else
                            If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "added" Then
                                If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 10).Value > (savedDateandTime + ((usageTime - 0.5) / 60 / 24)) Then
                                    Exit For
                                Else
                                    If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 8).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value Then
                                        If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 3).Value = "added" Then
                                            If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value + Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value < starterMaxSize Then
                                                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value + Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value
                                                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value & " | " & Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 2).Value
                                                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Value & " | " & Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 16).Value
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 16).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 12).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 9).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 5).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 6).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 7).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 8).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 10).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 11).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 13).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 14).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 4).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 3).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 2).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value = ""
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol).Value = ""
                                            End If
                                        End If
                                    Else
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If Sheets("preferment_calc").Cells(ii, reFeedSourceCol).Value <> "" Then
                        If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "added" Then
                            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value, 2)
                            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 5).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value
                            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value = 0
                            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value = 0
                        End If
                    End If
                End If
            Next
        'Remove Blanks Spot, Columns needs to change as well
        Sheets("preferment_calc").Activate
        Sheets("preferment_calc").Range("R2:AH14000").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.Delete Shift:=xlUp
        Sheets("preferment_calc").Activate
'ENd of cleaning up some mess

'Second Parr of Balancing to the nearest five
        reFeedLastRow = Sheets("preferment_calc").Cells(Rows.Count, reFeedSourceCol).End(xlUp).Row
        For ii = 2 To reFeedLastRow
            If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value > 40 Then
                If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value < starterMaxSize Then
                    If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value > 0 Then
                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value - (Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value * starterRefeedPercentage)
                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value / starterRefeedPercentage, 2)
                        Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value + Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 5).Value
                    End If
                End If
            End If
        Next
        For ii = 2 To reFeedLastRow
            
            If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value < 40 Then
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = 40
            End If
            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Application.WorksheetFunction.RoundUp(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value / 5, 0) * 5
            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value - Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 5).Value
            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value / starterRefeedPercentage, 2)
        
        Next
        
        'Print Feed One
        For ii = 2 To reFeedLastRow + 100
            If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "FT" Then
                Sheets("preferment_calc").Range("R" & ii & ":AH" & ii).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol).Value
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = Round(Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 1).Value * starterRefeedPercentage, 2)
                If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value < 40 Then
                   Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = 40
                End If
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value = "FO"
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 10).Value - (starterfermentationTime / 60 / 24)
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value = "NA"
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 4).Value = "NA"
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 5).Value = "0"
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value / starterRefeedPercentage, 2)
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value = "NA"
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 9).Value = "NA"
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 11).Value = "NA"
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 12).Value = "NA"
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 13).Value = "NA"
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 14).Value = "NA"
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 15).Value = "NA"
                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 16).Value = "NA"
                ii = ii + 1
            End If
        Next
        'Sort
        reFeedLastRow = Sheets("preferment_calc").Cells(Rows.Count, reFeedSourceCol).End(xlUp).Row
        Range("R1:AH" & reFeedLastRow).Select
        ActiveWorkbook.Worksheets("preferment_calc").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("preferment_calc").Sort.SortFields.Add2 Key:=Range( _
            "AB2:AB" & reFeedLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("preferment_calc").Sort
            .SetRange Range("R1:AH" & reFeedLastRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        'Split value Greater than  starter Max
            For ii = 2 To reFeedLastRow + 100
                If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value > starterMaxSize Then
                    Sheets("preferment_calc").Range("R" & ii & ":AH" & ii).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    Sheets("preferment_calc").Range("R" & ii & ":AH" & ii).Value = Sheets("preferment_calc").Range("R" & ii + 1 & ":AH" & ii + 1).Value
                    Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value = starterMaxSize
                    Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 1).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 1).Value - starterMaxSize
                    Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 10).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 10).Value + 0.00694
                    If Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 1).Value < 40 Then
                       Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 1).Value = 40
                    End If
                    If Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 3).Value = "added1" Then
                        If Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 14).Value = "NA" Then
                            Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 14).Value = "Refeed Or RoundingScrap | 1 | " & Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 1).Value
                            Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 6).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 1).Value
                            Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 7).Value = Round(Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 6).Value / starterRefeedPercentage, 2)
                        End If
                        If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 14).Value = "NA" Then
                            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 14).Value = "Refeed Or RoundingScrap | 1 | " & Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value
                            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value
                            Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value = Round(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 6).Value / starterRefeedPercentage, 2)
                        End If
                    End If
                    
                End If
                ii = ii + 1
            Next
            
        'Print to Column
        reFeedLastRow = Sheets("preferment_calc").Cells(Rows.Count, reFeedSourceCol).End(xlUp).Row
        'Correction Comments
        For ii = 2 To reFeedLastRow
            If Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 3).Value <> "FT" Then
                If Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 3).Value <> "FO" Then
                    If Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 14).Value = "NA" Then
                        Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 14).Value = Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 9).Value & " | " & "Dough | " & Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 5).Value & " | Refeed | " & Sheets("preferment_calc").Cells(ii + 1, reFeedSourceCol + 6).Value
                    End If
                End If
            End If
        Next
        printOutLastRow = Sheets("preferment_calc").Cells(Rows.Count, 38).End(xlUp).Row + 1
        
        For ii = 2 To reFeedLastRow
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 1).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 1).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 2).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 2).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 3).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 4).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 4).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 13).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 5).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 6).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 12).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 7).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 14).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 8).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 9).Value = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 3).Value
            Sheets("preferment_calc").Cells(printOutLastRow + ii - 2, printOutCol + 10).FormulaR1C1 = "=VLOOKUP(RC[-10],DOUGH_DATA!C[-47]:C[-46],2,FALSE)"
        Next

    
        Sheets("preferment_calc").Range("R2:AH14000").ClearContents

     
    End If
    'End of Counter Two / known as refeeds

Application.ScreenUpdating = True

Next
    'Execute formatting
    numberingandUpperLimit
    'Excute make preferment Sheet
    printMixer_PrefermentSheet
'ClearEmpty Spot underneath the page
    Sheets("preferment_calc").Activate
    Rows("15001:15001").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Sheets("scheduler").Activate
Exit Sub

ResExit:
    MsgBox ("Please Message IT / Rick Lam for troubleshoot")
End Sub
Sub numberingandUpperLimit()
            
        'Add Numbering SEQUENCE BASED ON REQUIREMENT
        Dim x, i, ii, iii, yyy, numberingLastRow, cntNum As Integer
        Dim svdTimeCntSrce, svdTimeCntTrgt As Double
        Dim testValue As String
        Dim summarizeCol As Integer
        Dim LArray() As String
        Dim LArray2() As Date
        Dim LArray3() As String
        Dim lastCountofArray, dividerLineCnt, dividerSplitCnt As Integer
        Dim lineValue As String
        Dim splitValue As String
        Dim LArray4(0 To 30, 0 To 30) As Variant
        Dim testRemainder As Double
        Dim lastSavedValue As Double
        Dim refeedAmountCheck As Double
        Dim binCountRefeed As Double
        Dim lastBinValue As Double
        Dim yy As Integer
        Dim checkRefeed As String
        Dim placeHolderNumOne, placeHolderNumTwo As Double
        Dim phaseCheck As Boolean
        Dim summationValue As Double
        Dim checkRefeedtoPrntGrp As Boolean
        Dim binVarOne As Integer
        Dim usageTime As Double
        Dim fermTime As Double
        Dim checkUpperGapTime As Double
        Dim refeedscrapPercentage As Double
        Dim counterclean As Integer
        refeedscrapPercentage = 1.005
        binVarOne = 1
        summarizeCol = 38
        numberingLastRow = Sheets("preferment_calc").Cells(Rows.Count, summarizeCol).End(xlUp).Row
        
        cntNum = 1
        svdTimeCntSrce = 0
        
        With Sheets("preferment_calc")
            .Range("AX2:BD15000").ClearContents
            For x = 2 To numberingLastRow
                If .Cells(x, summarizeCol + 3).Value = "NEW" Then
                    If .Cells(x, summarizeCol + 6).Value <> "NA" Then
                        svdTimeCntTrgt = Application.WorksheetFunction.RoundDown(.Cells(x, summarizeCol + 6).Value, 0)
                        If .Cells(x, summarizeCol).Value <> .Cells(x - 1, summarizeCol).Value Then
                            cntNum = 1
                            svdTimeCntSrce = Application.WorksheetFunction.RoundDown(.Cells(x, summarizeCol + 6).Value, 0)
                        End If
                    End If
                    If svdTimeCntTrgt - svdTimeCntSrce >= 1 Then
                        cntNum = 1
                        svdTimeCntSrce = svdTimeCntTrgt
                    End If
                End If
                
                If .Cells(x, summarizeCol + 9).Value = "FT" Or .Cells(x, summarizeCol + 9).Value = "FO" Then
                    cntNum = 1
                    svdTimeCntSrce = 0
                End If

                If .Cells(x, summarizeCol + 9).Value <> "FT" Then
                    If .Cells(x, summarizeCol + 9).Value <> "FO" Then
                        Sheets("preferment_calc").Cells(x, summarizeCol - 1).Value = cntNum
                        cntNum = cntNum + 1
                    End If
                End If
            Next
        
        
            'Sort UsageTime
            For i = 2 To numberingLastRow
                If IsDate(.Cells(i, summarizeCol + 2).Value) = False Then
                    If .Cells(i, summarizeCol + 2).Value <> "NA" Then
                        LArray = Split(.Cells(i, summarizeCol + 2).Text, " | ")
                        lastCountofArray = UBound(LArray)
                        ReDim LArray2(0 To lastCountofArray)
                        For ii = 0 To lastCountofArray
                            LArray2(ii) = CDate(LArray(ii))
                        Next
                        If lastCountofArray = 0 Then
                        'Blanks if count = zero
                        Else
                            SortAr LArray2
                            .Cells(i, summarizeCol + 11).Value = Format(LArray2(lastCountofArray), "mm/dd/yyyy hh:mm:ss AM/PM")
                        End If
                    Else
                        .Cells(i, summarizeCol + 11).Value = .Cells(i, summarizeCol + 2).Value
                    End If
                Else
                    .Cells(i, summarizeCol + 11).Value = .Cells(i, summarizeCol + 2).Value
                End If
            Next
            'End of Sort Usage time to get the correct association
            'Draw an Array First than do a comparison
            For i = 2 To numberingLastRow
            
                .Cells(i, summarizeCol + 12).Value = ""
                Erase LArray
                Erase LArray3
                LArray = Split(.Cells(i, summarizeCol + 7).Text, Chr(10))
                dividerLineCnt = UBound(LArray)
                For ii = 0 To dividerLineCnt
                    LArray3 = Split(LArray(ii), " | ")
                    dividerSplitCnt = UBound(LArray3)
                    For iii = 0 To dividerSplitCnt
                        LArray4(ii, iii) = LArray3(iii)
                    Next

'                   'Test
                    For iii = 0 To dividerSplitCnt
                        If LArray4(ii, 0) <> "" Then
                            .Cells(i, summarizeCol + 12).Value = .Cells(i, summarizeCol + 12).Value & " " & LArray4(ii, iii)
                        End If
                    Next
                    If LArray4(ii, 0) <> "" Then
                        .Cells(i, summarizeCol + 12).Value = .Cells(i, summarizeCol + 12).Value & Chr(10)
                    End If
'End of Test
                Next
                Erase LArray4
            Next
'Start writing the +4 series Require fix for the remainder additions
            For i = 2 To numberingLastRow
                
                placeHolderNumTwo = 1 ' For number of bins
                summationValue = .Cells(i, summarizeCol + 1).Value 'To determine the summationVaue for Refeeds
                If IsNumeric(.Cells(i, summarizeCol + 4).Value) = True Then
                    placeHolderNumTwo = .Cells(i, summarizeCol + 4).Value
                End If
                .Cells(i, summarizeCol + 4).Value = ""
                .Cells(i, summarizeCol + 13).FormulaR1C1 = "=VLOOKUP(RC[-13],C[-49]:C[-46],4,FALSE)"
                checkRefeed = .Cells(i, summarizeCol + 13).Value
                matNumber = .Cells(i, summarizeCol).Value
                Erase LArray
                Erase LArray3
                LArray = Split(.Cells(i, summarizeCol + 12).Text, Chr(10))
                dividerLineCnt = UBound(LArray)
                yyy = 0
                For ii = 0 To dividerLineCnt
                    LArray3 = Split(LArray(ii), " ")
                    dividerSplitCnt = UBound(LArray3)
                    If Trim(LArray(ii)) <> "NA" Then
                        If dividerSplitCnt > 0 Then
                            'Check for Row with the same number
                            For yyy = 0 To 30
                                If LArray4(yyy, 1) = LArray3(1) Then
                                    If LArray4(yyy, 2) = LArray3(2) Then
                                        Exit For
                                    End If
                                End If
                            Next
                            If yyy = 31 Then
                                For yyy = 0 To 30
                                    If LArray4(yyy, 1) = "" Then
                                        If LArray4(yyy, 2) = "" Then
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            'If there is no same find the first blanks
                            For iii = 0 To dividerSplitCnt
                                If iii = 6 Then
                                    LArray4(yyy, 6) = CDbl(LArray4(yyy, 6) + LArray3(6))
                                Else 'Substitute other text if necessary
                                    LArray4(yyy, iii) = LArray3(iii)
                                End If
                            Next
                        End If
                    End If
                Next
                
                
                'Continuation SavedValue, Temporary erase 9 only because array has already been built
                usageTime = Application.WorksheetFunction.VLookup(.Cells(i, summarizeCol).Value, Sheets("DOUGH_DATA").Range("A1:E1000"), 5, False)
                If .Cells(i, summarizeCol).Value <> .Cells(i - 1, summarizeCol).Value Then
                        For ii = 0 To 30
                            LArray4(ii, 9) = 0
                        Next
                Else
                    If (.Cells(i, summarizeCol + 5).Value - .Cells(i - 1, summarizeCol + 5).Value) > (usageTime / 60 / 24) Then
                        For ii = 0 To 30
                            LArray4(ii, 9) = 0
                        Next
                    End If
                End If
                
                
                'End of Continuation SavedValue check
                'Also include usage time for the starter dough
                'Same ii array to to compare against two value in the written array
                Select Case checkRefeed
                    Case "NO"
                        For ii = 0 To 30
                            'Write and should only clear out value 6 once it is done
                            If LArray4(ii, 1) <> "" Then
                                If LArray4(ii, 9) > 0 Then
                                    LArray4(ii, 6) = LArray4(ii, 6) + LArray4(ii, 9)
                                    LArray4(ii, 9) = 0
                                End If
                                'Print
                                placeHolderNumOne = Application.WorksheetFunction.RoundDown(CDbl(LArray4(ii, 2)) * 1.01, 2)
                                LArray4(ii, 5) = Application.WorksheetFunction.RoundDown(LArray4(ii, 6) / placeHolderNumOne, 0)
                                LArray4(ii, 7) = CDbl(LArray4(ii, 6)) - (CDbl(LArray4(ii, 5)) * placeHolderNumOne)
                                LArray4(ii, 8) = CDbl(LArray4(ii, 6)) - (CDbl(LArray4(ii, 5)) * CDbl(LArray4(ii, 2)))
                                If LArray4(ii, 5) > 0 Then
                                    .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & "For Line " & LArray4(ii, 1) & " Divide " & LArray4(ii, 5) & " Bins at " & LArray4(ii, 2) & " KG and with " & Round(LArray4(ii, 7), 2) & " Remainder +/-" ' & Round(CDbl(LArray4(ii, 8)) - CDbl(LArray4(ii, 7)), 2) & " Scrap"
                                    .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & Chr(10)
                                End If
                                'Saved Remainder
                                If CDbl(LArray4(ii, 6)) >= Round(CDbl(LArray4(ii, 2) * 1.01), 2) Then
                                    LArray4(ii, 9) = Round(LArray4(ii, 6) - ((Application.WorksheetFunction.RoundDown(LArray4(ii, 6) / (Application.WorksheetFunction.RoundDown(LArray4(ii, 2) * (1.01), 2)), 0) * (Application.WorksheetFunction.RoundDown(LArray4(ii, 2) * (1.01), 2)))), 2)
                                Else
                                    LArray4(ii, 9) = LArray4(ii, 6)
                                End If
                            End If
                            LArray4(ii, 6) = 0
                        Next
                    Case "YES"
                        For ii = 0 To 30
                        'For Sequence that can be divided into line
                            If LArray4(ii, 1) <> "" Then
                                If LArray4(ii, 1) <> "NA" Then
                                    If LArray4(ii, 1) <> "Refeed" Then
                                        If Trim(LArray4(ii, 2)) <> "Dough" Then
                                            If LArray4(ii, 9) > 0 Then
                                                LArray4(ii, 6) = LArray4(ii, 6) + LArray4(ii, 9)
                                                summationValue = summationValue + LArray4(ii, 9)
                                                LArray4(ii, 9) = 0
                                            End If
                                            'Print
                                            placeHolderNumOne = Application.WorksheetFunction.RoundDown(CDbl(LArray4(ii, 2)) * refeedscrapPercentage, 2)
                                            LArray4(ii, 5) = Application.WorksheetFunction.RoundDown(LArray4(ii, 6) / placeHolderNumOne, 0)
                                            LArray4(ii, 7) = CDbl(LArray4(ii, 6)) - (CDbl(LArray4(ii, 5)) * placeHolderNumOne)
                                            LArray4(ii, 8) = CDbl(LArray4(ii, 6)) - (CDbl(LArray4(ii, 5)) * CDbl(LArray4(ii, 2)))
                                            If LArray4(ii, 6) > 0 Then
                                                .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & "For Line " & LArray4(ii, 1) & " Divide " & LArray4(ii, 5) & " Bins at " & LArray4(ii, 2) & " KG and with " & Round(LArray4(ii, 7), 2) & " Remainder +/-" '& Round(CDbl(LArray4(ii, 8)) - CDbl(LArray4(ii, 7)), 2) & " Scrap"
                                                .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & Chr(10)
                                            End If
                                            'Saved Remainder
                                            If CDbl(LArray4(ii, 6)) >= Round(CDbl(LArray4(ii, 2)) * refeedscrapPercentage, 2) Then
                                                LArray4(ii, 9) = Round(LArray4(ii, 6) - ((Application.WorksheetFunction.RoundDown(LArray4(ii, 6) / (Application.WorksheetFunction.RoundDown(LArray4(ii, 2), 2)), 0) * (Application.WorksheetFunction.RoundDown(LArray4(ii, 2), 2)))), 2)
                                            Else
                                                LArray4(ii, 9) = LArray4(ii, 6)
                                            End If
                                            summationValue = summationValue - Round(LArray4(ii, 5) * LArray4(ii, 2), 2) - Round(LArray4(ii, 7), 2)
                                            LArray4(ii, 6) = 0
                                        End If
                                    End If
                                End If
                            End If
                            If LArray4(ii, 1) <> "" Then
                                If LArray4(ii, 1) <> "NA" Then
                                    If LArray4(ii, 1) <> "Refeed" Then
                                        If Trim(LArray4(ii, 2)) = "Dough" Then
                                            .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & "Dough Require: " & LArray4(ii, 3)
                                            .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & Chr(10)
                                            summationValue = summationValue - LArray4(ii, 3)
                                            LArray4(ii, 1) = ""
                                            LArray4(ii, 2) = ""
                                            LArray4(ii, 3) = ""
                                            LArray4(ii, 4) = ""
                                            LArray4(ii, 5) = ""
                                            LArray4(ii, 6) = ""
                                            LArray4(ii, 7) = ""
                                            LArray4(ii, 8) = ""
                                            LArray4(ii, 9) = 0
                                        End If
                                    End If
                                End If
                            End If
                            'Counter Newly added ,make sure to comment to compare
                            If LArray4(ii, 1) <> "" Then
                                If LArray4(ii, 5) = 0 Then
                                    If LArray4(ii, 11) = "" Then
                                        LArray4(ii, 11) = 1
                                    Else
                                        LArray4(ii, 11) = LArray4(ii, 11) + 1
                                    End If
                                End If
                                If LArray4(ii, 11) <> "" Then
                                    If LArray4(ii, 11) > 1 Then
                                        LArray4(ii, 1) = ""
                                        LArray4(ii, 2) = ""
                                        LArray4(ii, 3) = ""
                                        LArray4(ii, 4) = ""
                                        LArray4(ii, 5) = ""
                                        LArray4(ii, 6) = ""
                                        LArray4(ii, 7) = ""
                                        LArray4(ii, 8) = ""
                                        LArray4(ii, 9) = 0
                                        LArray4(ii, 11) = 0
                                    End If
                                End If
                            End If
                            'End of Counter to see if it works
                        Next
                        If summationValue > 2.99 Then
                            If summationValue >= 13 Then
                                .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & "Refeed QTY and/or Scrap QTY: " & Application.WorksheetFunction.RoundDown(summationValue / 13, 0) & " Bin(s) at " & Round(summationValue / (Application.WorksheetFunction.RoundDown(summationValue / 13, 0)), 2) & " KG"
                            ElseIf summationValue < 13 Then
                                .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & "Refeed QTY and/or Scrap QTY: 1 Bin at " & Round(summationValue, 2) & " KG"
                            End If
                        End If
                End Select
                
                If .Cells(i, summarizeCol).Value <> .Cells(i + 1, summarizeCol).Value Then
                    Erase LArray4
                End If
                
            Next
            
            Erase LArray
            Erase LArray3
            Erase LArray4
        
'

'Beginnin of the correspongig sequence Should work, not added any change with except time as lookup diff from v.40
        For i = 2 To numberingLastRow
            usageTime = Application.WorksheetFunction.VLookup(.Cells(i, summarizeCol).Value, Sheets("DOUGH_DATA").Range("A1:E1000"), 5, False)
            fermTime = Application.WorksheetFunction.VLookup(.Cells(i, summarizeCol).Value, Sheets("DOUGH_DATA").Range("A1:D1000"), 4, False)
            checkRefeedtoPrntGrp = False
            Erase LArray
            Erase LArray3
            LArray = Split(.Cells(i, summarizeCol + 4).Text, Chr(10))
            dividerLineCnt = UBound(LArray)
            For ii = 0 To dividerLineCnt
                LArray3 = Split(LArray(ii), " ")
                dividerSplitCnt = UBound(LArray3)
                For iii = 0 To dividerSplitCnt
                    If LArray3(iii) = "Refeed" Then
                        checkRefeedtoPrntGrp = True
                    End If
                Next
            Next
            If .Cells(i, summarizeCol + 13).Value = "YES" Then
                If .Cells(i, summarizeCol + 9).Value <> "FO" Then
                    If .Cells(i, summarizeCol + 9).Value <> "FT" Then
                        If checkRefeedtoPrntGrp = True Then
                            .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & "Can be used for Batch No"
                            For ii = i + 1 To numberingLastRow
                                If .Cells(i, summarizeCol).Value = .Cells(ii, summarizeCol).Value Then
                                    If .Cells(ii, summarizeCol + 5).Value > (CDate(.Cells(i, summarizeCol + 5).Value + ((fermTime - 0.5) / 60 / 24))) Then
                                        If .Cells(ii, summarizeCol + 5).Value < (CDate(.Cells(i, summarizeCol + 5).Value + ((fermTime + usageTime - 0.5) / 60 / 24))) Then
                                            If .Cells(ii, summarizeCol + 8).Value = .Cells(i, summarizeCol + 8).Value Then
                                                .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & ", " & .Cells(ii, summarizeCol - 1).Value
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            End If

            If .Cells(i, summarizeCol + 13).Value = "YES" Then
                If .Cells(i, summarizeCol + 9).Value = "FT" Then
                    .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & "Can be used for Batch No"
                    For ii = i + 1 To numberingLastRow
                        If .Cells(i, summarizeCol).Value = .Cells(ii, summarizeCol).Value Then
                            If .Cells(ii, summarizeCol + 5).Value > (CDate(.Cells(i, summarizeCol + 5).Value + ((fermTime - 0.5) / 60 / 24))) Then
                                If .Cells(ii, summarizeCol + 5).Value < (CDate(.Cells(i, summarizeCol + 5).Value + ((fermTime + usageTime - 0.5) / 60 / 24))) Then
                                    If .Cells(ii, summarizeCol + 8).Value = .Cells(i, summarizeCol + 8).Value Then
                                        .Cells(i, summarizeCol + 4).Value = .Cells(i, summarizeCol + 4).Value & ", " & .Cells(ii, summarizeCol - 1).Value
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
            .Cells(i, summarizeCol + 14).Value = .Cells(i, summarizeCol + 5).Value - (fermTime / 60 / 24)

        Next
        .Range("AK2:BB15000").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


        Worksheets("preferment_calc").Activate
        ActiveWorkbook.Worksheets("preferment_calc").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("preferment_calc").Sort.SortFields.Add2 Key:=Range( _
            "AZ2:AZ15000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("preferment_calc").Sort
            .SetRange Range("AK1:BB15000")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        'Check for Duplicate

        For i = 2 To numberingLastRow
            .Cells(i, summarizeCol + 15).FormulaR1C1 = "=COUNTIF(C[-1],RC[-1])"
        Next

        For i = 2 To numberingLastRow
            If .Cells(i, summarizeCol + 15).Value > 1 Then
                If .Cells(i, summarizeCol + 13).Value = "NO" Then
                    usageTime = Application.WorksheetFunction.VLookup(.Cells(i, summarizeCol).Value, Sheets("DOUGH_DATA").Range("A1:E1000"), 5, False)
                    If (.Cells(i, summarizeCol + 11).Value - .Cells(i, summarizeCol + 5).Value) < (usageTime / 60 / 24) Then
                        .Cells(i, summarizeCol + 16).Value = (.Cells(i, summarizeCol + 11).Value - .Cells(i, summarizeCol + 5).Value) * 60 * 24
                    Else
                        .Cells(i, summarizeCol + 16).Value = 0
                    End If
                End If
            End If
        Next

        For i = 3 To numberingLastRow
            checkUpperGapTime = 0
            If .Cells(i, summarizeCol + 15).Value > 1 Then
                If .Cells(i, summarizeCol + 13).Value = "NO" Then
                    If .Cells(i, summarizeCol + 16).Value <> "" Then
                        If .Cells(i, summarizeCol + 16).Value > 0 Then
                            For ii = 1 To 5
                                If i - ii > 1 Then
                                    checkUpperGapTime = .Cells(i, summarizeCol + 5).Value - .Cells(i - ii, summarizeCol + 5).Value
                                End If
                                If checkUpperGapTime <> 0 Then
                                    Exit For
                                End If
                            Next
                            If (checkUpperGapTime * 60 * 24) < (15 / 60 / 24) Then
                                .Cells(i, summarizeCol + 16).Value = 0
                            Else
                                .Cells(i, summarizeCol + 16).Value = 0
                                .Cells(i, summarizeCol + 5).Value = .Cells(i, summarizeCol + 5).Value - (15 / 60 / 24)
                            End If
                        End If
                    End If
                End If
            End If

        Next

        
        
        End With
        
        'Clear the auto Wrap Text for the spread sheet
        Sheets("preferment_calc").Range("AP2:AP14000").Select
        Selection.WrapText = False


End Sub

Sub SortAr(arr() As Date)
    Dim Temp As Date
    Dim i As Long, j As Long

    For j = 2 To UBound(arr)
        Temp = arr(j)
        For i = j - 1 To 1 Step -1
            If (arr(i) <= Temp) Then GoTo 10
                arr(i + 1) = arr(i)
        Next i
        i = 0
10:     arr(i + 1) = Temp
    Next j
End Sub


Public Function add_preferment(ByVal dMixTime As Date, ByVal intSearchingDough As Integer, ByVal oneTotalMix As Integer, starterArray() As String, ByVal batchSize As Double, ByVal intTime As Double, ByVal linetxt As String) As Boolean
    On Error GoTo ErrHandle
    Dim blnCheckTotalMix As Boolean
    'Variables
    Dim dblBeingMixed As Double
    Dim dblMixRequired As Double
    Dim dtePrefermentMixTime As Date
    Dim strLookingFor As String
    Dim y, yy As Integer
    Dim upperArray As Integer
    Dim formattedUsageTime As Date
    Dim starterDoughSearch, starterDoughCol As Integer
    Dim outputRowcount As Integer
    Dim addedIntervalTime As Double
    'Found
    'Dim test As String
    
    
    formattedUsageTime = Format(Application.WorksheetFunction.MRound(dMixTime, "0:01"), "mm/dd/yyyy hh:mm:ss AM/PM")
    outputRowcount = Sheets("preferment_calc").Cells(Rows.Count, 2).End(xlUp).Row + 1
    upperArray = UBound(starterArray)
    addedIntervalTime = 0
    
    For y = 0 To upperArray - 1
    starterDoughSearch = 0
    starterDoughCol = 0
    addedIntervalTime = 0
        Set starterFoundArray = Sheets("DOUGH_DATA").Columns(1).Find(what:=starterArray(y), LookIn:=xlValues)
        If starterFoundArray Is Nothing Then
            MsgBox ("There is no Starter Dough for " & starterArray(y))
            GoTo FuncExit
        Else
           starterDoughSearch = starterFoundArray.Row
            If Sheets("DOUGH_DATA").Cells(starterDoughSearch, 9).Value > 0 Then
                starterDoughCol = Sheets("DOUGH_DATA").Cells(starterDoughSearch, 9).Value
            Else
                MsgBox ("There is no Starter Dough Col Ref number for" & starterArray(y))
                GoTo FuncExit
            End If
            'Currently zero percentage scrap facto
           If Sheets("DOUGH_DATA").Cells(intSearchingDough, starterDoughCol).Value > 0 Then
                For yy = 1 To oneTotalMix
                    Sheets("preferment_calc").Cells(outputRowcount, 2).Value = starterArray(y)
                    Sheets("preferment_calc").Cells(outputRowcount, 3).Value = Round(batchSize * Sheets("DOUGH_DATA").Cells(intSearchingDough, starterDoughCol).Value, 2) 'This control the amount, currently zero scrap
                    Sheets("preferment_calc").Cells(outputRowcount, 4).Value = formattedUsageTime + addedIntervalTime / 60 / 24
                    Sheets("preferment_calc").Cells(outputRowcount, 5).Value = Sheets("DOUGH_DATA").Cells(starterDoughSearch, 8).Value
                    Sheets("preferment_calc").Cells(outputRowcount, 6).Value = linetxt
                    If yy = 1 Then
                        Sheets("preferment_calc").Cells(outputRowcount, 7).Value = "NEW"
                    Else
                        Sheets("preferment_calc").Cells(outputRowcount, 7).Value = "NA"
                    End If
                    outputRowcount = outputRowcount + 1
                    addedIntervalTime = addedIntervalTime + intTime
                Next
            End If
           
        End If
        
    Next

        add_preferment = True
    
FuncExit:
    Exit Function
ErrHandle:
    add_preferment = False
    GoTo FuncExit
End Function
'print to sheet
Sub printMixer_PrefermentSheet()

Dim prefermentICount, shtReprintNum2, rowNumCopyPre, lastRowPreferment, yyy, prefermentRowind, lastRowPrintPrefer, counterCheckReNum, WHTFLRCOL As Integer
Dim starterDoughRowSearch, usageTime, starterMaxSize, starterfermentationTime, starterColRef, starterRefeedPercentage, starterRefeedSize, minBatchSizeToMax, interiorColourNum As Double
    
    
    Application.ScreenUpdating = False
    'clear preferment schedule
    Worksheets("preferment_calc").Activate
    
    'Other Variables
    lastRowPreferment = Sheets("preferment_calc").Cells(Rows.Count, 38).End(xlUp).Row
    rowNumCopyPre = 44 'Edit should be ending+1
    prefermentRowind = 13 ' Starting Row
    shtReprintNum2 = Application.WorksheetFunction.RoundUp(lastRowPreferment / 9, 0) 'Edit
    counterCheckReNum = 1
    
    'Set WhiteFlour Loc
    Set WHTFLR = Sheets("DOUGH_DATA").Rows(1).Find(what:="White Flour", LookIn:=xlValues)
    If WHTFLR Is Nothing Then
        MsgBox ("The WhiteFlour Ing is not setup properly")
        Exit Sub
    Else
        WHTFLRCOL = WHTFLR.Column
    End If
    Sheets("mixer_Preferment_All_Line").Activate
    Sheets("mixer_Preferment_All_Line").Rows("13:29").ClearContents 'Edit
    Sheets("mixer_Preferment_All_Line").Rows("13:29").Select 'Edit for row change
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets("mixer_Preferment_All_Line").Rows("44:15000").Select 'Should be ending +1
    Selection.Delete Shift:=xlUp
    
    For prefermentICount = 1 To shtReprintNum2
        Sheets("mixer_Preferment_All_Line").Activate
        Sheets("mixer_Preferment_All_Line").Rows("12:43").Select ' Edit
        Selection.Copy
        Sheets("mixer_Preferment_All_Line").Rows("" & rowNumCopyPre & ":" & rowNumCopyPre + 32 & "").Select 'Edit
        ActiveSheet.Paste
        Worksheets("mixer_Preferment_All_Line").HPageBreaks.Add before:=Range(Range("A" & rowNumCopyPre).Rows(1).Address)
        rowNumCopyPre = rowNumCopyPre + 32
    Next
    Sheets("mixer_Preferment_All_Line").Activate
    ActiveSheet.PageSetup.PrintArea = "$A$1:$AB" & rowNumCopyPre - 1
    
    For yyy = 2 To lastRowPreferment
        
        With Sheets("preferment_calc")
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 1).Value = .Cells(yyy, 37).Value
            If .Cells(yyy, 37).Value = 1 Then
                Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 2).Value = "MO NUMBER:" & Chr(10) & "__________" & Chr(10) & .Cells(yyy, 38).Value & Chr(10) & .Cells(yyy, 48).Value
            Else
                Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 2).Value = .Cells(yyy, 38).Value & Chr(10) & .Cells(yyy, 48).Value
            End If
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 5).Value = Format(.Cells(yyy, 43).Value, "mm/dd/yyyy hh:mm:ss AM/PM")
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value = .Cells(yyy, 39).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 23).Value = .Cells(yyy, 42).Value
            If .Cells(yyy, 47).Value = "FO" Then
                Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 26).Value = .Cells(yyy, 47).Value
            ElseIf .Cells(yyy, 47).Value = "FT" Then
                Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 26).Value = .Cells(yyy, 47).Value
            End If
            
            
            'Set found row to located fermentation time and usage time
            Set foundStarter = Sheets("DOUGH_DATA").Columns(1).Find(what:=.Cells(yyy, 38).Value, LookIn:=xlValues)
            If foundStarter Is Nothing Then
                MsgBox ("The Starter Dough is not setup properly")
                Exit Sub
            Else
               starterDoughRowSearch = foundStarter.Row
               usageTime = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 5).Value
               starterMaxSize = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 3).Value
               starterfermentationTime = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 4).Value
               If Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 9).Value = "" Then
                    MsgBox ("Please Fill in Starter Dough Data Column number")
                    Exit Sub
               End If
               starterColRef = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 9).Value
               starterRefeedPercentage = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, starterColRef).Value
               starterRefeedSize = starterMaxSize * starterRefeedPercentage
               minBatchSizeToMax = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 10).Value
               interiorColourNum = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, 1).Interior.Color
            End If
            
            If starterRefeedPercentage = "" Then
                starterRefeedPercentage = 0
            End If
            
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 4).Value = Format(.Cells(yyy, 43).Value - (starterfermentationTime / 60 / 24), "mm/dd/yyyy hh:mm:ss AM/PM")
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 6).Value = Format(.Cells(yyy, 43).Value + (usageTime / 60 / 24), "mm/dd/yyyy hh:mm:ss AM/PM")
            Sheets("mixer_Preferment_All_Line").Range("A" & prefermentRowind & ":AB" & prefermentRowind).Select
            Selection.Interior.Color = interiorColourNum
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 11).Value = Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value * starterRefeedPercentage
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 12).Value = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, WHTFLRCOL) * Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 13).Value = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, WHTFLRCOL + 1) * Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 14).Value = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, WHTFLRCOL + 2) * Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 15).Value = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, WHTFLRCOL + 3) * Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 16).Value = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, WHTFLRCOL + 4) * Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 17).Value = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, WHTFLRCOL + 5) * Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 18).Value = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, WHTFLRCOL + 6) * Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 19).Value = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, WHTFLRCOL + 7) * Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 20).Value = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, WHTFLRCOL + 8) * Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 21).Value = Sheets("DOUGH_DATA").Cells(starterDoughRowSearch, WHTFLRCOL + 9) * Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value = Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value
            Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 22).Value = Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 7).Value - Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 11).Value
            If .Cells(yyy, 37).Value = 1 Then
                Sheets("mixer_Preferment_All_Line").Cells(prefermentRowind, 1).Interior.Color = 255
            End If
'            If .Cells(yyy, 38).Value <> .Cells(yyy + 1, 38).Value Then
'                counterCheckReNum = 0
'                If (prefermentRowind + 3) Mod 26 = 0 Then
'                    prefermentRowind = prefermentRowind + 16
'                Else
'                    prefermentRowind = prefermentRowind + (26 - ((prefermentRowind + 3) Mod 26)) + 16
'                End If
'            Else
                If counterCheckReNum Mod 9 = 0 Then 'Edit to 9
                    prefermentRowind = prefermentRowind + 16
                Else
                    prefermentRowind = prefermentRowind + 2
                End If
'            End If

            
        End With
        
        counterCheckReNum = counterCheckReNum + 1
    Next
    
    
    Application.ScreenUpdating = True

   
End Sub

'print mixing and fritsch sheet
Sub printDoughAndFritsch()

   'On Error GoTo ErrHandle
    Dim strStartTime As String
    Dim workingDate As Date
    Dim intErrors As Integer
    Dim intReading As Integer
    Dim intWriting As Integer
    Dim strRef As String
    Dim blnCompletedFirst As Boolean
    Dim intDoughSearch As Integer
    Dim intIngredientSearch As Integer
    Dim intIngredientAdd As Integer
    Dim intMixes As Integer
    Dim intTotalMixes As Integer
    Dim intInterval As Integer
    Dim intFritschWriting As Integer
    Dim intFritschCounter As Integer
    Dim lngFritschMix As Double
    Dim lngFritschCumm As Long
    Dim intPackingWriting As Integer
    Dim intFermentation As Integer
    Dim intRequiredLines As Integer
    Dim intRequiredCounter As Integer
    Dim blnRequirementFound As Boolean
    Dim firstIntWriting As Integer
    Dim shtReprintNum As Integer
    Dim schMixesLastRow As Integer
    Dim schICount As Integer
    Dim mixDoughICount As Integer
    Dim rowNumCopyDough As Integer
    'Starter find
    Dim prefOWDLastRow As Integer
    Dim prefORGOWDLastRow As Integer
    Dim prefPOOLISHLastRow As Integer
    Dim prefBIGALastRow As Integer
    Dim prefSourSLastRow As Integer
    Dim prefLiquidSLastRow As Integer
    Dim prefLightRyeLLastRow As Integer
    Dim prefOrganicLiquidSLastRow As Integer
    Dim prefOrganicSourSLastRow As Integer
    Dim ii As Integer
    Dim SH_FRITSCH, SH_MIXER As String
    Dim refCol As Integer
    Dim lineRef As String
    Dim ingSavedValue(1 To 5) As Double
    Dim yy As Integer
    Dim remainValueTestCaseCompare As Double
    Dim remainValueTestCase As Double
    Dim binCount As Double
    Dim savedDoughrow As Double
    Dim savedworkingDate As Date
    Dim savedTotalMixDelta As Integer
    Dim addedValueDelta As Integer
    Dim changeoverAdd As Double

    
    Application.ScreenUpdating = False
    
For ii = 1 To 5
    intErrors = 0
    blnCompletedFirst = False
    intTotalMixes = 0
    intFritschWriting = 13 'New Format as of 05/08/2024
    rowNumCopyDough = 59
    intWriting = 14
    
    Select Case ii
        Case 1:
            SH_FRITSCH = "fritsch_schedule_Line_One"
            SH_MIXER = "Mixer_Report_Line_One"
            refCol = 3
            lineRef = "SL01"
        Case 2:
            SH_FRITSCH = "fritsch_schedule_Line_Two"
            SH_MIXER = "Mixer_Report_Line_Two"
            refCol = 16
            lineRef = "SL02"
        Case 3:
            SH_FRITSCH = "fritsch_schedule_Line_Three"
            SH_MIXER = "Mixer_Report_Line_Three"
            refCol = 29
            lineRef = "SL03"
        Case 4:
            SH_FRITSCH = "fritsch_schedule_Line_Four"
            SH_MIXER = "Mixer_Report_Line_Four"
            refCol = 42
            lineRef = "SL04"
        Case 5:
            SH_FRITSCH = "fritsch_schedule_Line_Five"
            SH_MIXER = "Mixer_Report_Line_Five"
            refCol = 55
            lineRef = "SL05"
    End Select
    
    With Worksheets(SH_FRITSCH)
        .Activate
        'clear old data
        'Change 12/08/2022 Rick change to FGF format
        .Rows("13:2000").Select
        Selection.Delete Shift:=xlUp
    End With

    'Clear Dough schedule
    With Worksheets(SH_MIXER)
        .Activate
        Worksheets(SH_MIXER).ResetAllPageBreaks
        'clear old data
        .Rows("59:1000").Select
        Selection.Delete Shift:=xlUp
        .Range("K14:AB14").Select
        Selection.ClearContents
        .Range("A15:AD40").Select
        Selection.ClearContents
        .Range("B10:B11").Select
        Selection.ClearContents
        .Range("E10:E11").Select
        Selection.ClearContents
        .Range("M10:M11").Select
        Selection.ClearContents
    End With

    If Sheets("scheduler").Cells(3, refCol).Value <> "" Then
    
    Else
        GoTo nxtLoopofIteration
    End If
    
    strStartTime = Sheets("scheduler").Cells(1, refCol + 4).Value + Sheets("scheduler").Cells(1, refCol + 6).Value
    workingDate = CDate(strStartTime)
    workingDate = DateAdd("s", 1, workingDate)
    
    'Insert Repearting Page of Dough Based on the # of Mix Divide 25
    shtReprintNum = 0
    schMixesLastRow = Sheets("scheduler").Cells(Rows.Count, refCol + 4).End(xlUp).Row
    For schICount = 3 To schMixesLastRow
        If Sheets("scheduler").Cells(schICount, refCol + 4).Value > 0 Then
            shtReprintNum = shtReprintNum + Application.WorksheetFunction.RoundUp(Sheets("scheduler").Cells(schICount, refCol + 4).Value / 25, 0)
        End If
    Next
    
    For mixDoughICount = 1 To shtReprintNum - 1
        Sheets(SH_MIXER).Activate
        Sheets(SH_MIXER).Rows("9:58").Select
        Selection.Copy
        Sheets(SH_MIXER).Rows("" & rowNumCopyDough & ":" & rowNumCopyDough + 49 & "").Select
        ActiveSheet.Paste
        Worksheets(SH_MIXER).HPageBreaks.Add before:=Range(Range("A" & rowNumCopyDough).Rows(1).Address)
        rowNumCopyDough = rowNumCopyDough + 50
    Next
    
    ActiveSheet.PageSetup.PrintArea = "$A$1:$AD" & rowNumCopyDough - 1
    savedDoughrow = 0
    savedworkingDate = 0
    savedTotalMixDelta = 0
    addedValueDelta = 0
    changeoverAdd = 0
    'Beginning of inputing Dough Mix, Fritsch Line and
    For intReading = 3 To 100
        'New SKU starts here
        If intReading = 3 Then
            changeoverAdd = 0
        Else
            changeoverAdd = (Sheets("scheduler").Cells(intReading - 1, refCol + 7).Value / 60 / 24) + (Sheets("scheduler").Cells(intReading - 1, refCol + 4).Value / 60 / 24) - (Sheets("scheduler").Cells(intReading, refCol + 4).Value / 60 / 24) + (Sheets("scheduler").Cells(intReading, refCol + 3).Value / 60 / 24)
        End If
        ingSavedValue(1) = 0
        ingSavedValue(2) = 0
        ingSavedValue(3) = 0
        ingSavedValue(4) = 0
        ingSavedValue(5) = 0
        strRef = Trim(Worksheets("scheduler").Cells(intReading, refCol).Value)
        If (Len(strRef) > 0) Then   'add the mix
            'insert recipe on the first row
            With Worksheets("DOUGH_DATA")
                intIngredientAdd = 11
                For intDoughSearch = 2 To 300
                    If (Worksheets("scheduler").Cells(intReading, refCol + 1).Value = .Cells(intDoughSearch, 1)) Then
                        For intIngredientSearch = 11 To 300
                            If (.Cells(1, intIngredientSearch).Value = "end") Then
                                Exit For
                            ElseIf (CDbl(.Cells(intDoughSearch, intIngredientSearch).Value) > 0) Then
                                firstIntWriting = intWriting
                                ' Added weird location but does the job
                                Worksheets(SH_FRITSCH).HPageBreaks.Add before:=Range(Range("A" & intFritschWriting).Rows(1).Address)
                                Worksheets(SH_FRITSCH).Cells(8, 11).Value = Application.WorksheetFunction.WeekNum(strStartTime)
                                Worksheets(SH_MIXER).Cells(8, 17).Value = Application.WorksheetFunction.WeekNum(strStartTime)
                                Worksheets(SH_MIXER).Cells(intWriting - 4, 2).Value = Sheets("scheduler").Cells(intReading, refCol - 1).Value
                                Worksheets(SH_MIXER).Cells(intWriting - 3, 2).Value = Sheets("scheduler").Cells(intReading, refCol + 1).Value
                                Worksheets(SH_MIXER).Cells(intWriting - 4, 5).FormulaR1C1 = "=IFERROR(VLOOKUP(R[1]C[-3],dough_data!C[-4]:C[-3],2,FALSE),"""")"
                                Worksheets(SH_MIXER).Cells(intWriting - 3, 5).Value = Sheets("scheduler").Cells(intReading, refCol + 2).Value
                                Worksheets(SH_MIXER).Cells(intWriting - 4, 13).Value = Sheets("scheduler").Cells(intReading, refCol).Value
                                Worksheets(SH_MIXER).Cells(intWriting - 3, 13).Value = Sheets("scheduler").Cells(intReading, refCol - 2).Value
                                '/////////////////////////////////////////
                                Worksheets(SH_MIXER).Cells(intWriting, intIngredientAdd).Value = .Cells(1, intIngredientSearch)
                                'Worksheets(SH_MIXER).Cells(intWriting, intIngredientAdd).Interior.Color = .Cells(1, intIngredientSearch).Interior.Color
                                Worksheets(SH_MIXER).Cells(intWriting + 1, intIngredientAdd).Value = _
                                        CDbl(.Cells(intDoughSearch, intIngredientSearch)) * _
                                        CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5).Value)
                                ' change from 10 to 2.1 kg 12/06/2022 - Rick
                                'if less than 2 set decimal to 3 places, else 1
                                Worksheets(SH_MIXER).Activate
                                Worksheets(SH_MIXER).Cells(intWriting + 1, intIngredientAdd).Select
                                If (CDbl(Worksheets(SH_MIXER).Cells(intWriting + 1, intIngredientAdd).Value) < 10) Then
                                    Selection.NumberFormat = "0.000"
                                Else
                                    Selection.NumberFormat = "0.0"
                                End If
                                ingSavedValue(intIngredientAdd - 10) = Worksheets(SH_MIXER).Cells(intWriting + 1, intIngredientAdd).Value
                                intIngredientAdd = intIngredientAdd + 1
                    
                            End If
                        Next
                        Exit For
                        
                    End If
                Next
            End With
            intWriting = intWriting + 2
            
            'calculate throughput
            intFritschCounter = 0
            lngFritschCumm = 0
            lngFritschMix = CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5)) / CDbl(Worksheets("scheduler").Cells(intReading, refCol + 10))
            
            
            'retrieve fermentation minutes
            intFermentation = CInt(Worksheets("scheduler").Cells(intReading, refCol + 4))
            'Update 02/16/2023
            'add the mixes
            'Mixes
            Dim sDblInterval As Double
            
            With Worksheets(SH_MIXER)
                intInterval = CInt(Worksheets("scheduler").Cells(intReading, refCol + 3))
                sDblInterval = Worksheets("scheduler").Cells(intReading, refCol + 3) * 60
                intTotalMixes = 0
                
                For intMixes = 1 To CInt(Worksheets("scheduler").Cells(intReading, refCol + 6).Value)
                    If intMixes = 1 Then
                        workingDate = workingDate + changeoverAdd
                    End If
                    If (intMixes > 1) Then workingDate = DateAdd("s", sDblInterval, workingDate)
                    intTotalMixes = intTotalMixes + 1
                    
                    'Saved Total Mix Delta Newly added
                    
                    If intTotalMixes = 1 Then
                        If savedDoughrow = intDoughSearch Then
                            If savedworkingDate = CDate(Int(workingDate + intFermentation / 60 / 24)) Then
                                savedTotalMixDelta = savedTotalMixDelta + Worksheets("scheduler").Cells(intReading - 2, refCol + 6).Value
                            End If
                        Else
                            savedTotalMixDelta = 0
                        End If
                    End If
                                        
                    'Newly added 04/09/2024///////////////////////
                    If savedDoughrow = intDoughSearch Then
                        If intTotalMixes = 1 Then
                            If savedworkingDate = CDate(Int(workingDate + intFermentation / 60 / 24)) Then
                                addedValueDelta = savedTotalMixDelta
                            Else
                                addedValueDelta = 0
                            End If
                        End If
                    Else
                        addedValueDelta = 0
                    End If
                    'Newlyadded ending
                    'Newly added 04/09/2024 for only addedvaluedelat

                    .Cells(intWriting, 1) = intTotalMixes + addedValueDelta ' Here to change for batch continuation
                    .Cells(intWriting, 2) = Format(workingDate, "dddd mmm/d/yy h:mm AM/PM")
                    '.Cells(intWriting, COL_M_DOUGH_NAME) = Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_NAME) & _
                            " (#" & Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_REF) & ")" & _
                              " ( " & Worksheets("scheduler").Cells(intReading, COL_SAP_PROCESS_ORDER_NUMBER).Value & ")"
                    .Cells(intWriting, 3) = Round(Worksheets("scheduler").Cells(intReading, refCol + 5), 0)
                    
                    
                    
                    
                    'Input the Number of Bins for the corresponding starters
                    For yy = 1 To UBound(ingSavedValue)
                        remainValueTestCaseCompare = 10
                        If ingSavedValue(yy) > 0 Then
                            For iii = 0 To 5
                                remainValueTestCase = ingSavedValue(yy) Mod (15 - iii)
                                If remainValueTestCase < remainValueTestCaseCompare Then
                                    binCount = Application.WorksheetFunction.RoundDown(ingSavedValue(yy) / (15 - iii), 0)
                                    If binCount = 0 Then
                                        binCount = 1
                                    End If
                                    remainValueTestCaseCompare = remainValueTestCase
                                End If
                            Next
                            .Cells(intWriting, yy + 10).Value = binCount & " Bin(s)"
                        End If
                    Next
                    'format the line
                    Sheets(SH_MIXER).Activate
                    format_line intWriting
                    '''
                    'Might need to add preferment here but to preferment_calc Rahter than directly to mixer_Preferment Sheet
                    'You can add 1 % scrap factor to this by simply adding *1.01 at the end of the line like this
                    'CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5)) * CDbl(Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_OWD))*1.01
'                    If (add_preferment(workingDate, _
'                                    CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5)) * CDbl(Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_OWD)) * 1.01, _
'                                    CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5)) * CDbl(Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_SEED)) * 1.01, _
'                                    CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5)) * CDbl(Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_LIQUID)) * 1.01, _
'                                    CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5)) * CDbl(Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_BIGA)) * 1.01, _
'                                    CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5)) * CDbl(Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_POOLISH)) * 1.01, _
'                                    CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5)) * CDbl(Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_ORGANIC_OWD)) * 1.01, _
'                                    CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5)) * CDbl(Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_ORGANIC_LIQUID)) * 1.01, _
'                                    CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5)) * CDbl(Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_ORGANIC_SEED)) * 1.01, _
'                                    CDbl(Worksheets("scheduler").Cells(intReading, refCol + 5)) * CDbl(Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_LIGHTRYELEVAIN)) * 1.01, _
'                                    Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_NAME), intTotalMixes) = False) Then
'
'                        MsgBox "Error adding preferments for " & Worksheets("DOUGH_DATA").Cells(intDoughSearch, COL_DD_NAME) & _
'                                " mix at " & workingDate
'                        GoTo ResExit
'                    End If
                    'Added to Warn Mixer not to mix to previous batches
                    'Dim lastRowPreferment As Integer
                    'If intTotalMixes = 1 Then
                    '    lastRowPreferment = Sheets(SH_PREFERMENT).Cells(Rows.Count, 5).End(xlUp).Row
                    '    Sheets(SH_PREFERMENT).Cells(lastRowPreferment, 7).Value = "*"
                   ' End If

                    'Modify 2023
                    'Begining of the formatting the fritsch line
                    'add it to the fritsch schedule
                    intFritschCounter = intFritschCounter + 1
                    lngFritschCumm = lngFritschCumm + lngFritschMix
                    
                    With Worksheets(SH_FRITSCH)
                        .Cells(intFritschWriting, 1).Value = intFritschCounter + addedValueDelta 'Here to change for batch continuation 2
                        .Cells(intFritschWriting, 2).Value = Worksheets("scheduler").Cells(intReading, refCol) & " " & Worksheets("scheduler").Cells(intReading, refCol + 2).Value
                        .Cells(intFritschWriting, 4).Value = FormatDateTime(DateAdd("n", intFermentation, CDate(workingDate)), vbShortTime)
                        .Cells(intFritschWriting, 5).Value = Worksheets(SH_MIXER).Cells(intWriting, 3)
                        .Cells(intFritschWriting, 6).Value = lngFritschCumm
                        .Cells(intFritschWriting, 3).Value = Format(CDate(workingDate) + (intFermentation / 60 / 24), "dddd mmm/d/yy")
                        Dim z As Integer
                        'Additional as of 08/23/2023 /////////////////////////////////////////////
                        'Added 08/18/2023 for Fritsch Line & 05/08/2024
                        If .Cells(intFritschWriting, 1).Value = 1 Then
                            .Cells(intFritschWriting, 11).Value = (60 / (sDblInterval / 60)) * .Cells(intFritschWriting, 6).Value
                            z = 1
                        End If
                        '08/28/2023 for Fritsch line Comments
                        If .Cells(intFritschWriting, 1).Value = z Then
                            If z <> 1 Then
                                .Cells(intFritschWriting, 11).Value = "Check all belts for damages. [  ]"
                            End If
                            z = z + Round((60 / (sDblInterval / 60)), 0)
                        End If
                        '////////////////////////////////////////////////////////////////////////////////////
                    End With
                    intFritschWriting = intFritschWriting + 1
                    
                    If intMixes = CInt(Worksheets("scheduler").Cells(intReading, refCol + 6).Value) Then
                        Exit For
                    End If
                    If intMixes Mod 25 = 0 Then 'must be whole
                        Application.DisplayAlerts = False
                        Worksheets(SH_MIXER).Rows("" & firstIntWriting - 4 & ":" & firstIntWriting + 1 & "").Copy
                        Worksheets(SH_MIXER).Rows(intWriting + 20).Select
                        ActiveSheet.Paste
                        intWriting = intWriting + 26
                        Application.DisplayAlerts = True
                    Else
                        intWriting = intWriting + 1
                    End If
                    
                    'Here to addifStatement for Referesh 04/09/2024
                    If intTotalMixes = 1 Then
                        savedDoughrow = intDoughSearch
                        savedworkingDate = CDate(Int(workingDate + intFermentation / 60 / 24))
                    End If
                    '/////////////////
                    
                Next
                
            If ((intMixes) Mod 25) > 0 Then
                intWriting = intWriting + (25 - ((intMixes - 1) Mod 25)) + 23
            Else
                intWriting = intWriting + 24
            End If
                    
                
            End With
            
        Else 'check for changeover
            If (Trim(Worksheets("scheduler").Cells(intReading, refCol + 3).Value) <> "") Then
                intInterval = CInt(Worksheets("scheduler").Cells(intReading, refCol + 3))
                workingDate = DateAdd("n", intInterval, workingDate)
            End If
        End If
    Next

    'format the fritsch schedule
    Sheets(SH_FRITSCH).Activate
    format_fritsch intFritschWriting - 1
    Sheets(SH_FRITSCH).Activate
    Sheets(SH_FRITSCH).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$K" & intFritschWriting - 1
    ActiveSheet.HPageBreaks(1).DragOff Direction:=xlUp, RegionIndex:=1
nxtLoopofIteration:
Next

Application.ScreenUpdating = True
Exit Sub

ErrHandle:
    MsgBox "Error creating production.  Please check data and try again.  If problem persists please email file and description of error to Rick.Lam@fgfbrands.com", vbOKOnly
   


End Sub
Private Sub format_line(ByVal lCurrentLine As Long)
    Dim intColumnCounter As Integer

    ActiveSheet.Range("A" & lCurrentLine & ":AA" & lCurrentLine).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    ActiveSheet.Range("A" & lCurrentLine).Select


End Sub

Public Function format_fritsch(ByVal lTotalLines As Long) As Boolean
    On Error GoTo ErrHandle
    
    With ActiveSheet
        
        'Change 12/08/2022 -Rick - change 5 to 13
        'format cells borders
        .Range("A13:K" & lTotalLines).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        If (lTotalLines > 1) Then
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End If
        .Range("A1").Select
    End With

    format_fritsch = True
    
FuncExit:
    Exit Function
ErrHandle:
    format_fritsch = False
    GoTo FuncExit
End Function
Sub cmdPrint_Click()
    Worksheets("mixer_Preferment_All_Line").Activate
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    'print the schedules
Dim SH_MIXER As String
Dim SH_FRITSCH As String
Dim ii As Integer

For ii = 1 To 5
    Select Case ii
        Case 1:
            SH_FRITSCH = "fritsch_schedule_Line_One"
            SH_MIXER = "Mixer_Report_Line_One"

        Case 2:
            SH_FRITSCH = "fritsch_schedule_Line_Two"
            SH_MIXER = "Mixer_Report_Line_Two"

        Case 3:
            SH_FRITSCH = "fritsch_schedule_Line_Three"
            SH_MIXER = "Mixer_Report_Line_Three"
  
        Case 4:
            SH_FRITSCH = "fritsch_schedule_Line_Four"
            SH_MIXER = "Mixer_Report_Line_Four"

        Case 5:
            SH_FRITSCH = "fritsch_schedule_Line_Five"
            SH_MIXER = "Mixer_Report_Line_Five"
    End Select
    
    Worksheets(SH_MIXER).Activate
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
    Worksheets(SH_FRITSCH).Activate
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Next
 '   Worksheets(SH_PACKING).Activate
  ' ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
 '
'    Worksheets(SH_REQUIRE).Activate
 '   ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
'    With Worksheets(SH_INVENTORY)
'        .Activate
'        Call .cmdUnhide_Click
'        Call .cmdHide_Click
'    End With
'    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    

End Sub

'Another part of the calculation
Public Function formatRefeedPart2(ByVal usageTimeRF2 As Double, ByVal starterMaxSizeRF2 As Double, ByVal starterFermentationTimeRF2 As Double, ByVal starterRefeedPercentageRF2 As Double, ByVal starterRefeedSizeRF2 As Double, ByVal mibBatchSizeToMaxRF2 As Double, ByVal reFeedLastRowRF2 As Integer, ByVal reFeedSourceCol As Integer) As Boolean

' Beginning to TEST NEW VERSION //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'Step:Find Source, Check against Source for Start match,Check if it can max once it reaches to limit of bowl, Add based on the added1 and added rule
    'On Error GoTo ErrHandleSecond
    
    Dim threeDLArray As Variant
    Dim reFeedValueOne, reFeedValueTwo As Double
    Dim endForLoopCheck As Boolean
    Dim endForLoopCheckOne As Boolean
    Dim ii, iii, iiii, yy As Integer
    Dim savedValue, savedValueDelta As Double
    Dim savedDateandTime, savedDateandTimeSecond As Date
    Dim returnArrayVar As Variant
    Dim testValueUnderTrgt As Double
    Dim starterMaxDelta As Double
    Dim grpChrSrce, grpChrTrgt As String
    Dim groupLetterSouce As String
    Dim groupLetterTrgt, groupLetterTrgt2 As String
    Dim counterTrgt As Integer
    Dim groupCheck As Boolean
    Dim varDiffAdd As Double
    Dim varAdd As Double
    Dim lineNameGroup As String
    Dim binSizeDet As Double
    Dim returnText As Variant
    Dim twoDLArray(1 To 30, 1 To 6) As Variant
    Dim yxy, x2 As Integer
    Dim numberOfBins As Double
    Dim qtyToBeAddedtoLarr As Double
    Dim checkMatchArr As Boolean
    Dim counterXXArray As Integer
    Dim addedOneTFCheck As Boolean
    Dim xxy As Integer
    Dim rowCountVerThree As Integer
    rowCountVerThree = 3
    For ii = 2 To reFeedLastRowRF2
    counterXXArray = 2
    'First loop to determine prior volume to be consumed
        If Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value <> 0 Then 'Check source does have an amount
            savedValue = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value
            savedDateandTime = CDate(Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 10).Value)
            grpChrSrce = Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 8).Value
            For iii = rowCountVerThree To reFeedLastRowRF2
                reFeedValueOne = 0
                reFeedValueTwo = 0
                If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 3).Value <> "FT" Then 'Check first is build on top is not FT
                    If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value <> "" Then
                        If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value < starterMaxSizeRF2 Then 'Check to max Size
                            If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value < savedValue Then ' Check if current has enough
                                If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 10).Value >= (savedDateandTime + (starterFermentationTimeRF2 / 24 / 60)) Then
                                    If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 10).Value < (savedDateandTime + ((starterFermentationTimeRF2 + usageTimeRF2) / 24 / 60)) Then 'Check time is within range for the first added or added1
                                        grpChrTrgt = Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 8).Value
                                        If grpChrSrce = grpChrTrgt Then
                                            savedValueDelta = savedValue
                                            savedDateandTimeSecond = CDate(Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 10).Value)
                                            'Group Limit
                                            testValueUnderTrgt = Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value
                                            counterTrgt = 0
                                            'This is to check that only one next group will be added to the prior one
                                            For iiii = iii + 1 To reFeedLastRowRF2
                                                groupLetterTrgt2 = Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 8).Value
                                                If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 3).Value <> "FT" Then
                                                    If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 1).Value <> "" Then
                                                        If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 1).Value <= savedValue Then
                                                            If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 10).Value > (savedDateandTimeSecond + (usageTimeRF2 - 0.05) / 60 / 24) Then
                                                                Exit For
                                                            Else
                                                                If grpChrTrgt <> groupLetterTrgt2 Then
                                                                    grpChrTrgt = groupLetterTrgt2
                                                                    counterTrgt = counterTrgt + 1
                                                                End If
                                                                If counterTrgt = 2 Then
                                                                    Exit For
                                                                End If
                                                                'added Value
                                                                If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 3).Value = "added1" Then
                                                                    If (Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 10).Value - Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 10).Value) < 0.00694 Then
                                                                        testValueUnderTrgt = testValueUnderTrgt + Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 1).Value
                                                                    Else
                                                                        GoTo NextIterationFour
                                                                    End If
                                                                Else
                                                                    testValueUnderTrgt = testValueUnderTrgt + Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 1).Value
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
NextIterationFour:
                                            Next
                                            If savedValue < starterMaxSizeRF2 Then
                                                If testValueUnderTrgt >= starterMaxSizeRF2 Then
                                                    groupCheck = False
                                                Else
                                                    groupCheck = True
                                                End If
                                            ElseIf savedValue >= starterMaxSizeRF2 Then
                                                If testValueUnderTrgt >= starterMaxSizeRF2 Then
                                                    groupCheck = False
                                                Else
                                                    groupCheck = False
                                                End If
                                            End If
                                            'Group Determination
                                            If savedValueDelta >= starterMaxSizeRF2 Then
                                                starterMaxDelta = starterMaxSizeRF2
                                            ElseIf savedValueDelta < starterMaxSizeRF2 Then
                                                starterMaxDelta = savedValueDelta
                                            End If
                                            
                                            If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 3).Value = "added" Then
                                                reFeedValueOne = Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value
                                                savedValueDelta = savedValueDelta - reFeedValueOne
                                                reFeedValueTwo = 0
                                                twoDLArray(1, 1) = Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 9).Value
                                                twoDLArray(1, 2) = Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 13).Value
                                                twoDLArray(1, 3) = Round(Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value / Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 13).Value, 0)
                                                twoDLArray(1, 4) = Round(Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value, 2)
                                                twoDLArray(1, 5) = Round(Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value / Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 13).Value, 0)
                                                twoDLArray(1, 6) = Round(Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value, 2)
                                            ElseIf Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 3).Value = "added1" Then
                                                reFeedValueOne = 0
                                                reFeedValueTwo = Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value
                                                savedValueDelta = savedValueDelta - reFeedValueTwo
                                                twoDLArray(1, 1) = "Refeed Or RoundingScrap"
                                                twoDLArray(1, 2) = 1
                                                twoDLArray(1, 3) = Round(Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value, 2)
                                                twoDLArray(1, 4) = Round(Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value, 2)
                                                twoDLArray(1, 5) = Round(Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value, 2)
                                                twoDLArray(1, 6) = Round(Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value, 2)
                                            End If
                                            
                                            '///// breakpoint check
                                            For iiii = iii + 1 To reFeedLastRowRF2
                                                'check added1 or added, 'check if savedValue is greater than max size or not, located here because this is the part where it will change, 4 different conditions

                                                If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 3).Value <> "FT" Then
                                                    If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 1).Value <> "" Then
                                                        If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 1).Value <= savedValue Then
                                                            If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 10).Value > (savedDateandTimeSecond + (usageTimeRF2 - 0.5) / 60 / 24) Then
                                                                Exit For
                                                            Else
                                                                If groupCheck = True Then
                                                                    If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 8).Value = Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 8).Value Then
                                                                        'Nothing
                                                                    Else
                                                                        Exit For
                                                                    End If
                                                                ElseIf groupCheck = False Then
                                                                    'Nothing
                                                                End If
                                                                addedOneTFCheck = True
                                                                If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 3).Value = "added1" Then
                                                                    If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 8).Value = Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 8).Value Then
                                                                        For xxy = 1 To iiii - 1
                                                                            If Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 8).Value = Sheets("preferment_calc").Cells(iiii - xxy, reFeedSourceCol + 8).Value Then
                                                                                If Sheets("preferment_calc").Cells(iiii - xxy, reFeedSourceCol + 7).Value > 0 Then
                                                                                    If (Sheets("preferment_calc").Cells(iiii - xxy, reFeedSourceCol + 10).Value + (starterFermentationTimeRF2 / 60 / 24)) <= Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 10).Value Then
                                                                                        If (Sheets("preferment_calc").Cells(iiii - xxy, reFeedSourceCol + 10).Value + (starterFermentationTimeRF2 + usageTimeRF2 / 60 / 24)) > Sheets("preferment_calc").Cells(iiii, reFeedSourceCol + 10).Value Then
                                                                                            addedOneTFCheck = True
                                                                                            Exit For
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        Next
                                                                    Else
                                                                        addedOneTFCheck = False
                                                                    End If
                                                                    
                                                                    If xxy = iiii Then
                                                                        addedOneTFCheck = False
                                                                    End If
                                                                End If
                                                                If addedOneTFCheck = True Then 'New check that is not based on 10 min
                                                                    returnArrayVar = myfuncSummarize(reFeedValueOne, reFeedValueTwo, starterMaxDelta, reFeedSourceCol, iiii, iii, savedValueDelta)
                                                                    reFeedValueOne = returnArrayVar(1)
                                                                    reFeedValueTwo = returnArrayVar(2)
                                                                    savedValueDelta = returnArrayVar(7)
                                                                    endForLoopCheck = returnArrayVar(6)
                                                                    varAdd = returnArrayVar(5)
                                                                    varDiffAdd = returnArrayVar(3)
                                                                    lineNameGroup = returnArrayVar(8)
                                                                    binSizeDet = returnArrayVar(9)
                                                                    'Beginning of array check
                                                                    If endForLoopCheck = False Then
                                                                        numberOfBins = Round(varAdd / binSizeDet, 0)
                                                                        qtyToBeAddedtoLarr = Round(varAdd, 2)
                                                                    Else
                                                                        numberOfBins = Round(varDiffAdd / binSizeDet, 0)
                                                                        qtyToBeAddedtoLarr = Round(varDiffAdd, 2)
                                                                    End If
                                                                    'ArrayCheck 'Line/Refeed+Rounding Remainder, Bin Size, Number of Bins, Total Bins, Total QTY
                                                                    checkMatchArr = False
                                                                    For yxy = 1 To 30
                                                                        If twoDLArray(yxy, 1) = lineNameGroup Then
                                                                            If twoDLArray(yxy, 2) = binSizeDet Then
                                                                                If twoDLArray(yxy, 3) = numberOfBins Then
                                                                                    If twoDLArray(yxy, 4) = qtyToBeAddedtoLarr Then
                                                                                        twoDLArray(yxy, 5) = twoDLArray(yxy, 5) + numberOfBins
                                                                                        twoDLArray(yxy, 6) = twoDLArray(yxy, 6) + qtyToBeAddedtoLarr
                                                                                        checkMatchArr = True
                                                                                        Exit For
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next
                                                                    If checkMatchArr = False Then
                                                                        twoDLArray(counterXXArray, 1) = lineNameGroup
                                                                        twoDLArray(counterXXArray, 2) = binSizeDet
                                                                        twoDLArray(counterXXArray, 3) = numberOfBins
                                                                        twoDLArray(counterXXArray, 4) = qtyToBeAddedtoLarr
                                                                        twoDLArray(counterXXArray, 5) = numberOfBins
                                                                        twoDLArray(counterXXArray, 6) = qtyToBeAddedtoLarr
                                                                        counterXXArray = counterXXArray + 1
                                                                    End If
                                                                    'End of Array Check
                                                                    If endForLoopCheck = True Then
                                                                        Exit For
                                                                    End If
                                                                Else
                                                                    GoTo nextiterationThree
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
nextiterationThree:
                                            Next
                                            If Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 3).Value <> "FT" Then
                                                For yyy = 1 To counterXXArray
                                                    Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 14).Value = Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 14).Value & twoDLArray(yyy, 1) & " | " & twoDLArray(yyy, 2) & " | " & twoDLArray(yyy, 3) & " | " & twoDLArray(yyy, 4) & " | " & twoDLArray(yyy, 5) & " | " & twoDLArray(yyy, 6) & Chr(10)
                                                Next
                                                Erase twoDLArray
                                                
                                            Else
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 14).Value = "NA"
                                            End If
                                            'Should print here after the for loop for the additional of refeed one and two has been completed
                                            If (reFeedValueOne + reFeedValueTwo) > 0 Then
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 1).Value = reFeedValueOne + reFeedValueTwo
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 5).Value = reFeedValueOne
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 6).Value = reFeedValueTwo
                                                Sheets("preferment_calc").Cells(iii, reFeedSourceCol + 7).Value = Round(reFeedValueTwo / starterRefeedPercentageRF2, 2)
                                                savedValue = savedValueDelta
                                                Sheets("preferment_calc").Cells(ii, reFeedSourceCol + 7).Value = savedValueDelta
                                                rowCountVerThree = iii + 1
                                                savedValue = savedValueDelta
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        
    Next
'

formatRefeedPart2 = True


FuncExitSecond:
    Exit Function
ErrHandleSecond:
    formatRefeedPart2 = False
    GoTo FuncExitSecond
End Function

Public Function myfuncSummarize(ByVal varFeedOne As Double, ByVal varFeedTwo As Double, ByVal starterMaxSizeRF23 As Double, ByVal reFeedSourceCol2 As Integer, ByVal rowNum As Integer, ByVal rowNum2 As Integer, ByVal savedValue As Double) As Variant

'rowNum = iiii:IncrementalRow, rowNum1 = iii:Initial Row
Dim varDifference, varRemainder As Double
Dim returnValueCheck As Boolean
Dim binSize As Double
Dim lineName As String
Dim qtypreDetermination As Double

qtypreDetermination = Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value

If (varFeedOne + varFeedTwo + Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value) < starterMaxSizeRF23 Then
    'Volume added depending on whether its a refeed or volume to be used on the dough
    If Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 3).Value = "added" Then
        varFeedOne = varFeedOne + Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value
        binSize = Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 13).Value
        lineName = Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 9).Value
    ElseIf Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 3).Value = "added1" Then
        varFeedTwo = varFeedTwo + Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value
        binSize = 1
        lineName = "Refeed Or RoundingScrap"
    End If
    'Establish New Timing Counter for Numbering Or split for Process Order
    If Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 4).Value = "NEW" Then
        Sheets("preferment_calc").Cells(rowNum2, reFeedSourceCol2 + 12).Value = Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 12).Value
        Sheets("preferment_calc").Cells(rowNum2, reFeedSourceCol2 + 4).Value = "NEW"
    End If
    'Cross Check with savedValue to update it
    savedValue = Round(savedValue - Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value, 2)
    Sheets("preferment_calc").Cells(rowNum2, reFeedSourceCol2 + 2).Value = Sheets("preferment_calc").Cells(rowNum2, reFeedSourceCol2 + 2).Value & " | " & Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 2).Value
    'Clear Range
    Sheets("preferment_calc").Range("R" & rowNum & ":AH" & rowNum).ClearContents 'Manual change of Range
    varDifference = 0
    varRemainder = 0
    returnValueCheck = False
ElseIf (varFeedOne + varFeedTwo + Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value) = starterMaxSizeRF23 Then
    'Volume added depending on whether its a refeed or volume to be used on the dough
    If Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 3).Value = "added" Then
        varFeedOne = varFeedOne + Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value
        binSize = Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 13).Value
        lineName = Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 9).Value
    ElseIf Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 3).Value = "added1" Then
        varFeedTwo = varFeedTwo + Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value
        binSize = 1
        lineName = "Refeed Or RoundingScrap"
    End If
    'Establish New Timing Counter for Numbering Or split for Process Order
    If Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 4).Value = "NEW" Then
        Sheets("preferment_calc").Cells(rowNum2, reFeedSourceCol2 + 12).Value = Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 12).Value
        Sheets("preferment_calc").Cells(rowNum2, reFeedSourceCol2 + 4).Value = "NEW"
    End If
    'Cross Check with savedValue to update it
    savedValue = Round(savedValue - Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value, 2)
    Sheets("preferment_calc").Cells(rowNum2, reFeedSourceCol2 + 2).Value = Sheets("preferment_calc").Cells(rowNum2, reFeedSourceCol2 + 2).Value & " | " & Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 2).Value
    'Clear Range
    Sheets("preferment_calc").Range("R" & rowNum & ":AH" & rowNum).ClearContents 'Manual change of Range
    varDifference = 0
    varRemainder = 0
    returnValueCheck = True
ElseIf (varFeedOne + varFeedTwo + Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value) > starterMaxSizeRF23 Then
    If Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 3).Value = "added" Then
        varDifference = starterMaxSizeRF23 - varFeedOne - varFeedTwo
        varRemainder = Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value - varDifference
        varFeedOne = starterMaxSizeRF23 - varFeedTwo
        Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value = varRemainder
        Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 5).Value = varRemainder
        binSize = Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 13).Value
        lineName = Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 9).Value
    ElseIf Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 3).Value = "added1" Then
        varDifference = starterMaxSizeRF23 - varFeedOne - varFeedTwo
        varRemainder = Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value - varDifference
        varFeedTwo = starterMaxSizeRF23 - varFeedOne
        Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 1).Value = varRemainder
        Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 6).Value = varRemainder
        binSize = 1
        lineName = "Refeed Or RoundingScrap"
    End If
    'Cross Check with savedValue to update it
    savedValue = Round(savedValue - varDifference, 2)
    Sheets("preferment_calc").Cells(rowNum2, reFeedSourceCol2 + 2).Value = Sheets("preferment_calc").Cells(rowNum2, reFeedSourceCol2 + 2).Value & " | " & Sheets("preferment_calc").Cells(rowNum, reFeedSourceCol2 + 2).Value
    returnValueCheck = True
End If

Dim vArr(1 To 9) As Variant
vArr(1) = varFeedOne
vArr(2) = varFeedTwo
vArr(3) = varDifference ' Value needed to determine in split
vArr(4) = varRemainder
vArr(5) = qtypreDetermination  'Value to determine amount
vArr(6) = returnValueCheck 'if true exit for
vArr(7) = savedValue
vArr(8) = lineName ' Line name
vArr(9) = binSize 'Bin Size saved

myfuncSummarize = vArr

End Function
'another part of the calculation

Public Function preDetermination(ByVal columnRef As Integer, ByVal starerMxSze As Double, ByVal starterPercentage As Double, ByVal lastRowDet As Integer, ByVal usageTime As Double) As Boolean

        Dim svdValueDiffNum, rndUpNum As Double
        Dim xxCheckAdd As Boolean
        Dim xxContSvd As Integer
        Dim counterYNCheck As Integer
        Dim runningSum As Double
        counterYNCheck = 0

        Dim testRowSeqVerTwo As Integer
        Dim firstAddressRow, finalAddressRow, compareAddressRowOne, compareAddressRowTwo As Integer
        Dim searchlast As Range
        Dim search As Range
        Dim countDiffGroup As Integer
        Dim beginAddress, endingAddress As String
        Dim trgtValueGroup As Double
        Dim backwardCheckValue As Double
        Dim backwardCheckValueTF As Boolean
        Dim dblCheckStr As Boolean
        Dim adjFacValue As Double
        Dim rowDeterm As Integer
        Dim compareTimeOne, compareTimeTwo As Double
        Dim dblCheckStr2 As Boolean
        With Sheets("preferment_calc")
            dblCheckStr = True
            adjFacValue = 0
            svdValueDiffNum = 0
            rowDeterm = 2
            dblCheckStr2 = True
            Do While (dblCheckStr = True)
            
                For ii = rowDeterm To lastRowDet
                    If .Cells(ii, columnRef + 3).Value = "added1" Then
                        If .Cells(ii + 1, columnRef + 3).Value = "added1" Then
                            .Cells(ii, columnRef + 15).Value = "NO"
                        End If
                        If .Cells(ii - 1, columnRef + 3).Value = "added1" Then
                            .Cells(ii, columnRef + 15).Value = "NO"
                        End If
                    End If
                Next
                For ii = rowDeterm To lastRowDet
                    If .Cells(ii, columnRef + 15).Value = "NO" Then
                        For iiii = 1 To ii - 1
                            If .Cells(ii - iiii, columnRef + 3).Value <> "added" Then
                                If .Cells(ii - iiii, columnRef + 8).Value = .Cells(ii, columnRef + 8).Value Then
                                    .Cells(ii - iiii, columnRef + 15).Value = "NO"
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                Next
                
                For ii = rowDeterm To lastRowDet
                    backwardCheckValueTF = False
                    backwardCheckValue = 0
                    If .Cells(ii, columnRef + 3).Value <> "added" Then
                        trgtValueGroup = 0
                        If .Cells(ii, columnRef + 14).Value = "" Then
                            .Cells(ii, columnRef + 14) = .Cells(ii, columnRef + 7).Value
                        End If
                        testRowSeqVerTwo = ii
                        For iii = ii To lastRowDet
                            If .Cells(iii + 1, columnRef + 8).Value = .Cells(testRowSeqVerTwo, columnRef + 8).Value Then
                                testRowSeqVerTwo = iii + 1
                            Else
                                Exit For
                            End If
                        Next
                        
                        If .Cells(ii, columnRef + 3) = "added1" Then
                            For iiii = 1 To ii - 1
                                If .Cells(ii - iiii, columnRef + 3).Value <> "added" Then
                                    If .Cells(ii - iiii, columnRef + 8).Value = .Cells(ii, columnRef + 8).Value Then
                                        backwardCheckValue = .Cells(ii - iiii, columnRef + 16).Value
                                        backwardCheckValueTF = True
                                        Exit For
                                    End If
                                End If
                            Next
                            If iiii > (ii - 1) Then
                                backwardCheckValueTF = False
                            End If
                        End If
                        
                        Set search = ActiveSheet.Range("Z" & iii + 1 & ":Z14000")
                        Set searchlast = search.Cells(search.Cells.Count)
                        Set rngFindValue = ActiveSheet.Range("Z" & iii + 1 & ":Z14000").Find(.Cells(iii, columnRef + 8).Value, searchlast, xlValues)
                        countDiffGroup = 1
                        If Not rngFindValue Is Nothing Then
                            firstAddressRow = rngFindValue.Row
                            finalAddressRow = firstAddressRow
                            compareAddressRowOne = firstAddressRow
                            Do While (countDiffGroup = 1)
                                Set rngFindValue = search.FindNext(rngFindValue)
                                compareAddressRowTwo = rngFindValue.Row
                                countDiffGroup = compareAddressRowTwo - compareAddressRowOne
                                If countDiffGroup = 1 Then
                                    compareAddressRowOne = compareAddressRowTwo
                                    finalAddressRow = compareAddressRowTwo
                                End If
                            Loop
                            trgtValueGroup = Application.WorksheetFunction.Sum(.Range("S" & firstAddressRow & ":S" & finalAddressRow))
                            Set rngFindValue = Nothing
                            If trgtValueGroup < starerMxSze Then
                                GoTo nxxtItterationss
                            End If
                        Else
                            trgtValueGroup = 0
                            Set rngFindValue = Nothing
                            GoTo nxxtItterationss
                        End If
                        
                        rndUpNum = Application.WorksheetFunction.RoundUp((trgtValueGroup - svdValueDiffNum) / starerMxSze, 0) * starerMxSze
                        svdValueDiffNum = rndUpNum - .Cells(ii, columnRef + 14).Value

                                             
                        For iii = ii + 1 To lastRowDet
                            If .Cells(iii, columnRef + 3).Value <> "added" Then
                                If .Cells(iii, columnRef + 7).Value >= svdValueDiffNum Then
                                    compareTimeOne = CDbl(.Cells(iii, columnRef + 10).Value)
                                    compareTimeTwo = CDbl(.Cells(ii, columnRef + 10).Value) + ((usageTime + 5) / 60 / 24)
                                    If compareTimeOne <= compareTimeTwo Then
                                        If .Cells(ii, columnRef + 15).Value <> "NO" Then
                                            If (Asc(.Cells(iii, columnRef + 8).Text) - Asc(.Cells(ii, columnRef + 8).Text)) = 1 Or (Asc(.Cells(iii, columnRef + 8).Text) - Asc(.Cells(ii, columnRef + 8).Text)) = -1 Then
                                                If backwardCheckValueTF = True Then
                                                    If backwardCheckValue >= (rndUpNum * starterPercentage) Then
                                                        .Cells(ii, columnRef + 15).Value = "YES"
                                                        .Cells(ii, columnRef + 16).Value = rndUpNum
                                                        .Cells(ii, columnRef + 17).Value = Round(((.Cells(ii, columnRef + 16).Value * starterPercentage) - .Cells(ii, columnRef + 6).Value), 2)
                                                        .Cells(iii, columnRef + 14).Value = Round((.Cells(iii, columnRef + 7).Value - svdValueDiffNum), 2)
                                                        Exit For
                                                    Else
                                                        GoTo nxxtItterationss
                                                    End If
                                                ElseIf backwardCheckValueTF = False Then
                                                    .Cells(ii, columnRef + 15).Value = "YES"
                                                    .Cells(ii, columnRef + 16).Value = rndUpNum
                                                    .Cells(ii, columnRef + 17).Value = Round(((.Cells(ii, columnRef + 16).Value * starterPercentage) - .Cells(ii, columnRef + 6).Value), 2)
                                                    .Cells(iii, columnRef + 14).Value = Round((.Cells(iii, columnRef + 7).Value - svdValueDiffNum), 2)
                                                    Exit For
                                                End If
                                            Else
                                                GoTo nxxtItterationss
                                            End If
                                        Else
                                            GoTo nxxtItterationss
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        If iii > lastRowDet Then
nxxtItterationss:
                            .Cells(ii, columnRef + 15).Value = "NO"
                            .Cells(ii, columnRef + 16).Value = .Cells(ii, columnRef + 14).Value
                            .Cells(ii, columnRef + 17).Value = Round(((.Cells(ii, columnRef + 16).Value * starterPercentage) - .Cells(ii, columnRef + 6).Value), 2)
                            svdValueDiffNum = 0
                        End If
                    End If
                Next
                dblCheckStr = False
                dblCheckStr2 = False
                svdValueDiffNum = 0
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                For ii = rowDeterm To lastRowDet

                    If .Cells(ii, columnRef + 3).Value <> "added" Then
                        trgtValueGroup = 0
                        testRowSeqVerTwo = ii
                        For iii = ii To lastRowDet
                            If .Cells(iii + 1, columnRef + 8).Value = .Cells(testRowSeqVerTwo, columnRef + 8).Value Then
                                testRowSeqVerTwo = iii + 1
                            Else
                                Exit For
                            End If
                        Next

                        Set search = ActiveSheet.Range("Z" & iii + 1 & ":Z14000")
                        Set searchlast = search.Cells(search.Cells.Count)
                        Set rngFindValue = ActiveSheet.Range("Z" & iii + 1 & ":Z14000").Find(.Cells(iii, columnRef + 8).Value, searchlast, xlValues)
                        countDiffGroup = 1
                        If Not rngFindValue Is Nothing Then
                            firstAddressRow = rngFindValue.Row
                            finalAddressRow = firstAddressRow
                            compareAddressRowOne = firstAddressRow
                            Do While (countDiffGroup = 1)
                                Set rngFindValue = search.FindNext(rngFindValue)
                                compareAddressRowTwo = rngFindValue.Row
                                countDiffGroup = compareAddressRowTwo - compareAddressRowOne
                                If countDiffGroup = 1 Then
                                    compareAddressRowOne = compareAddressRowTwo
                                    finalAddressRow = compareAddressRowTwo
                                End If
                            Loop
                            adjFacValue = Application.WorksheetFunction.Sum(.Range("AI" & ii & ":AI" & finalAddressRow)) 'ii / iii+1 / firstaddressRow, could be require testing
                            Set rngFindValue = Nothing
                        Else
                            adjFacValue = 0
                            Set rngFindValue = Nothing
                            GoTo nxxtItterationss
                        End If
                        If .Cells(ii, columnRef + 15).Value = "YES" Then
                            If (.Cells(ii, columnRef + 14).Value + adjFacValue) > .Cells(ii, columnRef + 16).Value Then 'stop here for future testing
                                For iiii = ii + 1 To lastRowDet
                                    If .Cells(iiii, columnRef + 3).Value <> "added" Then
                                        If (.Cells(ii, columnRef + 16).Value + starerMxSze - .Cells(ii, columnRef + 14).Value) < .Cells(iiii, columnRef + 7).Value Then
                                            .Cells(ii, columnRef + 16).Value = .Cells(ii, columnRef + 16).Value + starerMxSze
                                            svdValueDiffNum = .Cells(ii, columnRef + 16).Value - .Cells(ii, columnRef + 14).Value
                                            .Range("AF" & ii + 1 & ":AI14000").ClearContents
                                            .Cells(iiii, columnRef + 14).Value = .Cells(iiii, columnRef + 7) - svdValueDiffNum
                                            dblCheckStr = True
                                            dblCheckStr2 = True
                                            Exit For
                                        Else
                                           .Cells(ii, columnRef + 16).Value = .Cells(ii, columnRef + 14).Value
                                           .Cells(ii, columnRef + 15).Value = "NO"
                                           .Range("AF" & ii + 1 & ":AI14000").ClearContents
                                           svdValueDiffNum = 0
                                           dblCheckStr = True
                                           dblCheckStr2 = True
                                           Exit For
                                        End If
                                    End If
                                Next
                                rowDeterm = ii + 1
                             ElseIf (.Cells(ii, columnRef + 14).Value + adjFacValue) < (.Cells(ii, columnRef + 16).Value - starerMxSze) Then
                                If (.Cells(ii, columnRef + 14).Value + adjFacValue) < .Cells(ii, columnRef + 14).Value Then 'stop here for future testing
                                    .Cells(ii, columnRef + 14).Value = .Cells(ii, columnRef + 14).Value + adjFacValue
                                    .Cells(ii, columnRef + 15).Value = "NO"
                                    .Cells(ii, columnRef + 16).Value = .Cells(ii, columnRef + 14).Value
                                    .Range("AF" & ii + 1 & ":AI14000").ClearContents
                                    rowDeterm = ii
                                    svdValueDiffNum = 0
                                    dblCheckStr = True
                                    dblCheckStr2 = True
                                    Exit For
                                End If
                             End If
                         End If
                    End If
                    If dblCheckStr2 = True Then
                        Exit For
                    End If
                Next
            Loop
'Check Part
        'Round all the preset if it is YES
            For ii = 2 To lastRowDet
                If .Cells(ii, columnRef + 3).Value <> "added" Then
                    If .Cells(ii, columnRef + 16).Value > 0 Then
                        .Cells(ii, columnRef + 7).Value = Application.WorksheetFunction.RoundUp(.Cells(ii, columnRef + 16).Value, 2)
                        .Cells(ii, columnRef + 6).Value = Application.WorksheetFunction.RoundUp(.Cells(ii, columnRef + 7).Value * starterPercentage, 2)
                        .Cells(ii, columnRef + 1).Value = .Cells(ii, columnRef + 6).Value
                        .Cells(ii, columnRef + 7).Value = Round(.Cells(ii, columnRef + 1).Value / starterPercentage, 2)
                    End If
                    If .Cells(ii, columnRef + 1).Value < 40 Then
                        .Cells(ii, columnRef + 6).Value = Application.WorksheetFunction.RoundUp(.Cells(ii, columnRef + 7).Value * starterPercentage, 1)
                        .Cells(ii, columnRef + 1).Value = .Cells(ii, columnRef + 6).Value
                        .Cells(ii, columnRef + 7).Value = Round(.Cells(ii, columnRef + 1).Value / starterPercentage, 1)
                    End If
                End If
                
                    
            Next
            
        End With

preDetermination = True

End Function


