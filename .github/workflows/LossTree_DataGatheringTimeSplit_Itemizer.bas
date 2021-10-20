Sub splitOperationalDowntimebyShiftAndValidate()

'Shift1 is determined by material Number
'Normal Shift for L03 is 6:00, 2:00pm,10:00pm
'Selected Shift startsfrom 4:00, 12, 8...

Dim stdShftOne, stdShftTwo, stdShftThree As Double
Dim erlyShftOne, erlyShftTwo, erlyShftThree As Double
Dim i, ii, lastRow, rowAddcount, startTimeSetup, storedTime, addDay As Double
Dim startTimeSplit As Double
Dim shftStartDet As Integer
Dim differDet, ordRunLength As Long
'//////////////////////////////////
Dim wsMain As Worksheet
Set wsMain = Sheets("LOGICTREE EXTRACTS INSTUCTIONS")
Dim xRng As Range
Set xRng = wsMain.Range("R:R")
'///////////////////////////
Dim x As Variant
Dim matNum As String
Dim t As Boolean
Dim shftsSeperate(1 To 7) As String
Dim counter2 As Long
'////////// For highlights
Dim checkValue, checkValue2, checkTimeError As Double
Sheets("SUMMARIZETABLE").Activate

shftsSeperate(1) = "SHIFT 1"
shftsSeperate(2) = "SHIFT 2"
shftsSeperate(3) = "SHIFT 3"
shftsSeperate(4) = "SHIFT 1"
shftsSeperate(5) = "SHIFT 2"
shftsSeperate(6) = "SHIFT 3"
shftsSeperate(7) = "SHIFT 1"

stdShftOne = 0.25 ' Starts at 6:00 am
stdShftTwo = 0.58333333
stdShftThree = 0.91666667

erlyShftOne = 0.16666667 'Starts at 4:00am
erlyShftTwo = 0.5
erlyShftThree = 0.83333333

lastRow = Sheets("SUMMARIZETABLE").Cells(Rows.Count, 1).End(xlUp).Row
i = 2
Do While i < lastRow + 1
    rowAddcount = 0
 '   matNum = Sheets("SUMMARIZETABLE").Cells(i, 2).Value
 '   x = Application.VLookup(matNum, xRng, 1, False)
        'Check if it is the selected material number....
 '       If IsError(x) = False Then
        '4:00 Run

'        Else
        '6:00 Run
'///////////////////COPYRUN
            ordRunLength = Sheets("SUMMARIZETABLE").Cells(i, 9).Value + _
                Sheets("SUMMARIZETABLE").Cells(i, 10).Value
            startTimeSetup = Round(Sheets("SUMMARIZETABLE").Cells(i, 6).Value, 8)
            Select Case startTimeSetup
                Case 0.25 To 0.583333329
                    shftStartDet = 1
                    differDet = stdShftTwo - startTimeSetup
                Case 0.58333333 To 0.916666669
                    shftStartDet = 2
                    differDet = stdShftThree - startTimeSetup
                Case 0.91666667 To 1.25
                    shftStartDet = 3
                    differDet = 1.25 - startTimeSetup
                Case 0 To 0.24999999
                    shftStartDet = 4
                    differDet = stdShftOne - startTimeSetu
            End Select
            
            rowAddcount = (Application.RoundDown(ordRunLength / 24 / 60, 8) - differDet) / 0.33333333
            If rowAddcount >= 0.00069444 Then
                rowAddcount = Application.RoundUp(rowAddcount, 0)
            Else
                rowAddcount = 0
            End If
            'Create Rows and Assign Shifts
            If rowAddcount = 0 Then
                If shftStartDet = 4 Then
                    Sheets("SUMMARIZETABLE").Cells(i, 12).Value = shftsSeperate(3)
                Else
                    Sheets("SUMMARIZETABLE").Cells(i, 12).Value = shftsSeperate(shftStartDet)
                End If
            Else
            'Insert Row
                For ii = 1 To rowAddcount
                    Sheets("SUMMARIZETABLE").Activate
                    Rows(i).Select
                    Selection.Offset(1, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrAbove
                    With Sheets("SUMMARIZETABLE")
                        .Cells(i + 1, 1).Value = .Cells(i, 1).Value
                        .Cells(i + 1, 2).Value = .Cells(i, 2).Value
                        .Cells(i + 1, 13).Value = .Cells(i, 13).Value
                        .Cells(i + 1, 14).Value = .Cells(i, 14).Value
                        .Cells(i + 1, 15).Value = .Cells(i, 15).Value
                        .Cells(i + 1, 16).Value = .Cells(i, 16).Value
                        .Cells(i + 1, 17).Value = .Cells(i, 17).Value
                        .Cells(i + 1, 18).Value = .Cells(i, 18).Value
                        lastRow = lastRow + 1
                    End With
                Next
            'Insert Time difference and case seperation
                storedTime = Round(ordRunLength / 24 / 60, 8)
                
                For ii = 0 To rowAddcount
                    startTimeSplit = Round(Sheets("SUMMARIZETABLE").Cells(i + ii, 6).Value, 8)
                    Sheets("SUMMARIZETABLE").Cells(i + ii, 1).Value = Sheets("SUMMARIZETABLE").Cells(i, 1).Value
                    Sheets("SUMMARIZETABLE").Cells(i + ii, 2).Value = Sheets("SUMMARIZETABLE").Cells(i, 2).Value
                    Select Case startTimeSplit
                        Case 0.25 To 0.583333329
                            If storedTime >= 0.33333333 Then
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = 0.58333333
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 6).Value = 0.58333333
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 5).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                            ElseIf (startTimeSplit + storedTime) > 0.58333333 Then
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = 0.58333333
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 6).Value = 0.58333333
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 5).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                            Else
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 6).Value + _
                                    storedTime
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                            End If
                            Sheets("SUMMARIZETABLE").Cells(i + ii, 12).Value = "SHIFT 1"
                            storedTime = storedTime - (stdShftTwo - startTimeSplit)
                        Case 0.58333333 To 0.916666669
                            If storedTime >= 0.33333333 Then
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = 0.91666667
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 6).Value = 0.91666667
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 5).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                            ElseIf (startTimeSplit + storedTime) > 0.91666667 Then
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = 0.91666667
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 6).Value = 0.91666667
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 5).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                            Else
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 6).Value + _
                                    storedTime
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                            End If
                            Sheets("SUMMARIZETABLE").Cells(i + ii, 12).Value = "SHIFT 2"
                            storedTime = storedTime - (stdShftThree - startTimeSplit)
                        Case 0.91666667 To 1.25
                            If storedTime >= 0.33333333 Then
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = 0.25
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 6).Value = 0.25
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value + 1
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 5).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value + 1
                            ElseIf (startTimeSplit + storedTime - 1) > 0.25 Then
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = 0.25
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 6).Value = 0.25
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value + 1
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 5).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value + 1
                            Else
                                If storedTime >= (1 - 0.91666667) Then
                                    Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value + 1
                                    Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = storedTime - (1 - 0.91666667)
                                ElseIf storedTime < (1 - 0.9166667) Then
                                    Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                                    Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 6).Value + _
                                        storedTime
                                End If
                            End If
                            Sheets("SUMMARIZETABLE").Cells(i + ii, 12).Value = "SHIFT 3"
                            storedTime = storedTime - (1.25 - startTimeSplit)
                        Case 0 To 0.249999999
                            If storedTime >= 0.33333333 Then
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = 0.25
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 6).Value = 0.25
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 5).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                            ElseIf (startTimeSplit + storedTime) > 0.25 Then
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = 0.25
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 6).Value = 0.25
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                                Sheets("SUMMARIZETABLE").Cells(i + ii + 1, 5).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                            Else
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 8).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 6).Value + _
                                    storedTime
                                Sheets("SUMMARIZETABLE").Cells(i + ii, 7).Value = Sheets("SUMMARIZETABLE").Cells(i + ii, 5).Value
                            End If
                            Sheets("SUMMARIZETABLE").Cells(i + ii, 12).Value = "SHIFT 3"
                            storedTime = storedTime - (stdShftOne - startTimeSplit)
                    End Select
                    
                Next
            End If
            i = i + rowAddcount
 '////////////////////COPY RANGE
'        End If
i = i + 1
Loop

For i = 2 To lastRow
    With Sheets("SUMMARIZETABLE")
        .Cells(i, 9).FormulaR1C1 = "=(RC[-1]+RC[-2]-RC[-3]-RC[-4])*24*60"
    End With
Next

'Validation of the spliting time to make sure it make sense, if it does not it will highlight it blue
'Check the difference between the End Time of Previous Row is not Greater than Current Row Start Time
'If it is a downtime issue, Repost the confirmation in SAP
'Check the difference, any misalignment will be highlighted

lastRow = Sheets("SUMMARIZETABLE").Cells(Rows.Count, 1).End(xlUp).Row

'Highlight problem order and ask user to adjust manaully
    For i = 3 To lastRow
    
        checkTimeError = Sheets("SUMMARIZETABLE").Cells(i, 5).Value + Sheets("SUMMARIZETABLE").Cells(i, 6).Value - _
            Sheets("SUMMARIZETABLE").Cells(i, 7).Value - Sheets("SUMMARIZETABLE").Cells(i, 8).Value
        If Round(checkTimeError * 24 * 60 * -1, 0) <> Round(Sheets("SUMMARIZETABLE").Cells(i, 9).Value + _
            Sheets("SUMMARIZETABLE").Cells(i, 10).Value, 0) Then
            If checkTimeError <= -0.0006944 Then
                Sheets("SUMMARIZETABLE").Activate
                Rows(i).Interior.ColorIndex = 3
                GoTo highlghterrMoveNxtRow2
            End If
        End If
        
        checkValue = Sheets("SUMMARIZETABLE").Cells(i, 5).Value + Sheets("SUMMARIZETABLE").Cells(i, 6).Value - _
            Sheets("SUMMARIZETABLE").Cells(i, 11).Value / 60 / 24 - _
            Sheets("SUMMARIZETABLE").Cells(i - 1, 7).Value - Sheets("SUMMARIZETABLE").Cells(i - 1, 8).Value
        If checkValue <= -0.00069444 Then
            MsgBox "There is a diffence in time that does not make sense." & Chr(10) _
                & Round(checkValue * 60 * 24, 0) & " " & "Min(s) need to be allocated to Current Run Time / Previous Run Time / Current ChangeOver" & Chr(10) _
                & "AFFECTED ROWS HAS BEEN HIGHLIGHTED IN BLUE"
            Sheets("SUMMARIZETABLE").Activate
            Rows(i).Interior.ColorIndex = 37
            If Rows(i - 1).Interior.ColorIndex <> 3 Then
                Sheets("SUMMARIZETABLE").Activate
                Rows(i - 1).Interior.ColorIndex = 37
            End If
        End If
highlghterrMoveNxtRow2:

    Next
    
For i = 2 To lastRow

    With Sheets("SUMMARIZETABLE")
        If Round(.Cells(i, 9).Value, 0) = 0 Then
            .Rows(i).EntireRow.Delete
        End If
    End With

Next

End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'AddRuntimeinbetweenoperatialdowntime

Sub DOWNTIMEITEMIZERS()
    Dim proOrderCount, i, ii, iii, counterCheck As Integer
    Dim rowNum As Long
    Dim proOrder As String
    Dim lastRow, lastRow2, lastRowErrror As Double
    Dim itemStartDate, itemStartTime, itemEndDate, itemEndTime As Double
    Dim rowFind As Long
    Dim itemCombineStart, itemCombineEnd, combineStart, combineEnd As Double
    Dim startDiff, endDiff As Double
    Dim dwnTime As Double
    Dim checkTableTime As Boolean
    
If Sheets("ITEMIZERS").Range("B3").Value = "" Then
    GoTo incorrectRow
End If

lastRow2 = Sheets("ITEMIZERS").Cells(Rows.Count, 2).End(xlUp).Row
lastRow = Sheets("SUMMARIZETABLE").Cells(Rows.Count, 1).End(xlUp).Row
If IsNumeric(Sheets("ITEMIZERS").Range("B3").Value) = False Then
    GoTo incorrectRow
End If

'Risk any error go to end
On Error GoTo incorrectRow
'Delcare for errors
lastRowErrror = Sheets("SUMMARIZETABLE").Cells(Rows.Count, 5).End(xlUp).Row + 1
rowFind = lastRowErrror
ii = 0
'////////////////////////////////////////////////////
proOrder = Sheets("ITEMIZERS").Range("B3").Value
rowFind = Sheets("SUMMARIZETABLE").Columns(1).Find(What:=proOrder, LookIn:=xlValues).Row

For i = 6 To lastRow2
    counterCheck = 0
    proOrderCount = Application.WorksheetFunction.CountIf(Sheets("SUMMARIZETABLE").Range("A:A"), proOrder)
    itemStartDate = Sheets("ITEMIZERS").Cells(i, 7).Value
    itemStartTime = Sheets("ITEMIZERS").Cells(i, 8).Value
    itemEndDate = Sheets("ITEMIZERS").Cells(i, 9).Value
    itemEndTime = Sheets("ITEMIZERS").Cells(i, 10).Value
    itemCombineStart = itemStartDate + itemStartTime
    itemCombineEnd = itemEndDate + itemEndTime
    dwnTime = itemCombineEnd - itemCombineStart
    ii = 0
    Do While ii <= proOrderCount - 1
        combineStart = Sheets("SUMMARIZETABLE").Cells(rowFind + ii, 6).Value + Sheets("SUMMARIZETABLE").Cells(rowFind + ii, 5).Value
        combineEnd = Sheets("SUMMARIZETABLE").Cells(rowFind + ii, 7).Value + Sheets("SUMMARIZETABLE").Cells(rowFind + ii, 8).Value
        'if smaller than start time Go to Next Loop
        If itemCombineStart < combineStart Then
            counterCheck = counterCheck + 1
        ElseIf itemCombineStart > 73051 Then 'new added in 07-24-2020
            counterCheck = counterCheck + 1
        End If
        'If between time in single row, then insert below two columns-> subtract from the same line
        If itemCombineStart > combineStart Then
            startDiff = itemCombineStart - combineStart
            If itemCombineStart < combineEnd Then
                If itemCombineEnd < combineEnd Then
                    With Sheets("SUMMARIZETABLE")
                    
                        .Activate
                        Rows(rowFind + ii).Select
                        Selection.Offset(1, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrAbove
                        Selection.Offset(1, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrAbove
                        'Update Mid Time
                        .Cells(rowFind + ii + 1, 1).Value = .Cells(rowFind + ii, 1).Value
                        .Cells(rowFind + ii + 1, 2).Value = .Cells(rowFind + ii, 2).Value
                        .Cells(rowFind + ii + 1, 5).Value = itemStartDate
                        .Cells(rowFind + ii + 1, 6).Value = itemStartTime
                        .Cells(rowFind + ii + 1, 7).Value = itemEndDate
                        .Cells(rowFind + ii + 1, 8).Value = itemEndTime
                        .Cells(rowFind + ii + 1, 9).Value = Sheets("ITEMIZERS").Cells(i, 11).Value
                        .Cells(rowFind + ii + 1, 13).Value = "UNPLANNED DOWNTIME"
                        .Cells(rowFind + ii + 1, 14).Value = Sheets("ITEMIZERS").Cells(i, 2).Value
                        .Cells(rowFind + ii + 1, 15).Value = Sheets("ITEMIZERS").Cells(i, 4).Value
                        .Cells(rowFind + ii + 1, 16).Value = Sheets("ITEMIZERS").Cells(i, 5).Value
                        .Cells(rowFind + ii + 1, 17).Value = Sheets("ITEMIZERS").Cells(i, 6).Value
                        .Cells(rowFind + ii + 1, 18).Value = Sheets("ITEMIZERS").Cells(i, 3).Value & " " & Sheets("ITEMIZERS").Cells(i, 12).Value
                        'Update End Time
                        .Cells(rowFind + ii + 2, 1).Value = .Cells(rowFind + ii, 1).Value
                        .Cells(rowFind + ii + 2, 2).Value = .Cells(rowFind + ii, 2).Value
                        .Cells(rowFind + ii + 2, 5).Value = itemEndDate
                        .Cells(rowFind + ii + 2, 6).Value = itemEndTime
                        .Cells(rowFind + ii + 2, 7).Value = .Cells(rowFind + ii, 7).Value
                        .Cells(rowFind + ii + 2, 8).Value = .Cells(rowFind + ii, 8).Value
                        'Update Start Time
                        .Cells(rowFind + ii, 7).Value = itemStartDate
                        .Cells(rowFind + ii, 8).Value = itemStartTime
                      
                    End With
                    proOrderCount = proOrderCount + 2
                    
                ElseIf Round(itemCombineEnd, 8) = Round(combineEnd, 8) Then
                'The same as ending
                    With Sheets("SUMMARIZETABLE")
                        .Activate
                        Rows(rowFind + ii).Select
                        Selection.Offset(1, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrAbove
                        .Cells(rowFind + ii + 1, 1).Value = .Cells(rowFind + ii, 1).Value
                        .Cells(rowFind + ii + 1, 2).Value = .Cells(rowFind + ii, 2).Value
                        .Cells(rowFind + ii + 1, 5).Value = itemStartDate
                        .Cells(rowFind + ii + 1, 6).Value = itemStartTime
                        .Cells(rowFind + ii + 1, 7).Value = itemEndDate
                        .Cells(rowFind + ii + 1, 8).Value = itemEndTime
                        .Cells(rowFind + ii + 1, 9).Value = Sheets("ITEMIZERS").Cells(i, 11).Value
                        .Cells(rowFind + ii + 1, 18).Value = Sheets("ITEMIZERS").Cells(i, 3).Value & " " & Sheets("ITEMIZERS").Cells(i, 12).Value
                        .Cells(rowFind + ii + 1, 13).Value = "UNPLANNED DOWNTIME"
                        .Cells(rowFind + ii + 1, 14).Value = Sheets("ITEMIZERS").Cells(i, 2).Value
                        .Cells(rowFind + ii + 1, 15).Value = Sheets("ITEMIZERS").Cells(i, 4).Value
                        .Cells(rowFind + ii + 1, 16).Value = Sheets("ITEMIZERS").Cells(i, 5).Value
                        .Cells(rowFind + ii + 1, 17).Value = Sheets("ITEMIZERS").Cells(i, 6).Value
                        'Update Current Row Start Time
                        .Cells(rowFind + ii, 7).Value = itemStartDate
                        .Cells(rowFind + ii, 8).Value = itemStartTime
                    End With
                    
                    proOrderCount = proOrderCount + 1
                    
                ElseIf itemCombineEnd > combineEnd Then
                'Greater than the ending
                    With Sheets("SUMMARIZETABLE")
                        .Activate
                        Rows(rowFind + ii).Select
                        Selection.Offset(1, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrAbove
                        .Cells(rowFind + ii + 1, 1).Value = .Cells(rowFind + ii, 1).Value
                        .Cells(rowFind + ii + 1, 2).Value = .Cells(rowFind + ii, 2).Value
                        .Cells(rowFind + ii + 1, 5).Value = itemStartDate
                        .Cells(rowFind + ii + 1, 6).Value = itemStartTime
                        .Cells(rowFind + ii + 1, 7).Value = itemEndDate
                        .Cells(rowFind + ii + 1, 8).Value = itemEndTime
                        .Cells(rowFind + ii + 1, 9).Value = Sheets("ITEMIZERS").Cells(i, 11).Value
                        .Cells(rowFind + ii + 1, 13).Value = "UNPLANNED DOWNTIME"
                        .Cells(rowFind + ii + 1, 14).Value = Sheets("ITEMIZERS").Cells(i, 2).Value
                        .Cells(rowFind + ii + 1, 15).Value = Sheets("ITEMIZERS").Cells(i, 4).Value
                        .Cells(rowFind + ii + 1, 16).Value = Sheets("ITEMIZERS").Cells(i, 5).Value
                        .Cells(rowFind + ii + 1, 17).Value = Sheets("ITEMIZERS").Cells(i, 6).Value
                        .Cells(rowFind + ii + 1, 18).Value = Sheets("ITEMIZERS").Cells(i, 3).Value & " " & Sheets("ITEMIZERS").Cells(i, 12).Value
                        'Update Start Time
                        .Cells(rowFind + ii, 7).Value = itemStartDate
                        .Cells(rowFind + ii, 8).Value = itemStartTime
                        'Update End Time
                        .Cells(rowFind + ii + 2, 5).Value = itemEndDate
                        .Cells(rowFind + ii + 2, 6).Value = itemEndTime

                    End With
                    proOrderCount = proOrderCount + 1
                End If
            End If
        End If
nxtForStatement:
        ii = ii + 1
    Loop
    
    If counterCheck = proOrderCount Then
        MsgBox "Please check item: " & Sheets("ITEMIZERS").Cells(i, 1).Value
    End If

Next

Exit Sub

incorrectRow:
    
    Sheets("SUMMARIZETABLE").Range("E" & rowFind + ii & ":H" & rowFind + ii).Interior.ColorIndex = 3
    MsgBox " Please Enter Correct Process Order / Please Check if Order has been entered correctly in Summarized Table"

End Sub

