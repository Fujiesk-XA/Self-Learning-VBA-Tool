'Only works for one sheet - Automatically updates after a change has been entered into the cell, Sheet name is NON-OPERATIONAL DT, This only works for 24 hour
Option Explicit

Private Const COL_S_REFS As Integer = 9
Private Const COL_S_DATES As Integer = 5
Private Const COL_S_TIMES As Integer = 6
Private Const COL_E_DATES As Integer = 7
Private Const COL_E_TIMES As Integer = 8

Private Sub Worksheet_Change(ByVal Target As Range)
    
    Dim introw As Integer
    Dim intBlanks As Integer
    Dim blnFound As Boolean
    Dim lngRef As Long
    Dim startDate As Double
    Dim startTime As Double
    Dim checkCombineTime As Double
    
    Dim i, ii, iii As Double
    Dim lastColumn As Double
    Dim RowofTarget As Double
    Dim lastRow As Double
    
    introw = 2
    intBlanks = 0
    blnFound = False
    
  'REF target excution Declare target.row > 1 so only change intarget row

If Target.Row > 1 Then
    Select Case Target.Column
        Case COL_S_REFS 'check if col is number 9 
            'Check if it is in number and not blank
            If (IsNumeric(Cells(Target.Row, Target.Column).Value) And _
                (Cells(Target.Row, Target.Column).Value <> "")) Then
                'populate from database and redirect to start time
                lngRef = CLng(Cells(Target.Row, Target.Column).Value) 'Copy min into Ref as Long
                If lngRef >= 1440 Then
                    blnFound = False
                Else
                    startDate = Cells(Target.Row, COL_S_DATES).Value
                    startTime = Cells(Target.Row, COL_S_TIMES).Value
                    checkCombineTime = startTime + lngRef / 60 / 24
                    
                    If checkCombineTime >= 1 Then
                        With Sheets("NON-OPERATIONAL DT")
                            .Cells(Target.Row, COL_E_DATES).Value = .Cells(Target.Row, COL_S_DATES).Value + 1
                            .Cells(Target.Row, COL_E_TIMES).Value = lngRef / 60 / 24 - (1 - .Cells(Target.Row, COL_S_TIMES).Value)
                        End With
                    Else
                        With Sheets("NON-OPERATIONAL DT")
                            .Cells(Target.Row, COL_E_DATES).Value = .Cells(Target.Row, COL_S_DATES).Value
                            .Cells(Target.Row, COL_E_TIMES).Value = .Cells(Target.Row, COL_S_TIMES).Value + lngRef / 60 / 24
                        End With
                    End If
                    blnFound = True
                End If
            End If
                
            If (blnFound = False) Then
                With Sheets("NON-OPERATIONAL DT")
                    .Cells(Target.Row, COL_E_DATES).Value = ""
                    .Cells(Target.Row, COL_E_TIMES).Value = ""
                End With
            End If
            
     End Select
End If

End Sub
