Public Function add_preferment(ByVal tMixTime As Date, ByVal sPreferment As String, ByVal dQty As Double, ByVal sproductName As String) As Boolean
    On Error GoTo errHandle
    Dim intRow As Integer
    Dim intBlanks As Integer
    Dim timeStart As Date
    Dim timeEnd As Date
    Dim blnFound As Boolean
    
    intRow = 3
    intBlanks = 0
    blnFound = False
    
'    If (tMixTime > CDate("23:59:00")) Then tMixTime = DateAdd("d", -1, tMixTime)
    
    Do While intBlanks < 1
        If (sPreferment = Worksheets(SHEET_DATA).Cells(intRow, 11)) Then
      '            'check if product is in range
            timeStart = CDate(Worksheets(SHEET_DATA).Cells(intRow, 12))
            timeEnd = CDate(Worksheets(SHEET_DATA).Cells(intRow, 13))
      
      'check if product is in range
            blnFound = False
            If (Hour(tMixTime) >= Hour(timeStart)) And (Hour(tMixTime) <= Hour(timeEnd)) Then
            
                blnFound = minute_calc(tMixTime, timeStart, timeEnd)
                
            ElseIf (Hour(timeStart) > Hour(timeEnd)) Then   ' dough spans over 0:00
                If ((Hour(tMixTime) >= Hour(timeStart)) And (Hour(tMixTime) <= 23) Or _
                    (Hour(tMixTime) <= Hour(timeEnd))) Then
                    
                    blnFound = minute_calc(tMixTime, timeStart, timeEnd)
                    
                End If
            End If
            
            If (blnFound = True) Then
                Worksheets(SHEET_DATA).Cells(intRow, 14) = CDbl(Worksheets(SHEET_DATA).Cells(intRow, 14)) + dQty
                If Trim(Worksheets(SHEET_DATA).Cells(intRow, 15)) = "" Then
                    Worksheets(SHEET_DATA).Cells(intRow, 15) = FormatDateTime(tMixTime, vbShortTime) & " - " & sProductName
                Else
                    Worksheets(SHEET_DATA).Cells(intRow, 15) = Worksheets(SHEET_DATA).Cells(intRow, 15) & ", " & _
                            FormatDateTime(tMixTime, vbShortTime) & " - " & sDoughName
                End If
                add_preferment = True
                Exit Do
            End If
            
        End If
        intRow = intRow + 1
    Loop

ResExit:
    Exit Function
errHandle:
    add_preferment = False
    MsgBox Err.Description & vbNewLine & _
            "Error adding parbaked preferment to data sheet." & vbNewLine & _
            "Please contact ", vbOKOnly
End Function
    
    '   If (dblOWD > 0) Then blnAdded = Worksheets(strPreferments).add_preferment(timeMix, "MaterialCode", , strProductName)
    
Private Function minute_calc(ByVal Mix As Date, _
        ByVal StarterStart As Date, _
        ByVal StarterEnd As Date) As Boolean
                
      If (Hour(Mix) = Hour(StarterStart)) Then
        'check the minute
        minute_calc = (Minute(Mix) >= Minute(StarterStart))
    ElseIf (Hour(Mix) = Hour(StarterEnd)) Then
        'check the minute
        minute_calc = (Minute(Mix) <= Minute(StarterEnd))
    Else
        minute_calc = True
    End If

End Function
