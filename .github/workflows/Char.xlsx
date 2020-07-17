'Changing Column Number to Corresponding Letters

Public Function charconversion(letterNum As Double) As String

    Dim multipleLetter As Integer
    Dim mulitpleLetter2 As Integer
    Dim firstLetter, secondLetter, thirdLetter As String

    If letterNum <= 26 Then
        charconversion = Chr(letterNum + 64)
    ElseIf (letterNum <= 676) Then
        multipleLetter = Application.WorksheetFunction.RoundDown(letterNum / 26, 0)
        firstLetter = Chr(multipleLetter + 64)
        secondLetter = Chr(letterNum - (26 * multipleLetter) + 64)
        charconversion = firstLetter & secondLetter
    ElseIf letterNum > 676 Then
        multipleLetter = Application.WorksheetFunction.RoundDown(letterNum / 26, 0)
        multipleLetter2 = Application.WorksheetFunction.RoundDown(multipleLetter / 26, 0)
        firstLetter = Chr(multipleLetter2 + 64)
        secondLetter = Chr(multipleLetter - (26 * multipleLetter2) + 64)
        thirdLetter = Chr(letterNum - (26 * multipleLetter) + 64)
        charconversion = firstLetter & secondLetter & thirdLetter
    End If
    
End Function

