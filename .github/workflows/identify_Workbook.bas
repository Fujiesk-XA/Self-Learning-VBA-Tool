'Using worksheet name to identify workbook

Public Function fileNamesfunc(filename1 As String) As Boolean

    Dim wb As Workbook
    Dim ws As Worksheet

    fileNamesfunc = False

        For Each ws In Workbooks(filename1).Worksheets
            If ws.Name = "MacroInputs" Then
                fileNamesfunc = True
            End If
        Next
    
End Function
