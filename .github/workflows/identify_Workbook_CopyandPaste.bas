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


Sub copyOrderHeaders()

Dim i, ii As Integer
Dim wb As Workbook
Dim ws As Worksheet
Dim sFile As String
Dim sFile2 As String
Dim answer As Integer

'Loop to determine sFile from Sheet Name "Must be unique compare to the other notebook
For Each wb In Workbooks
    If wb.Name <> "PERSONAL.XLSB" Then
        If wb.Name <> "PERSONAL (Autosaved).xlsb" Then
            If fileNamesfunc(wb.Name) = True Then
                sFile = wb.Name
            End If
        End If
    End If
Next

For Each wb In Workbooks
    If wb.Name <> "PERSONAL.XLSB" Then
        If wb.Name <> "PERSONAL (Autosaved).xlsb" Then
            If fileNamesfunc(wb.Name) = False Then
                sFile2 = wb.Name
                answer = MsgBox("Is this Your File " & sFile2, vbYesNo, "Confirmation")
                    If answer = vbYes Then
                        sFile2 = wb.Name
                        Exit For
                    ElseIf answer = vbNo Then
                        sFile2 = ""
                    End If
            End If
        End If
    End If
Next

If sFile2 = "" Then
    MsgBox "File did not detect export.", vCCritical, "Missing File(s)"
    Exit Sub
End If
'Delete
Windows(sFile).Activate
Sheets("MacroTest").Select ' Location of the sheets of copied
Cells.Select
Cells.Delete Shift:=xlUp
'Copy
Windows(sFile2).Activate
Cells.Select
Selection.Copy
Windows(sFile).Activate
Sheets("MacroTest").Select ' Location of the sheets of copied
Cells(1, 1).Select
ActiveSheet.Paste
Cells(1, 1).Copy
Workbooks(sFile2).Close savechanges:=False

End Sub
