'Module 1
Option Explicit

Public Const SH_TMTN_UPLOAD As String = "Tab - 7 TMTN UPLOAD Format"
Public Const SH_TEMP_ONE As String = "Tab 6 - Daily and Weekly print"


Public Const COL_T1_TMNAME As Integer = 1
Public Const COL_T1_TMQUA As Integer = 3
Public Const COL_T1_TMSTARTT As Integer = 6
Public Const COL_T1_TMENDT As Integer = 8

Public Const COL_T2_TMNAME As Integer = 1
Public Const COL_T2_TMQUA As Integer = 3
Public Const COL_T2_TMSTARTT As Integer = 6
Public Const COL_T2_TMENDT As Integer = 8

Public Const COL_T3_TMNAME As Integer = 1
Public Const COL_T3_TMQUA As Integer = 3
Public Const COL_T3_TMSTARTT As Integer = 6
Public Const COL_T3_TMENDT As Integer = 8

Public Const COL_T4_TMNAME As Integer = 1
Public Const COL_T4_TMQUA As Integer = 3
Public Const COL_T4_TMSTARTT As Integer = 6
Public Const COL_T4_TMENDT As Integer = 8

Public Const COL_T5_TMNAME As Integer = 1
Public Const COL_T5_TMQUA As Integer = 3
Public Const COL_T5_TMSTARTT As Integer = 6
Public Const COL_T5_TMENDT As Integer = 8

Public Const COL_DAYSHM_TMNAME As Integer = 1
Public Const COL_DAYSHM_TMSTARTT As Integer = 3

Public Const COL_NIGHTSHM_TMNAME As Integer = 1
Public Const COL_NIGHTSHM_TMSTARTT As Integer = 3

Public Const COL_DAYSHTMS_TMNAME As Integer = 2
Public Const COL_DAYSHTMS_TMQUA As Integer = 4

Public Const COL_NIGHTSHTMS_TMNAME As Integer = 2
Public Const COL_NIGHTSHTMS_TMQUA As Integer = 4


Public Const COL_S_REF As Integer = 4
Public Const COL_S_REF1 As Integer = 6
Public Const COL_S_REF2 As Integer = 8
Public Const COL_S_REF3 As Integer = 10
Public Const COL_S_REF4 As Integer = 12
Public Const COL_S_REF5 As Integer = 14
Public Const COL_S_REF6 As Integer = 16
Public Const COL_S_REF7 As Integer = 18
Public Const COL_S_REF8 As Integer = 20
Public Const COL_S_REF9 As Integer = 22
Public Const COL_S_REF10 As Integer = 24
Public Const COL_S_REF11 As Integer = 26
Public Const COL_S_REF12 As Integer = 28
Public Const COL_S_REF13 As Integer = 30
Public Const COL_S_REF14 As Integer = 32
Public Const COL_S_REF15 As Integer = 34
Public Const COL_S_REF16 As Integer = 36
Public Const COL_S_REF17 As Integer = 38
Public Const COL_S_REF18 As Integer = 40
Public Const COL_S_REF19 As Integer = 42
Public Const COL_S_REF20 As Integer = 44
Public Const COL_S_REF21 As Integer = 46
Public Const COL_S_REF22 As Integer = 48
Public Const COL_S_REF23 As Integer = 50
Public Const COL_S_REF24 As Integer = 52
Public Const COL_S_REF25 As Integer = 54
Public Const COL_S_REF26 As Integer = 56
Public Const COL_S_REF27 As Integer = 58
Public Const COL_S_REF28 As Integer = 60
Public Const COL_S_REF29 As Integer = 62
Public Const COL_S_REF30 As Integer = 64
Public Const COL_S_REF31 As Integer = 66
Public Const COL_S_REF32 As Integer = 68
Public Const COL_S_REF33 As Integer = 70
Public Const COL_S_REF34 As Integer = 72



Public Function format_L1(ByVal rowSelector As Double, ByVal columnSelector As Double) As Boolean
    
    On Error GoTo ErrHandle
        ActiveSheet.Cells(rowSelector, columnSelector + 1).Value = Format(ActiveSheet.Cells(rowSelector, columnSelector + 1).Value, "hh:mm AM/PM")
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 15773696
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 15773696
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With

    format_L1 = True
    
FuncExit:
    Exit Function
ErrHandle:
    format_L1 = False
    GoTo FuncExit
End Function
Public Function format_OFF(ByVal rowSelector As Double, ByVal columnSelector As Double) As Boolean
    
    On Error GoTo ErrHandle
        ActiveSheet.Cells(rowSelector, columnSelector + 1).Value = Format(ActiveSheet.Cells(rowSelector, columnSelector + 1).Value, "hh:mm AM/PM")
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With

    format_OFF = True
    
FuncExit:
    Exit Function
ErrHandle:
    format_OFF = False
    GoTo FuncExit
End Function

Public Function format_FILL(ByVal rowSelector As Double, ByVal columnSelector As Double) As Boolean
    
    On Error GoTo ErrHandle
        ActiveSheet.Cells(rowSelector, columnSelector + 1).Value = Format(ActiveSheet.Cells(rowSelector, columnSelector + 1).Value, "hh:mm AM/PM")
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = -0.499984740745262
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = -0.499984740745262
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With

    format_FILL = True
    
FuncExit:
    Exit Function
ErrHandle:
    format_FILL = False
    GoTo FuncExit
End Function

Public Function format_L2(ByVal rowSelector As Double, ByVal columnSelector As Double) As Boolean
    
    On Error GoTo ErrHandle
        ActiveSheet.Cells(rowSelector, columnSelector + 1).Value = Format(ActiveSheet.Cells(rowSelector, columnSelector + 1).Value, "hh:mm AM/PM")
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105 'Pink
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105 'Pink
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With

    format_L2 = True
    
FuncExit:
    Exit Function
ErrHandle:
    format_L2 = False
    GoTo FuncExit
End Function

Public Function format_L3(ByVal rowSelector As Double, ByVal columnSelector As Double) As Boolean
    
    On Error GoTo ErrHandle
        ActiveSheet.Cells(rowSelector, columnSelector + 1).Value = Format(ActiveSheet.Cells(rowSelector, columnSelector + 1).Value, "hh:mm AM/PM")
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 16775743 'Cyan
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 16775743 'Cyan
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With

    format_L3 = True
    
FuncExit:
    Exit Function
ErrHandle:
    format_L3 = False
    GoTo FuncExit
End Function

Public Function format_L4(ByVal rowSelector As Double, ByVal columnSelector As Double) As Boolean
    
    On Error GoTo ErrHandle
        ActiveSheet.Cells(rowSelector, columnSelector + 1).Value = Format(ActiveSheet.Cells(rowSelector, columnSelector + 1).Value, "hh:mm AM/PM")
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535 'Yellow
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535 'Yellow
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With

    format_L4 = True
    
FuncExit:
    Exit Function
ErrHandle:
    format_L4 = False
    GoTo FuncExit
End Function
Public Function format_L5(ByVal rowSelector As Double, ByVal columnSelector As Double) As Boolean
    
    On Error GoTo ErrHandle
        ActiveSheet.Cells(rowSelector, columnSelector + 1).Value = Format(ActiveSheet.Cells(rowSelector, columnSelector + 1).Value, "hh:mm AM/PM")
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 5296274
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector + 1).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 5296274
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With ActiveSheet.Cells(rowSelector, columnSelector).Font
            .Name = "Calibri"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Bold = True
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With

    format_L5 = True
    
FuncExit:
    Exit Function
ErrHandle:
    format_L5 = False
    GoTo FuncExit
End Function
'Module 2
Sub printScheduleTempOne()

Dim i, ii, iii, iiii, pagesToBePrinted, startColM, startColTMS As Integer
Dim lineOneRowDay, lineTwoRowDay, lineThreeRowDay, lineFourRowDay, lineFiveRowDay As Integer
Dim lineOneRowNight, lineTwoRowNight, lineThreeRowNight, lineFourRowNight, lineFiveRowNight As Integer
Dim weekDayAdd As Integer
Dim SH_DAY_SHIFTM, SH_NIGHT_SHIFTM, SH_DAY_TMS, SH_NIGHT_TMS, SH_KEYROLES As String
Dim fndLineOneRow, fndLineTwoRow, fndLineThreeRow, fndLineFourRow, fndLineFiveRow As String
Dim firstRowDTMS, firstRowNTMS As String
Dim weekDay As String
Dim lastColDayMix, lastColNightMix, lastRowDayMix, lastRowNightMix As Integer
Dim lastColDayTM, lastRowDayTM, lastColNightTM, lastRowNightTM As Integer
Dim startWkDate, endWkDate As Date
Dim startColDM, endColDM, startColNM, endColNM As Integer
Dim startColDTM, endColDTM, startColNTM, endColNTM As Integer
Dim startdateAsStr As String
Dim lineNumber, lineCheck As String
Dim startLineRowNumD, endLineRowNumD, startLineRowNumN, endLineRowNumN As Integer
Dim startColKR, endColKR As Integer
Dim lastRowKR, lastColKR, mornlastRowKR As Integer

'names to be replace: SH_TEMP_ONE,COL_T1_TMNAME,COL_T1_TMQUA,COL_T1_TMSTARTT,COL_T1_TMENDT

'Sheet Name
SH_DAY_SHIFTM = Sheets(SH_TEMP_ONE).Cells(3, 14).Value
SH_NIGHT_SHIFTM = Sheets(SH_TEMP_ONE).Cells(4, 14).Value
SH_DAY_TMS = Sheets(SH_TEMP_ONE).Cells(3, 16).Value
SH_NIGHT_TMS = Sheets(SH_TEMP_ONE).Cells(4, 16).Value
SH_KEYROLES = Sheets(SH_TEMP_ONE).Cells(6, 14).Value

'Sheets Name

startWkDate = CDate(Sheets(SH_TEMP_ONE).Cells(4, 21).Value)
endWkDate = CDate(Sheets(SH_TEMP_ONE).Cells(4, 22).Value)
lineNumber = Sheets(SH_TEMP_ONE).Cells(5, 14).Value
weekDay = Sheets(SH_TEMP_ONE).Cells(3, 23).Value

If weekDay <> "ALL" Then
    pagesToBePrinted = 0 'indicator for i so first one does not need to add extra pages
Else
    pagesToBePrinted = 6
End If


Select Case Sheets(SH_TEMP_ONE).Cells(3, 23).Value
    Case "SUN":
        weekDayAdd = 0
    Case "MON":
        weekDayAdd = 1
    Case "TUE":
        weekDayAdd = 2
    Case "WED":
        weekDayAdd = 3
    Case "THURS":
        weekDayAdd = 4
    Case "FRI":
        weekDayAdd = 5
    Case "SAT":
        weekDayAdd = 6
    Case "All":
        weekDayAdd = 0
End Select


fndLineOneRow = "L1.No"
fndLineTwoRow = "L2.No"
fndLineThreeRow = "L3.No"
fndLineFourRow = "L4.No"
fndLineFiveRow = "L5.No"


'Needs to rewrite for Night, in part 2
Set found1 = Sheets(SH_DAY_TMS).Columns(1).Find(What:=fndLineOneRow, LookIn:=xlValues)
Set found2 = Sheets(SH_DAY_TMS).Columns(1).Find(What:=fndLineTwoRow, LookIn:=xlValues)
Set found3 = Sheets(SH_DAY_TMS).Columns(1).Find(What:=fndLineThreeRow, LookIn:=xlValues)
Set found4 = Sheets(SH_DAY_TMS).Columns(1).Find(What:=fndLineFourRow, LookIn:=xlValues)
Set found5 = Sheets(SH_DAY_TMS).Columns(1).Find(What:=fndLineFiveRow, LookIn:=xlValues)

If found1 Is Nothing Then
        MsgBox ("Cannot find Determine LINE 1 in DAY TMS")
        Exit Sub
End If
If found2 Is Nothing Then
        MsgBox ("Cannot find Determine LINE 2 in in DAY TMS")
        Exit Sub
End If
If found3 Is Nothing Then
        MsgBox ("Cannot find Determine LINE 3 in Day TMS")
        Exit Sub
End If
If found4 Is Nothing Then
        MsgBox ("Cannot find Determine LINE 4 in in DAY TMS")
        Exit Sub
End If
If found5 Is Nothing Then
    MsgBox ("Cannot find Determine LINE 5 in in DAY TMS")
    Exit Sub
End If

lineOneRowDay = found1.Row
lineTwoRowDay = found2.Row
lineThreeRowDay = found3.Row
lineFourRowDay = found4.Row
lineFiveRowDay = found5.Row

' End of Copy of Part2 for night
Set found1 = Sheets(SH_NIGHT_TMS).Columns(1).Find(What:=fndLineOneRow, LookIn:=xlValues)
Set found2 = Sheets(SH_NIGHT_TMS).Columns(1).Find(What:=fndLineTwoRow, LookIn:=xlValues)
Set found3 = Sheets(SH_NIGHT_TMS).Columns(1).Find(What:=fndLineThreeRow, LookIn:=xlValues)
Set found4 = Sheets(SH_NIGHT_TMS).Columns(1).Find(What:=fndLineFourRow, LookIn:=xlValues)
Set found5 = Sheets(SH_NIGHT_TMS).Columns(1).Find(What:=fndLineFiveRow, LookIn:=xlValues)

If found1 Is Nothing Then
        MsgBox ("Cannot find Determine LINE 1 in Night TMS")
        Exit Sub
End If
If found2 Is Nothing Then
        MsgBox ("Cannot find Determine LINE 2 in in Night TMS")
        Exit Sub
End If
If found3 Is Nothing Then
        MsgBox ("Cannot find Determine LINE 3 in Night TMS")
        Exit Sub
End If
If found4 Is Nothing Then
        MsgBox ("Cannot find Determine LINE 4 in in Night TMS")
        Exit Sub
End If
If found5 Is Nothing Then
    MsgBox ("Cannot find Determine LINE 5 in in Night TMS")
    Exit Sub
End If

lineOneRowNight = found1.Row
lineTwoRowNight = found2.Row
lineThreeRowNight = found3.Row
lineFourRowNight = found4.Row
lineFiveRowNight = found5.Row
' End of night

'Find LastRow morning KeyRoles
Set found1 = Sheets(SH_KEYROLES).Cells.Find(What:="NIGHT SHIFT KEY ROLES", LookIn:=xlValues)

If found1 Is Nothing Then
        MsgBox ("NIGHT SHIFT KEY ROLES Text could not be founr")
        Exit Sub
End If

mornlastRowKR = found1.Row

'Check Start Date and End Date of all the sheets
With Sheets(SH_DAY_SHIFTM)
    lastColDayMix = .Cells(3, Columns.Count).End(xlToLeft).Column
    lastRowDayMix = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastColDayMix
        If .Cells(3, i).Value = startWkDate Then
            startColDM = i
        ElseIf .Cells(3, i).Value = endWkDate Then
            endColDM = i
        End If
    Next
End With
With Sheets(SH_NIGHT_SHIFTM)
    lastColNightMix = .Cells(3, Columns.Count).End(xlToLeft).Column
    lastRowNightMix = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastColNightMix
        If .Cells(3, i).Value = startWkDate Then
            startColNM = i
        ElseIf .Cells(3, i).Value = endWkDate Then
            endColNM = i
        End If
    Next
End With
With Sheets(SH_DAY_TMS)
    lastColDayTM = .Cells(3, Columns.Count).End(xlToLeft).Column
    lastRowDayTM = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastColDayTM
        If .Cells(3, i).Value = startWkDate Then
            startColDTM = i
        ElseIf .Cells(3, i).Value = endWkDate Then
            endColDTM = i
        End If
    Next
End With
With Sheets(SH_NIGHT_TMS)
    lastColNightTM = .Cells(3, Columns.Count).End(xlToLeft).Column
    lastRowNightTM = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastColNightTM
        If .Cells(3, i).Value = startWkDate Then
            startColNTM = i
        ElseIf .Cells(3, i).Value = endWkDate Then
            endColNTM = i
        End If
    Next
End With
'/////////////////////////////////////////////////////////////////////
'KeyRoles Column
With Sheets(SH_KEYROLES)
    lastColKR = .Cells(3, Columns.Count).End(xlToLeft).Column
    lastRowKR = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastColKR
        If .Cells(3, i).Value = startWkDate Then
            startColKR = i
        ElseIf .Cells(3, i).Value = endWkDate Then
            endColKR = i
        End If
    Next
End With
'/////////////////////////////////////////////////////////////
Select Case lineNumber
    Case "LINE 1":
        lineCheck = "L1"
        startLineRowNumD = lineOneRowDay
        endLineRowNumD = lineTwoRowDay - 2
        startLineRowNumN = lineOneRowNight
        endLineRowNumN = lineTwoRowNight - 2
    Case "LINE 2":
        lineCheck = "L2"
        startLineRowNumD = lineTwoRowDay
        endLineRowNumD = lineThreeRowDay - 2
        startLineRowNumN = lineTwoRowNight
        endLineRowNumN = lineThreeRowNight - 2
    Case "LINE 3":
        lineCheck = "L3"
        startLineRowNumD = lineThreeRowDay
        endLineRowNumD = lineFourRowDay - 2
        startLineRowNumN = lineThreeRowNight
        endLineRowNumN = lineFourRowNight - 2
    Case "LINE 4":
        lineCheck = "L4"
        startLineRowNumD = lineFourRowDay
        endLineRowNumD = lineFiveRowDay - 2
        startLineRowNumN = lineFourRowNight
        endLineRowNumN = lineFiveRowNight - 2
    Case "LINE 5":
        lineCheck = "L5"
        startLineRowNumD = lineFiveRowDay
        endLineRowNumD = lastRowDayTM
        startLineRowNumN = lineFiveRowNight
        endLineRowNumN = lastRowNightTM
End Select


For i = 0 To pagesToBePrinted
    'Day Shift
    iii = 12
    Sheets(SH_TEMP_ONE).Range("B12:J66").ClearContents
    Sheets(SH_TEMP_ONE).Cells(10, 1).Value = Sheets(SH_DAY_TMS).Cells(3, startColDTM + weekDayAdd + i).Value
    Sheets(SH_TEMP_ONE).Cells(10, 7).Value = "Morning Shift"
    For ii = 4 To mornlastRowKR
        With Sheets(SH_TEMP_ONE)
            If Sheets(SH_KEYROLES).Cells(ii, startColKR + weekDayAdd * 2 + i + i).Value = lineCheck Then
                .Cells(iii, 2).Value = Sheets(SH_KEYROLES).Cells(ii, 1).Value
                .Cells(iii, 3).Value = Sheets(SH_KEYROLES).Cells(ii, 3).Value
                .Cells(iii, 4).Value = Sheets(SH_KEYROLES).Cells(ii, 2).Value
                .Cells(iii, 7).Value = Sheets(SH_KEYROLES).Cells(ii, startColKR + weekDayAdd * 2 + i + i + 1).Value
                If .Cells(iii, 2).Value <> "" Then
                    .Cells(iii, 9).Value = .Cells(iii, 7).Value + 12 / 24
                End If
                iii = iii + 1
            End If
        End With
    Next
    For ii = 4 To lastRowDayMix
        With Sheets(SH_TEMP_ONE)
            If Sheets(SH_DAY_SHIFTM).Cells(ii, startColDM + weekDayAdd * 2 + i + i).Value = lineCheck Then
                .Cells(iii, 2).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 1).Value
                .Cells(iii, 3).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 2).Value
                .Cells(iii, 4).Value = "Mixers"
                .Cells(iii, 7).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, startColDM + weekDayAdd * 2 + i + i + 1).Value
                If .Cells(iii, 2).Value <> "" Then
                    .Cells(iii, 9).Value = .Cells(iii, 7).Value + 12 / 24
                End If
                iii = iii + 1
            End If
        End With
    Next
    For ii = startLineRowNumD + 2 To endLineRowNumD 'Stays with the Format
        With Sheets(SH_TEMP_ONE)
            If Sheets(SH_DAY_TMS).Cells(ii, startColDTM + weekDayAdd + i).Value <> "OFF" Then
                If Sheets(SH_DAY_TMS).Cells(ii, 1).Value <> "TOTAL" Then
                    .Cells(iii, 2).Value = Sheets(SH_DAY_TMS).Cells(ii, 2).Value
                    .Cells(iii, 3).Value = Sheets(SH_DAY_TMS).Cells(ii, 5).Value
                    .Cells(iii, 4).Value = Sheets(SH_DAY_TMS).Cells(ii, 3).Value
                    .Cells(iii, 7).Value = Sheets(SH_DAY_TMS).Cells(ii, startColDTM + weekDayAdd + i).Value
                    If .Cells(iii, 2).Value <> "" Then
                        .Cells(iii, 9).Value = .Cells(iii, 7).Value + 12 / 24
                    End If
                    iii = iii + 1
                End If
            End If
        End With
    Next
    
    Worksheets(SH_TEMP_ONE).Activate
    Worksheets(SH_TEMP_ONE).PrintOut Copies:=1, Collate:=True
    'Night Shift
    Sheets(SH_TEMP_ONE).Range("B12:J66").ClearContents
    Sheets(SH_TEMP_ONE).Cells(10, 1).Value = Sheets(SH_NIGHT_TMS).Cells(3, startColNTM + weekDayAdd + i).Value
    Sheets(SH_TEMP_ONE).Cells(10, 7).Value = "Night Shift"
    iii = 12
    For ii = mornlastRowKR To lastRowKR
        With Sheets(SH_TEMP_ONE)
            If Sheets(SH_KEYROLES).Cells(ii, startColKR + weekDayAdd * 2 + i + i).Value = lineCheck Then
                .Cells(iii, 2).Value = Sheets(SH_KEYROLES).Cells(ii, 1).Value
                .Cells(iii, 3).Value = Sheets(SH_KEYROLES).Cells(ii, 3).Value
                .Cells(iii, 4).Value = Sheets(SH_KEYROLES).Cells(ii, 2).Value
                .Cells(iii, 7).Value = Sheets(SH_KEYROLES).Cells(ii, startColKR + weekDayAdd * 2 + i + i + 1).Value
                If .Cells(iii, 2).Value <> "" Then
                    .Cells(iii, 9).Value = .Cells(iii, 7).Value + 12 / 24
                End If
                iii = iii + 1
            End If
        End With
    Next
    For ii = 4 To lastRowNightMix
        With Sheets(SH_TEMP_ONE)
            If Sheets(SH_NIGHT_SHIFTM).Cells(ii, startColNM + weekDayAdd * 2 + i + i).Value = lineCheck Then
                .Cells(iii, 2).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 1).Value
                .Cells(iii, 3).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 2).Value
                .Cells(iii, 4).Value = "Mixers"
                .Cells(iii, 7).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, startColNM + weekDayAdd * 2 + i + i + 1).Value
                If .Cells(iii, 2).Value <> "" Then
                    .Cells(iii, 9).Value = .Cells(iii, 7).Value + 12 / 24
                End If
                iii = iii + 1
            End If
        End With
    Next
    For ii = startLineRowNumN + 2 To endLineRowNumN
        With Sheets(SH_TEMP_ONE)
            If Sheets(SH_NIGHT_TMS).Cells(ii, startColNTM + weekDayAdd + i).Value <> "OFF" Then
                If Sheets(SH_NIGHT_TMS).Cells(ii, 1).Value <> "TOTAL" Then
                    .Cells(iii, 2).Value = Sheets(SH_NIGHT_TMS).Cells(ii, 2).Value
                    .Cells(iii, 3).Value = Sheets(SH_NIGHT_TMS).Cells(ii, 5).Value
                    .Cells(iii, 4).Value = Sheets(SH_NIGHT_TMS).Cells(ii, 3).Value
                    .Cells(iii, 7).Value = Sheets(SH_NIGHT_TMS).Cells(ii, startColNTM + weekDayAdd + i).Value
                If .Cells(iii, 2).Value <> "" Then
                    .Cells(iii, 9).Value = .Cells(iii, 7).Value + 12 / 24
                End If
                    iii = iii + 1
                End If
            End If
        End With
    Next
    
    Worksheets(SH_TEMP_ONE).Activate
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
Next

End Sub

'Module 3
Sub copyAllTM()

Dim i, ii, iii, iiii, pagesToBePrinted, startColM, startColTMS As Integer
Dim lineOneRowDay, lineTwoRowDay, lineThreeRowDay, lineFourRowDay, lineFiveRowDay As Integer
Dim lineOneRowNight, lineTwoRowNight, lineThreeRowNight, lineFourRowNight, lineFiveRowNight As Integer
Dim weekDayAdd As Integer
Dim SH_DAY_SHIFTM, SH_NIGHT_SHIFTM, SH_DAY_TMS, SH_NIGHT_TMS, SH_KEYROLES As String
Dim fndLineOneRow, fndLineTwoRow, fndLineThreeRow, fndLineFourRow, fndLineFiveRow As String
Dim firstRowDTMS, firstRowNTMS As String
Dim weekDay As String
Dim lastColDayMix, lastColNightMix, lastRowDayMix, lastRowNightMix As Integer
Dim lastColDayTM, lastRowDayTM, lastColNightTM, lastRowNightTM As Integer
Dim lastColLeaders, startColLead, endColLead, lastRowLeaders As Integer
Dim startWkDate, endWkDate As Date
Dim startColDM, endColDM, startColNM, endColNM As Integer
Dim startColDTM, endColDTM, startColNTM, endColNTM As Integer
Dim startdateAsStr As String
Dim lineNumber, lineCheck As String
Dim startLineRowNumD, endLineRowNumD, startLineRowNumN, endLineRowNumN As Integer
Dim startColKR, endColKR As Integer
Dim lastRowKR, lastColKR, mornlastRowKR As Integer
Dim nameCheck As Boolean

'Sheet Name
SH_DAY_SHIFTM = Sheets(SH_TEMP_ONE).Cells(3, 14).Value
SH_NIGHT_SHIFTM = Sheets(SH_TEMP_ONE).Cells(4, 14).Value
SH_DAY_TMS = Sheets(SH_TEMP_ONE).Cells(3, 16).Value
SH_NIGHT_TMS = Sheets(SH_TEMP_ONE).Cells(4, 16).Value
SH_KEYROLES = Sheets(SH_TEMP_ONE).Cells(6, 14).Value

'Sheets Name

startWkDate = CDate(Sheets("Tab - 7 TMTN UPLOAD Format").Cells(4, 6).Value)
endWkDate = CDate(Sheets("Tab - 7 TMTN UPLOAD Format").Cells(4, 12).Value)

Application.ScreenUpdating = False

Sheets("Tab - 7 TMTN UPLOAD Format").Range("A6:L9999").ClearContents

With Sheets(SH_DAY_SHIFTM)
    lastColDayMix = .Cells(3, Columns.Count).End(xlToLeft).Column
    lastRowDayMix = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastColDayMix
        If .Cells(3, i).Value = startWkDate Then
            startColDM = i
        ElseIf .Cells(3, i).Value = endWkDate Then
            endColDM = i
        End If
    Next
End With
With Sheets(SH_NIGHT_SHIFTM)
    lastColNightMix = .Cells(3, Columns.Count).End(xlToLeft).Column
    lastRowNightMix = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastColNightMix
        If .Cells(3, i).Value = startWkDate Then
            startColNM = i
        ElseIf .Cells(3, i).Value = endWkDate Then
            endColNM = i
        End If
    Next
End With
With Sheets(SH_DAY_TMS)
    lastColDayTM = .Cells(3, Columns.Count).End(xlToLeft).Column
    lastRowDayTM = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastColDayTM
        If .Cells(3, i).Value = startWkDate Then
            startColDTM = i
        ElseIf .Cells(3, i).Value = endWkDate Then
            endColDTM = i
        End If
    Next
End With
With Sheets(SH_NIGHT_TMS)
    lastColNightTM = .Cells(3, Columns.Count).End(xlToLeft).Column
    lastRowNightTM = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastColNightTM
        If .Cells(3, i).Value = startWkDate Then
            startColNTM = i
        ElseIf .Cells(3, i).Value = endWkDate Then
            endColNTM = i
        End If
    Next
End With
'/////////////////////////////////////////////////////////////////////
'KeyRoles Column
With Sheets(SH_KEYROLES)
    lastColKR = .Cells(3, Columns.Count).End(xlToLeft).Column
    lastRowKR = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastColKR
        If .Cells(3, i).Value = startWkDate Then
            startColKR = i
        ElseIf .Cells(3, i).Value = endWkDate Then
            endColKR = i
        End If
    Next
End With
'////////////////////////////////////////////////'
'Leaders
With Sheets("Tab - 3.5 Leaders")
    lastColLeaders = .Cells(3, Columns.Count).End(xlToLeft).Column
    lastRowLeaders = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastColLeaders
        If .Cells(3, i).Value = startWkDate Then
            startColLead = i
        ElseIf .Cells(3, i).Value = endWkDate Then
            endColLead = i
        End If
    Next
End With
i = 6
nameCheck = False

For ii = 4 To lastRowDayMix
    For iii = startColDM To endColDM
        With Sheets("Tab - 7 TMTN UPLOAD Format")
            If UCase(Sheets(SH_DAY_SHIFTM).Cells(ii, iii).Value) = "L1" Then
                .Cells(i, 1).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 2).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(2, 2).Value
                .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(3, 6).Value
                .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(13, 2).Value
                .Cells(i, ((iii - startColDM) / 2) + 6).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_DAY_SHIFTM).Cells(ii, iii).Value) = "L2" Then
                .Cells(i, 1).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 2).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(3, 2).Value
                .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(3, 6).Value
                .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(13, 2).Value
                .Cells(i, ((iii - startColDM) / 2) + 6).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_DAY_SHIFTM).Cells(ii, iii).Value) = "L3" Then
                .Cells(i, 1).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 2).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(4, 2).Value
                .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(3, 6).Value
                .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(13, 2).Value
                .Cells(i, ((iii - startColDM) / 2) + 6).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_DAY_SHIFTM).Cells(ii, iii).Value) = "L4" Then
                .Cells(i, 1).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 2).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(5, 2).Value
                .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(3, 6).Value
                .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(13, 2).Value
                .Cells(i, ((iii - startColDM) / 2) + 6).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_DAY_SHIFTM).Cells(ii, iii).Value) = "L5" Then
                .Cells(i, 1).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, 2).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(6, 2).Value
                .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(3, 6).Value
                .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(13, 2).Value
                .Cells(i, ((iii - startColDM) / 2) + 6).Value = Sheets(SH_DAY_SHIFTM).Cells(ii, iii + 1).Value
                nameCheck = True
            End If
        End With
    Next
    If nameCheck = True Then
        i = i + 1
        nameCheck = False
    End If
Next
'//////////////////////////////////////////////////Beginning of NightMix
For ii = 4 To lastRowNightMix
    For iii = startColNM To endColNM
        With Sheets("Tab - 7 TMTN UPLOAD Format")
            If UCase(Sheets(SH_NIGHT_SHIFTM).Cells(ii, iii).Value) = "L1" Then
                .Cells(i, 1).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 2).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(2, 2).Value
                .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(3, 6).Value
                .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(13, 2).Value
                .Cells(i, ((iii - startColNM) / 2) + 6).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_NIGHT_SHIFTM).Cells(ii, iii).Value) = "L2" Then
                .Cells(i, 1).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 2).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(3, 2).Value
                .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(3, 6).Value
                .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(13, 2).Value
                .Cells(i, ((iii - startColNM) / 2) + 6).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_NIGHT_SHIFTM).Cells(ii, iii).Value) = "L3" Then
                .Cells(i, 1).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 2).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(4, 2).Value
                .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(3, 6).Value
                .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(13, 2).Value
                .Cells(i, ((iii - startColNM) / 2) + 6).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_NIGHT_SHIFTM).Cells(ii, iii).Value) = "L4" Then
                .Cells(i, 1).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 2).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(5, 2).Value
                .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(3, 6).Value
                .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(13, 2).Value
                .Cells(i, ((iii - startColNM) / 2) + 6).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_NIGHT_SHIFTM).Cells(ii, iii).Value) = "L5" Then
                .Cells(i, 1).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, 2).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(6, 2).Value
                .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(3, 6).Value
                .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(13, 2).Value
                .Cells(i, ((iii - startColNM) / 2) + 6).Value = Sheets(SH_NIGHT_SHIFTM).Cells(ii, iii + 1).Value
                nameCheck = True
            End If
        End With
    Next
    If nameCheck = True Then
        i = i + 1
        nameCheck = False
    End If
Next
'///////////////////////////////////////////////////////////////////////////////////////////////////KeyRoles
For ii = 4 To lastRowKR
    For iii = startColKR To endColKR
        With Sheets("Tab - 7 TMTN UPLOAD Format")
            If UCase(Sheets(SH_KEYROLES).Cells(ii, iii).Value) = "L1" Then
                .Cells(i, 1).Value = Sheets(SH_KEYROLES).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_KEYROLES).Cells(ii, 3).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(2, 2).Value
                If UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "PTL" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(6, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(14, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "DVO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(4, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(15, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "OVO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(5, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(16, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "PKO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(7, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(17, 2).Value
                End If
                .Cells(i, ((iii - startColKR) / 2) + 6).Value = Sheets(SH_KEYROLES).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, iii).Value) = "L2" Then
                .Cells(i, 1).Value = Sheets(SH_KEYROLES).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_KEYROLES).Cells(ii, 3).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(3, 2).Value
                If UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "PTL" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(6, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(14, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "DVO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(4, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(15, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "OVO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(5, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(16, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "PKO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(7, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(17, 2).Value
                End If
                .Cells(i, ((iii - startColKR) / 2) + 6).Value = Sheets(SH_KEYROLES).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, iii).Value) = "L3" Then
                .Cells(i, 1).Value = Sheets(SH_KEYROLES).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_KEYROLES).Cells(ii, 3).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(4, 2).Value
                If UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "PTL" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(6, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(14, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "DVO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(4, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(15, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "OVO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(5, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(16, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "PKO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(7, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(17, 2).Value
                End If
                .Cells(i, ((iii - startColKR) / 2) + 6).Value = Sheets(SH_KEYROLES).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, iii).Value) = "L4" Then
                .Cells(i, 1).Value = Sheets(SH_KEYROLES).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_KEYROLES).Cells(ii, 3).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(5, 2).Value
                If UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "PTL" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(6, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(14, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "DVO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(4, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(15, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "OVO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(5, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(16, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "PKO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(7, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(17, 2).Value
                End If
                .Cells(i, ((iii - startColKR) / 2) + 6).Value = Sheets(SH_KEYROLES).Cells(ii, iii + 1).Value
                nameCheck = True
            ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, iii).Value) = "L5" Then
                .Cells(i, 1).Value = Sheets(SH_KEYROLES).Cells(ii, 1).Value
                .Cells(i, 2).Value = Sheets(SH_KEYROLES).Cells(ii, 3).Value
                .Cells(i, 3).Value = Sheets("Activity Code & Cost Center").Cells(6, 2).Value
                If UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "PTL" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(6, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(14, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "DVO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(4, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(15, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "OVO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(5, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(16, 2).Value
                ElseIf UCase(Sheets(SH_KEYROLES).Cells(ii, 2).Value) = "PKO" Then
                    .Cells(i, 4).Value = Sheets("Activity Code & Cost Center").Cells(7, 6).Value
                    .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(17, 2).Value
                End If
                .Cells(i, ((iii - startColKR) / 2) + 6).Value = Sheets(SH_KEYROLES).Cells(ii, iii + 1).Value
                nameCheck = True
            End If
        End With
    Next
    If nameCheck = True Then
        i = i + 1
        nameCheck = False
    End If
Next
'Morning TM
For ii = 4 To lastRowDayTM
    For iii = startColDTM To endColDTM
        With Sheets("Tab - 7 TMTN UPLOAD Format")
            If Sheets(SH_DAY_TMS).Cells(ii, iii).Value <> "OFF" Then
                If Sheets(SH_DAY_TMS).Cells(ii, 2).Value <> "" Then
                    If Sheets(SH_DAY_TMS).Cells(ii, 2).Value <> "Team Member Name" Then
                        .Cells(i, 1).Value = Sheets(SH_DAY_TMS).Cells(ii, 2).Value
                        .Cells(i, 2).Value = Sheets(SH_DAY_TMS).Cells(ii, 5).Value
                        .Cells(i, 3).Value = Sheets(SH_DAY_TMS).Cells(ii, 7).Value
                        .Cells(i, 4).Value = Sheets(SH_DAY_TMS).Cells(ii, 6).Value
                        .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(18, 2).Value
                        .Cells(i, ((iii - startColDTM)) + 6).Value = Sheets(SH_DAY_TMS).Cells(ii, iii).Value
                        nameCheck = True
                    End If
                End If
            End If
        End With
    Next
    If nameCheck = True Then
        i = i + 1
        nameCheck = False
    End If
Next
'Night Tm
For ii = 4 To lastRowNightTM
    For iii = startColNTM To endColNTM
        With Sheets("Tab - 7 TMTN UPLOAD Format")
            If Sheets(SH_NIGHT_TMS).Cells(ii, iii).Value <> "OFF" Then
                If Sheets(SH_NIGHT_TMS).Cells(ii, 2).Value <> "" Then
                    If Sheets(SH_NIGHT_TMS).Cells(ii, 2).Value <> "Team Member Name" Then
                        .Cells(i, 1).Value = Sheets(SH_NIGHT_TMS).Cells(ii, 2).Value
                        .Cells(i, 2).Value = Sheets(SH_NIGHT_TMS).Cells(ii, 5).Value
                        .Cells(i, 3).Value = Sheets(SH_NIGHT_TMS).Cells(ii, 7).Value
                        .Cells(i, 4).Value = Sheets(SH_NIGHT_TMS).Cells(ii, 6).Value
                        .Cells(i, 5).Value = Sheets("Activity Code & Cost Center").Cells(18, 2).Value
                        .Cells(i, ((iii - startColNTM)) + 6).Value = Sheets(SH_NIGHT_TMS).Cells(ii, iii).Value
                        nameCheck = True
                    End If
                End If
            End If
        End With
    Next
    If nameCheck = True Then
        i = i + 1
        nameCheck = False
    End If
Next
'Leaders Print
For ii = 5 To lastRowLeaders
    For iii = startColLead To endColLead
        With Sheets("Tab - 7 TMTN UPLOAD Format")
            If Sheets("Tab - 3.5 Leaders").Cells(ii, iii).Value <> "OFF" Then
                If Sheets("Tab - 3.5 Leaders").Cells(4, iii).Value = "Start Time" Then
                    
                    .Cells(i, 1).Value = Sheets("Tab - 3.5 Leaders").Cells(ii, 1).Value
                    .Cells(i, 2).Value = Sheets("Tab - 3.5 Leaders").Cells(ii, 3).Value
                    .Cells(i, 3).Value = Sheets("Tab - 3.5 Leaders").Cells(ii, 5).Value
                    .Cells(i, 4).Value = Sheets("Tab - 3.5 Leaders").Cells(ii, 4).Value
                    If .Cells(i, 5).Value = "" Then
                        .Cells(i, 5).Value = Sheets("Tab - 3.5 Leaders").Cells(ii, iii + 2).Value
                    End If
                    .Cells(i, ((iii - startColLead) / 3) + 6).Value = Sheets("Tab - 3.5 Leaders").Cells(ii, iii).Value
                    nameCheck = True
                    
                End If
            End If
        End With
    Next
    If nameCheck = True Then
        i = i + 1
        nameCheck = False
    End If
Next
Application.ScreenUpdating = True


End Sub
'Module 4
Sub refereshALL()


Dim lastRow As Integer
Dim i, ii As Integer

SH_MIX_IND = ActiveSheet.Name

Sheets(SH_MIX_IND).Cells(2, 66).Value = "On"

Application.ScreenUpdating = False

lastRow = Sheets(SH_MIX_IND).Cells(Rows.Count, 1).End(xlUp).Row
For ii = 4 To 58
    For i = 4 To lastRow
        Sheets(SH_MIX_IND).Cells(i, ii).Select
        Selection.Value = Selection.FormulaR1C1
    Next
    ii = ii + 1
Next

Application.ScreenUpdating = True

End Sub

Sub refereshALL3()


Dim lastRow As Integer
Dim i, ii As Integer

SH_KEYROLES = ActiveSheet.Name
Sheets(SH_KEYROLES).Cells(2, 77).Value = "On"

Application.ScreenUpdating = False

lastRow = Sheets(SH_KEYROLES).Cells(Rows.Count, 1).End(xlUp).Row
For ii = 4 To 72
    For i = 4 To lastRow
        If Sheets(SH_KEYROLES).Cells(i, 2) = "PTL" _
            Or Sheets(SH_KEYROLES).Cells(i, 2) = "DVO" _
            Or Sheets(SH_KEYROLES).Cells(i, 2) = "OVO" _
            Or Sheets(SH_KEYROLES).Cells(i, 2) = "PKO" Then
            Sheets(SH_KEYROLES).Cells(i, ii).Select
            Selection.Value = Selection.FormulaR1C1
        End If
    Next
    ii = ii + 1
Next

Application.ScreenUpdating = True



End Sub
