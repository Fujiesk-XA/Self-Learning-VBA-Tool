Private Sub ComboBox102_Change()
    Worksheets("ITEMIZERS").Range("AO6").Value = Right(ComboBox102.Value, 10)
    TextBox108_Change
End Sub

Private Sub ComboBox103_Change()
    Worksheets("ITEMIZERS").Range("AR6").Value = ComboBox103.Value
    TextBox108_Change
End Sub

Private Sub ComboBox104_Change()
    Worksheets("ITEMIZERS").Range("AS6").Value = Right(ComboBox104.Value, 10)
    TextBox108_Change
End Sub

Private Sub ComboBox105_Change()
    Worksheets("ITEMIZERS").Range("AV6").Value = ComboBox105.Value
    TextBox108_Change
End Sub

Private Sub ComboBox202_Change()
    Worksheets("ITEMIZERS").Range("AO7").Value = Right(ComboBox202.Value, 10)
    TextBox208_Change
End Sub

Private Sub ComboBox203_Change()
    Worksheets("ITEMIZERS").Range("AR7").Value = ComboBox203.Value
    TextBox208_Change
End Sub

Private Sub ComboBox204_Change()
    Worksheets("ITEMIZERS").Range("AS7").Value = Right(ComboBox204.Value, 10)
    TextBox208_Change
End Sub

Private Sub ComboBox205_Change()
    Worksheets("ITEMIZERS").Range("AV7").Value = ComboBox205.Value
    TextBox208_Change
End Sub

Private Sub CommandButton1_Click()
    Dim lastRow As Double
    
    lastRow = Sheets("OPERATIONAL DT").Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets("Operational DT").Cells(lastRow + 1, 1).Value = TextBox101.Value
    Sheets("Operational DT").Cells(lastRow + 1, 2).Value = TextBox102.Value
    Sheets("Operational DT").Cells(lastRow + 1, 3).Value = ComboBox101.Value
    
    If OptionButton1.Value = True Then
        Sheets("OPERATIONAL DT").Cells(lastRow + 1, 5).Value = "YES"
    ElseIf OptionButton2.Value = True Then
        Sheets("OPERATIONAL DT").Cells(lastRow + 1, 6).Value = "YES"
    ElseIf OptionButton3.Value = True Then
        Sheets("OPERATIONAL DT").Cells(lastRow + 1, 7).Value = "YES"
    End If
    
    Sheets("OPERATIONAL DT").Cells(lastRow + 1, 8).Value = Right(ComboBox102.Value, 10)
    Sheets("OPERATIONAL DT").Cells(lastRow + 1, 9).Value = TextBox103.Value & ":" & TextBox104.Value & " " & ComboBox103.Value
    Sheets("OPERATIONAL DT").Cells(lastRow + 1, 10).Value = Right(ComboBox104.Value, 10)
    Sheets("OPERATIONAL DT").Cells(lastRow + 1, 11).Value = TextBox105.Value & ":" & TextBox105.Value & " " & ComboBox105.Value
    
    Sheets("OPERATIONAL DT").Cells(lastRow + 1, 12).Value = TextBox108.Value
    Sheets("OPERATIONAL DT").Cells(lastRow + 1, 13).Value = TextBox107.Value
    Sheets("OPERATIONAL DT").Cells(lastRow + 1, 1).Select
    
End Sub


Private Sub CommandButton2_Click()

    Dim lastRow As Double
    Dim answer As Integer
    
    answer = MsgBox("Are you sure you want to Delete the Last Post", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
    
    If answer = vbYes Then
        lastRow = Sheets("OPERATIONAL DT").Cells(Rows.Count, 1).End(xlUp).Row
        Sheets("OPERATIONAL DT").Rows(lastRow).EntireRow.Delete
    End If
    
End Sub

Private Sub CommandButton3_Click()
    Dim lastRow2 As Double
    
    lastRow2 = Sheets("NON-OPERATIONAL DT").Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 1).Value = TextBox201.Value
    Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 2).Value = TextBox202.Value
    Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 14).Value = ComboBox201.Value
    
    If OptionButton4.Value = True Then
        Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 15).Value = "YES"
    ElseIf OptionButton5.Value = True Then
        Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 16).Value = "YES"
    ElseIf OptionButton6.Value = True Then
        Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 17).Value = "YES"
    End If
    
    Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 5).Value = Right(ComboBox202.Value, 10)
    Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 6).Value = TextBox203.Value & ":" & TextBox204.Value & " " & ComboBox203.Value
    Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 7).Value = Right(ComboBox204.Value, 10)
    Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 8).Value = TextBox205.Value & ":" & TextBox205.Value & " " & ComboBox205.Value
    
    Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 9).Value = TextBox208.Value
    Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 18).Value = TextBox207.Value
    Sheets("NON-OPERATIONAL DT").Cells(lastRow2 + 1, 1).Select

End Sub

Private Sub CommandButton4_Click()

    Dim lastRow As Double
    Dim answer As Integer
    
    answer = MsgBox("Are you sure you want to Delete the Last Post", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
    
    If answer = vbYes Then
        lastRow = Sheets("NON-OPERATIONAL DT").Cells(Rows.Count, 1).End(xlUp).Row
        Sheets("NON-OPERATIONAL DT").Rows(lastRow).EntireRow.Delete
    End If
End Sub

Private Sub CommandButton5_Click()
    Dim lastRow2 As Double
    
    lastRow3 = Sheets("FOODWASTES").Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets("FOODWASTES").Cells(lastRow3 + 1, 1).Value = Right(ComboBox302.Value, 10)
    Sheets("FOODWASTES").Cells(lastRow3 + 1, 2).Value = ComboBox303.Value
    Sheets("FOODWASTES").Cells(lastRow3 + 1, 3).Value = ComboBox301.Value
    Sheets("FOODWASTES").Cells(lastRow3 + 1, 4).Value = TextBox303.Value
    Sheets("FOODWASTES").Cells(lastRow3 + 1, 5).Value = TextBox301.Value
    Sheets("FOODWASTES").Cells(lastRow3 + 1, 6).Value = TextBox302.Value
    Sheets("FOODWASTES").Cells(lastRow3 + 1, 10).Value = TextBox304.Value
    Sheets("FOODWASTES").Cells(lastRow3 + 1, 1).Select
End Sub

Private Sub CommandButton6_Click()
    Dim lastRow As Double
    Dim answer As Integer
    
    answer = MsgBox("Are you sure you want to Delete the Last Post", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
    
    If answer = vbYes Then
        lastRow = Sheets("FOODWASTES").Cells(Rows.Count, 1).End(xlUp).Row
        Sheets("FOODWASTES").Rows(lastRow).EntireRow.Delete
    End If
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label308_Click()

End Sub

Private Sub MultiPage1_Change()

lastRow = Sheets("OPERATIONAL DT").Cells(Rows.Count, 1).End(xlUp).Row
lastRow2 = Sheets("NON-OPERATIONAL DT").Cells(Rows.Count, 1).End(xlUp).Row
lastRow3 = Sheets("FOODWASTES").Cells(Rows.Count, 1).End(xlUp).Row

    If MultiPage1.Value = 0 Then
        Sheets("OPERATIONAL DT").Activate
        Sheets("OPERATIONAL DT").Cells(lastRow, 1).Select
    ElseIf MultiPage1.Value = 1 Then
        Sheets("NON-OPERATIONAL DT").Activate
        Sheets("NON-OPERATIONAL DT").Cells(lastRow2, 1).Select
    ElseIf MultiPage1.Value = 2 Then
        Sheets("FOODWASTES").Activate
        Sheets("FOODWASTES").Cells(lastRow3, 5).Select
    End If
        
End Sub


Private Sub TextBox103_Change()
    Worksheets("ITEMIZERS").Range("AP6").Value = TextBox103.Value
    TextBox108_Change
End Sub

Private Sub TextBox104_Change()
    Worksheets("ITEMIZERS").Range("AQ6").Value = TextBox104.Value
    TextBox108_Change
End Sub

Private Sub TextBox105_Change()
    Worksheets("ITEMIZERS").Range("AT6").Value = TextBox105.Value
    TextBox108_Change
End Sub

Private Sub TextBox106_Change()
    Worksheets("ITEMIZERS").Range("AU6").Value = TextBox106.Value
    TextBox108_Change
End Sub

Private Sub TextBox108_Change()
    On Error Resume Next
    UserForm1.TextBox108.Value = Round(Worksheets("ITEMIZERS").Range("AW06").Value, 0)
    UserForm1.Repaint
End Sub

Private Sub TextBox203_Change()
    Worksheets("ITEMIZERS").Range("AP7").Value = TextBox203.Value
    TextBox208_Change
End Sub

Private Sub TextBox204_Change()
    Worksheets("ITEMIZERS").Range("AQ7").Value = TextBox204.Value
    TextBox208_Change
End Sub

Private Sub TextBox205_Change()
    Worksheets("ITEMIZERS").Range("AT7").Value = TextBox105.Value
    TextBox208_Change
End Sub

Private Sub TextBox206_Change()
    Worksheets("ITEMIZERS").Range("AU7").Value = TextBox206.Value
    TextBox208_Change
End Sub

Private Sub TextBox208_Change()
    On Error Resume Next
    UserForm1.TextBox208.Value = Round(Worksheets("ITEMIZERS").Range("AW07").Value, 0)
    UserForm1.Repaint
End Sub

Private Sub TextBox303_Change()

End Sub

Private Sub UserForm_Initialize()

'ORDER & MATERIAL\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim lastRow As Integer
lastRow = Sheets("OPERATIONAL DT").Cells(Rows.Count, 1).End(xlUp).Row
lastRow2 = Sheets("NON-OPERATIONAL DT").Cells(Rows.Count, 1).End(xlUp).Row
lastRow3 = Sheets("FOODWASTES").Cells(Rows.Count, 1).End(xlUp).Row
TextBox101.Value = Sheets("OPERATIONAL DT").Cells(lastRow, 1).Value
TextBox102.Value = Sheets("OPERATIONAL DT").Cells(lastRow, 2).Value

TextBox201.Value = Sheets("NON-OPERATIONAL DT").Cells(lastRow2, 1).Value
TextBox202.Value = Sheets("NON-OPERATIONAL DT").Cells(lastRow2, 2).Value

TextBox301.Value = Sheets("FOODWASTES").Cells(lastRow3, 5).Value
TextBox302.Value = Sheets("FOODWASTES").Cells(lastRow3, 6).Value

UserForm1.TextBox103.Value = "12"
UserForm1.TextBox104.Value = "00"
UserForm1.TextBox105.Value = "11"
UserForm1.TextBox106.Value = "59"

UserForm1.TextBox203.Value = "12"
UserForm1.TextBox204.Value = "00"
UserForm1.TextBox205.Value = "11"
UserForm1.TextBox206.Value = "59"

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'EquipmentList \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    ComboBox101.ColumnCount = 1
    ComboBox101.BoundColumn = 1
    ComboBox101.TextColumn = 1
    ComboBox101.List = Sheets("ITEMIZERS").Range("P6:P30").Value
    ComboBox101.Value = ("")
    
    ComboBox201.ColumnCount = 1
    ComboBox201.BoundColumn = 1
    ComboBox201.TextColumn = 1
    ComboBox201.List = Sheets("ITEMIZERS").Range("AL6:AL30").Value
    ComboBox201.Value = ("")
    
    ComboBox301.ColumnCount = 1
    ComboBox301.BoundColumn = 1
    ComboBox301.TextColumn = 1
    ComboBox301.AddItem ("MIXING - FOOD WASTE (KG)")
    ComboBox301.List(0, 1) = 0
    ComboBox301.AddItem ("FRITSCH LINE - FOOD WASTE (KG)")
    ComboBox301.List(1, 1) = 1
    ComboBox301.AddItem ("PROOFER - FOOD WASTE (KG)")
    ComboBox301.List(2, 1) = 2
    ComboBox301.AddItem ("SPRAYS & SEEDER - FOOD WASTE (KG)")
    ComboBox301.List(3, 1) = 3
    ComboBox301.AddItem ("OVEN - FOOD WASTE (KG)")
    ComboBox301.List(4, 1) = 4
    ComboBox301.AddItem ("SLICER - FOOD WASTE (KG)")
    ComboBox301.List(5, 1) = 5
    ComboBox301.AddItem ("PACKAGING - FOOD WASTE (KG)")
    ComboBox301.List(6, 1) = 6
    ComboBox301.AddItem ("END OF SHIFT REWORK (KG)")
    ComboBox301.List(7, 1) = 7
    ComboBox301.Value = ("")
    
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    ComboBox102.ColumnCount = 1
    ComboBox102.BoundColumn = 1
    ComboBox102.TextColumn = 1
    ComboBox102.AddItem (Format(Date - 1, "dddd, mm/dd/yyyy"))
    ComboBox102.List(0, 1) = 0
    ComboBox102.AddItem (Format(Date, "dddd, mm/dd/yyyy"))
    ComboBox102.List(1, 1) = 1
    ComboBox102.AddItem (Format(Date + 1, "dddd, mm/dd/yyyy"))
    ComboBox102.List(2, 1) = 2
    ComboBox102.Value = (Format(Date, "dddd, mm/dd/yyyy"))
    
    ComboBox104.ColumnCount = 1
    ComboBox104.BoundColumn = 1
    ComboBox104.TextColumn = 1
    ComboBox104.AddItem (Format(Date - 1, "dddd, mm/dd/yyyy"))
    ComboBox104.List(0, 1) = 0
    ComboBox104.AddItem (Format(Date, "dddd, mm/dd/yyyy"))
    ComboBox104.List(1, 1) = 1
    ComboBox104.AddItem (Format(Date + 1, "dddd, mm/dd/yyyy"))
    ComboBox104.List(2, 1) = 2
    ComboBox104.Value = (Format(Date, "dddd, mm/dd/yyyy"))
    
    ComboBox202.ColumnCount = 1
    ComboBox202.BoundColumn = 1
    ComboBox202.TextColumn = 1
    ComboBox202.AddItem (Format(Date - 1, "dddd, mm/dd/yyyy"))
    ComboBox202.List(0, 1) = 0
    ComboBox202.AddItem (Format(Date, "dddd, mm/dd/yyyy"))
    ComboBox202.List(1, 1) = 1
    ComboBox202.AddItem (Format(Date + 1, "dddd, mm/dd/yyyy"))
    ComboBox202.List(2, 1) = 2
    ComboBox202.Value = (Format(Date, "dddd, mm/dd/yyyy"))
    
    ComboBox204.ColumnCount = 1
    ComboBox204.BoundColumn = 1
    ComboBox204.TextColumn = 1
    ComboBox204.AddItem (Format(Date - 1, "dddd, mm/dd/yyyy"))
    ComboBox204.List(0, 1) = 0
    ComboBox204.AddItem (Format(Date, "dddd, mm/dd/yyyy"))
    ComboBox204.List(1, 1) = 1
    ComboBox204.AddItem (Format(Date + 1, "dddd, mm/dd/yyyy"))
    ComboBox204.List(2, 1) = 2
    ComboBox204.Value = (Format(Date, "dddd, mm/dd/yyyy"))
    
    ComboBox302.ColumnCount = 1
    ComboBox302.BoundColumn = 1
    ComboBox302.TextColumn = 1
    ComboBox302.AddItem (Format(Date - 3, "dddd, mm/dd/yyyy"))
    ComboBox302.List(0, 1) = 0
    ComboBox302.AddItem (Format(Date - 2, "dddd, mm/dd/yyyy"))
    ComboBox302.List(1, 1) = 1
    ComboBox302.AddItem (Format(Date - 1, "dddd, mm/dd/yyyy"))
    ComboBox302.List(2, 1) = 2
    ComboBox302.AddItem (Format(Date, "dddd, mm/dd/yyyy"))
    ComboBox302.List(3, 1) = 3
    ComboBox302.AddItem (Format(Date + 1, "dddd, mm/dd/yyyy"))
    ComboBox302.List(4, 1) = 4
    ComboBox302.AddItem (Format(Date + 2, "dddd, mm/dd/yyyy"))
    ComboBox302.List(5, 1) = 5
    ComboBox302.Value = (Format(Date, "dddd, mm/dd/yyyy"))
    
    ComboBox303.ColumnCount = 1
    ComboBox303.BoundColumn = 1
    ComboBox303.TextColumn = 1
    ComboBox303.AddItem ("SHIFT 1")
    ComboBox303.List(0, 1) = 0
    ComboBox303.AddItem ("SHIFT 2")
    ComboBox303.List(1, 1) = 1
    ComboBox303.AddItem ("SHIFT 3")
    ComboBox303.List(2, 1) = 2
    ComboBox303.Value = ("")
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    ComboBox103.ColumnCount = 1
    ComboBox103.BoundColumn = 1
    ComboBox103.TextColumn = 1
    ComboBox103.AddItem ("AM")
    ComboBox103.List(0, 1) = 0
    ComboBox103.AddItem ("PM")
    ComboBox103.List(1, 1) = 1
    ComboBox103.Value = ("AM")

    ComboBox105.ColumnCount = 1
    ComboBox105.BoundColumn = 1
    ComboBox105.TextColumn = 1
    ComboBox105.AddItem ("AM")
    ComboBox105.List(0, 1) = 0
    ComboBox105.AddItem ("PM")
    ComboBox105.List(1, 1) = 1
    ComboBox105.Value = ("PM")

    ComboBox203.ColumnCount = 1
    ComboBox203.BoundColumn = 1
    ComboBox203.TextColumn = 1
    ComboBox203.AddItem ("AM")
    ComboBox203.List(0, 1) = 0
    ComboBox203.AddItem ("PM")
    ComboBox203.List(1, 1) = 1
    ComboBox203.Value = ("AM")

    ComboBox205.ColumnCount = 1
    ComboBox205.BoundColumn = 1
    ComboBox205.TextColumn = 1
    ComboBox205.AddItem ("AM")
    ComboBox205.List(0, 1) = 0
    ComboBox205.AddItem ("PM")
    ComboBox205.List(1, 1) = 1
    ComboBox205.Value = ("PM")

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
''Format Box Size

Label101.Font.Size = 13
Label102.Font.Size = 13
Label103.Font.Size = 13
Label104.Font.Size = 13
Label105.Font.Size = 13
Label106.Font.Size = 13
Label107.Font.Size = 13
Label108.Font.Size = 13

ComboBox101.Font.Size = 13
ComboBox102.Font.Size = 11
ComboBox103.Font.Size = 13
ComboBox104.Font.Size = 11
ComboBox105.Font.Size = 13

TextBox101.Font.Size = 13
TextBox102.Font.Size = 13
TextBox103.Font.Size = 13
TextBox104.Font.Size = 13
TextBox105.Font.Size = 13
TextBox106.Font.Size = 13

CommandButton1.Font.Size = 13

Label201.Font.Size = 13
Label202.Font.Size = 13
Label203.Font.Size = 13
Label204.Font.Size = 13
Label205.Font.Size = 13
Label206.Font.Size = 13
Label207.Font.Size = 13
Label208.Font.Size = 13

ComboBox201.Font.Size = 13
ComboBox202.Font.Size = 11
ComboBox203.Font.Size = 13
ComboBox204.Font.Size = 11
ComboBox205.Font.Size = 13

TextBox201.Font.Size = 13
TextBox202.Font.Size = 13
TextBox203.Font.Size = 13
TextBox204.Font.Size = 13
TextBox205.Font.Size = 13
TextBox206.Font.Size = 13

CommandButton3.Font.Size = 13

Label301.Font.Size = 13
Label302.Font.Size = 13
Label303.Font.Size = 13
Label304.Font.Size = 13
Label305.Font.Size = 13
Label306.Font.Size = 13
Label307.Font.Size = 13
Label308.Font.Size = 13

ComboBox301.Font.Size = 13
ComboBox302.Font.Size = 11
ComboBox303.Font.Size = 13

TextBox301.Font.Size = 13
TextBox302.Font.Size = 13
TextBox303.Font.Size = 13

CommandButton5.Font.Size = 13

End Sub


