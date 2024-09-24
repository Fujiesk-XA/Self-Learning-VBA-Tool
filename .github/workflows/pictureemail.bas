Sub Huddle_Email()
    'assigning variables
    Dim OutApp As Object
    Dim OMail As Object
    Dim strbody As String
    Dim strbody2 As String
    Dim MakeJPG As String
    Dim loc As Date
    Dim signature As String
    
    'to make the macro run faster
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    loc = WorksheetFunction.WorkDay(Date, -1)
    'assigning objects to what application you want
    Set OutApp = CreateObject("Outlook.Application")
    Set OMail = OutApp.CreateItem(olMailItem)
    
    With OMail
        .Display
    End With
    
    signature = OMail.HTMLBody
    
    'email body
    strbody = "Hi Team," & "<br><br>" & "Please see below the updated huddle sheet with key points for you to discuss in your team huddles." & _
    " The hard copy of sheet can be found in office where radio/RF gun batteries are charged." & "<br><br>"
    
    'creating jpg file of the range
    MakeJPG = CopyRangeToJPG("PRODUCTION", "A1:F36")
    
    
    On Error Resume Next
    With OMail
        .To = "Gaurav.Kalia@wonderbrands.com; Nanthan.Nadarajah@wonderbrands.com; Thanendran.Parameshwaran@wonderbrands.com; jasjit.pannu@fgfbrands.com; Rick.Lam@wonderbrands.com; Claudia.Rios@wonderbrands.com; Manvesh.Paul@wonderbrands.com; Rosa.Arias@wonderbrands.com; Virajkumar.Patel@wonderbrands.com; Sivacumaran.Namasivayam@wonderbrands.com; Bharatkumar.Vekaria@wonderbrands.com; Prasanth.Karunairooban@wonderbrands.com" 'gaurav, nanthan, thanuesh, jasjit
        .Subject = "Secretariat DMS Daily Huddle Sheet"
        .Attachments.Add MakeJPG, 1, 0
        .HTMLBody = "<html><p>" & strbody & "</p><img src=""cid:NamePicture.jpg"" width=800 height=1200></html>" & "<br><br>" & signature
        .Display
    End With
    source_file = ThisWorkbook.FullName
    OMail.Attachments.Add source_file
     
    On Error GoTo 0
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OMail = Nothing
    Set OutApp = Nothing
    
End Sub


Function CopyRangeToJPG(NameWorksheet As String, RangeAddress As String) As String
   Dim PictureRange As Range
   
   With ActiveWorkbook
        On Error Resume Next
        .Worksheets(NameWorksheet).Activate
        Set PictureRange = .Worksheets(NameWorksheet).Range(RangeAddress)
        
        If PictureRange Is Nothing Then
            MsgBox "Sorry this is not a correct range"
            On Error GoTo 0
            Exit Function
        End If
        'copies range as picture in clipboard
        PictureRange.CopyPicture (xlScreen)
        'since its a pic and not a range now modify it to fit
        With .Worksheets(NameWorksheet).ChartObjects.Add(PictureRange.Left, PictureRange.Top, PictureRange.Width, PictureRange.Height)
            .Activate
            .Chart.Paste
            .Chart.Export Environ$("temp") & Application.PathSeparator & "NamePicture.jpg", "JPG"
        End With
        .Worksheets(NameWorksheet).ChartObjects(.Worksheets(NameWorksheet).ChartObjects.Count).Delete
    End With
    
    CopyRangeToJPG = Environ$("temp") & Application.PathSeparator & "NamePicture.jpg"
    Set PictureRange = Nothing
End Function

