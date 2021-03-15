Attribute VB_Name = "Emailer"
Sub Emailer_ActiveSub()
    subcontractor = ActiveCell.Value
    Emailer_Sub (subcontractor)
End Sub

Sub Emailer_EachYes()
    Dim subList As Collection
    ' Get Collection of Subs to Send
    Set subList = New Collection
    subcontractorsArray = Worksheets("Email").ListObjects("Sub_List").DataBodyRange.Value
    
    For rt = 1 To UBound(subcontractorsArray)
        If subcontractorsArray(rt, 2) = "YES" Then
            e = CollectionAddUnique(subList, CStr(subcontractorsArray(rt, 1)))
        End If
    Next rt
    ' ---------------------------------------------
    
    For Each subcont In subList
        Emailer_Sub (subcont)
    Next subcont

End Sub

Sub Emailer_Sub(inputSub As String)
    On Error GoTo errHandle
    Dim subcontractor As String
    Dim rngTable As Range
    Dim SENDorDISPLAYemailVar As String
    Dim emailSubjectVar As String
    Dim emailBodyVar As String
    
    subcontractor = inputSub 'ActiveCell.Value
    
    'Set cl = Worksheets("Contact Log").ListObjects("ContactLog")
    
    subContactsArray = Worksheets("Contact List").ListObjects("Contacts_Table").DataBodyRange.Value
    'subcontractorsArray = Worksheets("Updates").ListObjects("Subcontractor_Table").DataBodyRange.Value
    
    ' Get Collection of Unique Subs
    'Set subList = New Collection
    'subLogArray = Worksheets("Email Table").ListObjects("EMAIL_TABLE").DataBodyRange.Value
    
    ' Filter by sub / adjust row heights for printing
    emailHideClosed = Range("Email_Hide_Closed").Value
    Worksheets("Email Table").ListObjects("EMAIL_TABLE").Range.AutoFilter Field:=7, Criteria1:=subcontractor
    Worksheets("Email Table").ListObjects("Email_Table").Range.AutoFilter Field:=4
    
    If emailHideClosed = "HIDE" Then
        Worksheets("Email Table").ListObjects("Email_Table").Range.AutoFilter Field:=4, Criteria1 _
            :=Array("Assigned to Sub", "Design Review", "Draft", "Reviewed"), Operator:= _
            xlFilterValues
    End If
    
    Worksheets("Email Table").ListObjects("EMAIL_TABLE").Range.EntireRow.AutoFit
    Worksheets("Email Table").Columns("F:G").Hidden = True
    Set rngTable = Worksheets("Email Table").UsedRange.SpecialCells(xlCellTypeVisible)
    
    ' -----------------------------------------
    ' get emails
    to_Field = "" ' Clear from Previous subs
    For ctsc = 1 To UBound(subContactsArray)
        If subContactsArray(ctsc, 1) = subcontractor Then
            If to_Field = "" Then
                to_Field = subContactsArray(ctsc, 2) + " <" + subContactsArray(ctsc, 4) + ">"
            Else
                to_Field = to_Field + "; " + subContactsArray(ctsc, 2) + " <" + subContactsArray(ctsc, 4) + ">"
            End If
        End If
    Next ctsc
    
    ' -----------------------------------------
    ' Debug.Print to_field
    
    ' get templates
    emailSubjectVar = Range("Email_Subject").Value
    emailBodyVar = Range("Email_Body").Value
    emailPathToSignVar = Range("Email_Signature_Path").Value
    emailCCVar = Range("Email_CC").Value
    SENDorDISPLAYemailVar = Range("SENDorDISPLAYemail").Value
    emailAttachment1 = Range("Email_Attachment1").Value
    emailAttachment2 = Range("Email_Attachment2").Value
    
    If Dir(emailPathToSignVar) <> "" Then
        emailSignature = GetBoiler(emailPathToSignVar)
        startEmailSignPath = Environ("appdata") & _
                "\Microsoft\Signatures\"
        startEmailSignPath = Replace(startEmailSignPath, " ", "%20")
        emailSignature = Replace(emailSignature, "src=""", "src=""" & startEmailSignPath)
        emailSignature = Replace(emailSignature, "files/", "files\")
    Else
        emailSignature = ""
    End If
    
    AddLog (emailSignature)
    emailBodyVar = emailBodyVar + emailSignature
    

    ' -----------------------------------------
    
    ' update templates with variables
    emailSubjectVar = Replace(emailSubjectVar, "<<SUB NAME>>", subcontractor)
    emailBodyVar = Replace(emailBodyVar, "<<SUB NAME>>", subcontractor)
    
    emailSubjectVar = Replace(emailSubjectVar, "<<CAMRON DATE>>", Format(Now(), "yyyy-mm-dd"))
    emailBodyVar = Replace(emailBodyVar, "<<CAMRON DATE>>", Format(Now(), "yyyy-mm-dd"))
    
    ' -----------------------------------------
    ' update Template with EMAIL_TABLE
        'testVar = RangetoHTML(rngTable)
        emailBodyVar = Replace(emailBodyVar, "<<EMAIL TABLE>>", RangetoHTML(rngTable))
            'Open ThisWorkbook.Path & "\email body template.htm" For Output As #1
            'Print #1, emailBodyVar
            'Close #1
    
    ' -----------------------------------------
    
    ' create email
    e = SendEmail(SENDorDISPLAYemailVar, subcontractor, to_Field, emailSubjectVar, emailBodyVar, emailCCVar, emailAttachment1, emailAttachment2)
    
    ' -----------------------------------------
    
    
    Exit Sub
errHandle:
    AddLog ("Error: " & Err.Number & vbNewLine & Err.Description)
    e = MsgBox("Error: " & Err.Number & vbNewLine & Err.Description, vbExclamation)
End Sub
Sub Mail_Selection_Range_Outlook_Body()
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Don't forget to copy the function RangetoHTML in the module.
'Working in Excel 2000-2016
    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object

    Set rng = Nothing
    On Error Resume Next
    'Only the visible cells in the selection
    Set rng = Selection.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    'Set rng = Sheets("YourSheet").Range("D4:D12").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = "ron@debruin.nl"
        .CC = ""
        .BCC = ""
        .Subject = "This is the Subject line"
        .HTMLBody = RangetoHTML(rng)
        .Display   'or use .Send
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        '.Show
        On Error GoTo 0
    End With
    
    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
End Function

Function SendEmail(SENDorDISPLAYemail As String, subconName As Variant, subconEmail As Variant, emailSubject As String, emailBody As String, Optional ByVal ccEmail As String = "", Optional ByVal attachmentPath As String = "", Optional ByVal secondAttachmentPath As String = "")
    ' Requires Microsoft Outlook Object Library
    ' Based on http://www.exceltrainingvideos.com/automate-excel-to-pdf-and-email-pdf-document-using-vba/
    ' Camron 2019-03-12
    
    'Dim subconName As String
    'subconName = "Westland Construction"
    'Dim subconEmail As String
    'subconEmail = "camron@westlandconstruction.com"
    'Dim emailSubject As String
    'emailSubject = "Herriman High 2 Closeouts Test"
    'Dim emailBody As String
    'emailBody = "Please send me your closeouts ASAP"
    'Dim SENDorDISPLAYemail As String
    'SENDorDISPLAYemail = "SEND"
       
    Dim OutLookApp As Object
    Dim OutLookMailItem As Object
    Dim myAttachments As Object
    
    Set OutLookApp = CreateObject("Outlook.application")
    Set OutLookMailItem = OutLookApp.CreateItem(0)
    Set myAttachments = OutLookMailItem.Attachments
    
    Select Case SENDorDISPLAYemail
        Case Is = "DISPLAY"
            With OutLookMailItem
                .Importance = 2
                .To = subconEmail
                .CC = ccEmail
                .Subject = emailSubject
                .HTMLBody = emailBody
                
                If attachmentPath <> "" Then myAttachments.Add attachmentPath
                If secondAttachmentPath <> "" Then myAttachments.Add secondAttachmentPath
                .Display
            End With
        Case Is = "SEND"
            With OutLookMailItem
                .Importance = 2
                .To = subconEmail
                .CC = ccEmail
                .Subject = emailSubject
                .HTMLBody = emailBody
                If attachmentPath <> "" Then myAttachments.Add attachmentPath
                If secondAttachmentPath <> "" Then myAttachments.Add secondAttachmentPath
                .Send
            End With
        Case Else
        MsgBox ("SENDorDISPLAYemail is a required argument")
    End Select
    
    Set OutLookMailItem = Nothing
    Set OutLookApp = Nothing
End Function

Function GetBoiler(ByVal sFile As String) As String
    'Dick Kusleika https://www.rondebruin.nl/win/s1/outlook/signature.htm
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.ReadAll
    ts.Close
End Function


