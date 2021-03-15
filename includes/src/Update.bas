Attribute VB_Name = "Update"
Sub GetCustomSubmittalsExportLocation()
    On Error GoTo errHandle
    
    Dim strFileToOpen As String
    Dim currentFormula As String, finalFormula As String
    
    currentFormula = ActiveWorkbook.Queries.Item("Submittal").Formula
    If Not currentFormula = Range("Query_Formula").Value Then
        AddLog ("Query_Formula was changed!!! Current: " & currentFormula & "     vs     Old Range Value:" & Range("Query_Formula").Value)
    End If
    
    strFileToOpen = Application.GetOpenFilename _
        (Title:="Please select the Submittal Export file you exported from Viewpoint Team.", _
        FileFilter:="Excel Files *.xls* (*.xls*),")
    
    finalFormula = Replace(currentFormula, Range("Submittal_Export_Path").Value, strFileToOpen)
    
    If Not strFileToOpen = "False" Then
        ActiveWorkbook.Queries.Item("Submittal").Formula = finalFormula
        
        Range("Submittal_Export_Path").Value = strFileToOpen
        Range("Query_Formula").Value = finalFormula
        Range("Custom_File_Location").Value = True
    End If
    
    Exit Sub
errHandle:
    AddLog ("Error: " & Err.Number & vbNewLine & Err.Description)
    e = MsgBox("Error: " & Err.Number & vbNewLine & Err.Description, vbExclamation)
End Sub

Sub SetCurrentFolderAsExportLocation()
'If Range("Custom_File_Location").Value = False Then SetCurrentFolderAsExportLocation
    On Error GoTo errHandle

    Dim currentFormula As String, finalFormula As String, currentPath As String
    
    currentPath = Application.ActiveWorkbook.Path & "\Submittals Export.xlsx"
    currentFormula = ActiveWorkbook.Queries.Item("Submittal").Formula
    
    If Not currentPath = Range("Submittal_Export_Path").Value Then
        AddLog ("File Moved!  Old Path:" & currentPath & "     vs     New Path:" & Range("Submittal_Export_Path").Value)
    End If
    
    Range("Submittal_Export_Path").Value = currentPath
    
    finalFormula = Replace(currentFormula, Range("Submittal_Export_Path").Value, currentPath)
    
    If Not strFileToOpen = "False" Then
        ActiveWorkbook.Queries.Item("Submittal").Formula = finalFormula
        
        Range("Query_Formula").Value = finalFormula
    End If

    Exit Sub
errHandle:
    AddLog ("Error: " & Err.Number & vbNewLine & Err.Description)
    e = MsgBox("Error: " & Err.Number & vbNewLine & Err.Description, vbExclamation)
End Sub

Sub RefreshSubmittalQuery()
    If Range("Custom_File_Location").Value = False Then SetCurrentFolderAsExportLocation
    
    'On Error GoTo errHandle
    ActiveWorkbook.Connections("Query - Submittal").Refresh
    Application.CalculateUntilAsyncQueriesDone
    
    ResizeEmailLogTable
    ResizeOACLogTable
    
    Exit Sub
errHandle:
    AddLog ("Error: " & Err.Number & vbNewLine & Err.Description)
    If Err.Number = 1004 Then e = MsgBox("ERROR: 1004" & vbNewLine & "This error is likely because the Submittals Export.xlsx file is missing or named incorrectly. Please either select the file in settings or put it back in the same filder of this excel file.", vbExclamation)
End Sub

Sub ResizeEmailLogTable()
    Worksheets("Email Table").Rows.EntireRow.Hidden = False
    queryRowMax = Application.WorksheetFunction.Max(Worksheets("Query").ListObjects("Submittal").ListColumns("Index").DataBodyRange) + 1
    Worksheets("Email Table").ListObjects("Email_Table").Resize Range("A1:G" & queryRowMax)
    Worksheets("Email Table").Rows(queryRowMax & ":" & Worksheets("Email Table").Rows.Count).Delete
        
End Sub

Sub ResizeOACLogTable() 'TODO
    Worksheets("OAC Log").Rows.EntireRow.Hidden = False
    Worksheets("OAC Log").ListObjects("OAC_Table").Sort.SortFields.Clear
    queryRowMax = Application.WorksheetFunction.Max(Worksheets("Query").ListObjects("Submittal").ListColumns("Index").DataBodyRange) + 16
    Worksheets("OAC Log").ListObjects("OAC_Table").Resize Range("A15:H" & queryRowMax)
    Worksheets("OAC Log").Rows(queryRowMax & ":" & Worksheets("OAC Log").Rows.Count).Delete

End Sub

Sub UpdateSubList()
    Dim submittalQuery As Variant
    Dim subList As Collection
    
    submittalQuery = Worksheets("Query").ListObjects("Submittal").DataBodyRange.Value

    ' Get Collection of Unique Subs
    Set subList = New Collection
    subArray = Worksheets("Query").ListObjects("Submittal").ListColumns("Submitter Organization").DataBodyRange.Value
    
    For rst = 1 To UBound(subArray)
        e = CollectionAddUnique(subList, CStr(subArray(rst, 1)))
    Next rst
    ' -----------------------------------------------
    
    Worksheets("Email").ListObjects("Sub_List").Resize Range("A1:C2")
    Range("A3:C" & Rows.Count).Delete Shift:=xlUp
    
    
    For suba = 1 To subList.Count
         Worksheets("Email").ListObjects("Sub_List").DataBodyRange(suba, 1).Value = subList(suba)
         Worksheets("Email").ListObjects("Sub_List").DataBodyRange(suba, 2).Value = "NO"
    Next suba
    
    Worksheets("Email").ListObjects("Sub_List").Sort.SortFields. _
        Clear
    Worksheets("Email").ListObjects("Sub_List").Sort.SortFields. _
        Add2 Key:=Range("Sub_List[[#All],[Subcontractor]]"), SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Email").ListObjects("Sub_List").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
'===================================================
' ADDS ONLY UNIQUE ITEMS TO A COLLECTION
'===================================================
Public Function CollectionAddUnique(ByRef Target As Collection, Value As Variant) As Boolean

    Dim l As Long

    'SEE IF COLLECTION HAS ANY VALUES
    If Target.Count = 0 Then
        Target.Add Value
        Exit Function
    End If

    'SEE IF VALUE EXISTS IN COLLECTION
    For l = 1 To Target.Count
        If Target(l) = Value Then
            Exit Function
        End If
    Next l

    'NOT IN COLLECTION, ADD VALUE TO COLLECTION
    If Not Value = "" Then
        Target.Add Value
        CollectionAddUnique = True
    End If

End Function
