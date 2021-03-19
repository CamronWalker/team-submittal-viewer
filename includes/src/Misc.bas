Attribute VB_Name = "Misc"
Sub RestoreSettingsToDefaults()
    Range("Dev_Mode").Value = False
    Range("Logging").Value = True
    Range("Custom_File_Location").Value = False
    Range("SENDorDISPLAYemail").Value = "DISPLAY"
    Range("Email_Table_Filter").Value = "<>Closeout"
    Range("Email_Hide_Closed").Value = "SHOW"

End Sub

Sub Filter_MonthlyReport()
    Dim filterStart As Long, filterEnd As Long
    filterStart = Worksheets("Monthly Report").Range("MonthlyReport_Filter_Start").Value
    filterEnd = Worksheets("Monthly Report").Range("MonthlyReport_Filter_End").Value
    
    Worksheets("Monthly Report Table").ListObjects("MonthlyReport_Table").Range.AutoFilter Field:=3, _
        Criteria1:="<=" & filterEnd
        
    Worksheets("Monthly Report Table").ListObjects("MonthlyReport_Table").Range.AutoFilter Field:=4, _
        Criteria1:=">=" & filterEnd, _
        Operator:=xlOr, _
        Criteria2:="=" & ""
        
    Worksheets("Monthly Report Table").ListObjects("MonthlyReport_Table").Range.AutoFilter Field:=5, _
        Criteria1:=">=" & filterEnd, _
        Operator:=xlOr, _
        Criteria2:="=" & ""

    Worksheets("Monthly Report").Pictures("LinkedImage_MonthlyReport").Formula = "='Monthly Report Table'!$A$1:$A$" & Worksheets("Monthly Report").Range("B13").Value
    
End Sub

Sub Copy_Filtered_Monthly_Report()
    


End Sub

Function Submittal_Review_Time(designDate, reviewedDate, closedDate, Optional ByVal filterDateStart, Optional ByVal filterDateEnd)
    
    'Dim designDate, reviewedDate, closedDate
    'designDate = Range("C16").Value
    'reviewedDate = Range("D16").Value
    'closedDate = Range("E16").Value
    
    If IsMissing(filterDateEnd) Then filterDateEnd = Date
    
    If designDate = "" Then
        Submittal_Review_Time = ""
        Exit Function
    End If
    
    If reviewedDate = "" And closedDate <> "" Then
        reviewedDate = closedDate
    End If
    
    If reviewedDate = "" And closedDate = "" Then
        reviewedDate = filterDateEnd
    End If
    
    If reviewedDate - designDate < 0 Then
        reviewedDate = filterDateEnd
    End If
    
    If designDate > filterDateEnd Then
        Submittal_Review_Time = ""
        Exit Function
    End If
    
    Submittal_Review_Time = reviewedDate - designDate
End Function


Sub MakeImageLinkedPicture()

    Dim startDate As Long: startDate = Worksheets("Monthly Report").Range("MonthlyReport_Filter_Start").Value
    Dim endDate As Long: Worksheets("Monthly Report").Range("MonthlyReport_Filter_End").Value
    


Dim ws As Worksheet

Set ws = ActiveSheet

ws.Pictures("LinkedImage_MonthlyReport").Formula = "='Monthly Report Table'!$A$1:$A$629"

End Sub
