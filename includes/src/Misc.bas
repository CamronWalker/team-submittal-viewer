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

    Worksheets("Monthly Report").Pictures("LinkedImage_MonthlyReport").Formula = "='Monthly Report Table'!$A$1:$A$" & Worksheets("Monthly Report").Range("B13").Value
    
End Sub

Sub Copy_Filtered_Monthly_Report()
    


End Sub

Sub MakeImageLinkedPicture()

Dim ws As Worksheet

Set ws = ActiveSheet

ws.Pictures("LinkedImage_MonthlyReport").Formula = "='Monthly Report Table'!$A$1:$A$629"

End Sub
