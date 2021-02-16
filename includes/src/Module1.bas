Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWorkbook.Worksheets("Email").ListObjects("Sub_List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Email").ListObjects("Sub_List").Sort.SortFields. _
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
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveWorkbook.Worksheets("Email").ListObjects("Sub_List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Email").ListObjects("Sub_List").Sort.SortFields. _
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
