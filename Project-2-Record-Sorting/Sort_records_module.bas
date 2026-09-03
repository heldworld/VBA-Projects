Attribute VB_Name = "Module1"
Sub release_year_sort()
Attribute release_year_sort.VB_Description = "Sort descending by release_year"
Attribute release_year_sort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' release_year_sort Macro
' Sort descending by release_year
'

'
    Range("H1").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("H2:H51") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1:J51")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Buttons.Add(702.75, 73.5, 144, 45).Select
    Selection.OnAction = "release_year_sort"
    Selection.Characters.Text = "Descending Release Year"
    With Selection.Characters(Start:=1, Length:=23).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("R9").Select
End Sub
Sub date_added_ascending()
Attribute date_added_ascending.VB_Description = "Sorts ascending from oldest to newest by date_added"
Attribute date_added_ascending.VB_ProcData.VB_Invoke_Func = " \n14"
'
' date_added_ascending Macro
' Sorts ascending from oldest to newest by date_added
'

'
    Range("G1").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("G2:G51") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1:J51")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("N11").Select
End Sub
