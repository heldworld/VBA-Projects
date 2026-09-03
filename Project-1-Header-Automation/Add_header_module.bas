Attribute VB_Name = "Module1"
Sub Add_header()
Attribute Add_header.VB_Description = "This Macro can automatically create a header for the given data set."
Attribute Add_header.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Add_header Macro
' This Macro can automatically create a header for the given data set.
'

'
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "show_id"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "type"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "title"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "director"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "cast"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "country"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "date_added"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "release_year"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "rating"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "duration"
    Range("A1:J1").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Font.Bold = True
End Sub
