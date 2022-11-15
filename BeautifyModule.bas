Attribute VB_Name = "BeautifyModule"
Option Explicit

Public Sub beautifyGreenLightList()
Attribute beautifyGreenLightList.VB_ProcData.VB_Invoke_Func = " \n14"
'
' beautifyGreenLightList Macro
'

'
    columns("A:A").Select
    Selection.Font.Size = 8
    Selection.ColumnWidth = 11
    columns("B:B").Select
    Selection.Font.Size = 8
    Selection.ColumnWidth = 27
    columns("C:C").Select
    Selection.Font.Size = 9
    columns("D:P").Select
    With Selection.Font
        .name = "Calibri"
        .Size = 9
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
    columns("D:D").Select
    Selection.EntireColumn.Hidden = True
    columns("F:G").Select
    Selection.EntireColumn.Hidden = True
    columns("I:P").Select
    Selection.EntireColumn.Hidden = True
    columns("Q:Q").ColumnWidth = 8
    columns("R:AE").Select
    Selection.ColumnWidth = 8
    columns("AD:AE").Select
    Selection.EntireColumn.Hidden = True
    columns("AF:AJ").Select
    Selection.ColumnWidth = 4
    With Selection.Font
        .name = "Calibri"
        .Size = 8
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
    columns("Y:Z").Select
    Selection.ColumnWidth = 10
    Selection.Font.Size = 9
    columns("X:X").Select
    Selection.Font.Size = 9
    columns("W:W").Select
    Selection.Font.Size = 9
    
    Range("A1").Select
    On Error Resume Next
    Selection.AutoFilter

    ' to be done!
    '
    '
    'ActiveWorkbook.Worksheets("GREEN_LIGHT_20201019_ p2qo").AutoFilter.Sort. _
    '    SortFields.Clear
    'ActiveWorkbook.Worksheets("GREEN_LIGHT_20201019_ p2qo").AutoFilter.Sort. _
    '    SortFields.Add key:=Range("AA1:AA2042"), SortOn:=xlSortOnValues, Order:= _
    '    xlDescending, DataOption:=xlSortNormal
    'With ActiveWorkbook.Worksheets("GREEN_LIGHT_20201019_ p2qo").AutoFilter.Sort
    '    .Header = xlYes
    '    .MatchCase = False
    '    .Orientation = xlTopToBottom
    '    .SortMethod = xlPinYin
    '    .Apply
    'End With
    'ActiveSheet.Range("$A$1:$AI$2042").AutoFilter Field:=29, Criteria1:="="
End Sub

Public Sub beautifyReceptionList()
Attribute beautifyReceptionList.VB_ProcData.VB_Invoke_Func = " \n14"
'
' beautifyReceptionList Macro
'

'
    columns("A:B").Select
    Range("B1").Activate
    With Selection.Font
        .name = "Calibri"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.ColumnWidth = 3.43
    columns("C:C").Select
    Selection.Font.Size = 9
    Selection.ColumnWidth = 30.57
    columns("E:G").Select
    Selection.EntireColumn.Hidden = True
    columns("I:I").Select
    Range("N1").Select
    
    
    columns("W:X").Select
    Range("B1").Activate
    With Selection.Font
        .name = "Calibri"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    Selection.AutoFilter
    
    'ActiveSheet.Range("$A$1:$V$190").AutoFilter Field:=20, Criteria1:="="
    'ActiveWorkbook.Worksheets("RECEPTION_20201026_").AutoFilter.Sort.SortFields. _
    '    Clear
    'ActiveWorkbook.Worksheets("RECEPTION_20201026_").AutoFilter.Sort.SortFields. _
    '    Add key:=Range("R1:R190"), SortOn:=xlSortOnValues, Order:=xlDescending, _
    '    DataOption:=xlSortNormal
    'With ActiveWorkbook.Worksheets("RECEPTION_20201026_").AutoFilter.Sort
    '    .Header = xlYes
    '    .MatchCase = False
    '    .Orientation = xlTopToBottom
    '    .SortMethod = xlPinYin
    '    .Apply
    'End With
End Sub
