Attribute VB_Name = "Module1"
Sub removeDuplicates2()
Attribute removeDuplicates2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' removeDuplicates2 Macro
'

'
    Selection.End(xlToRight).Select
    Range("AN1").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("F8").Select
    Sheets(Array("Sheet33", "SrcPivot_20200408_")).Select
    Sheets("SrcPivot_20200408_").Activate
    ActiveWindow.SelectedSheets.Delete
    ActiveWorkbook.Save
    Selection.End(xlToRight).Select
    Range("AD12").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    Sheets("SrcPivot_20200408_").Select
    Range("H1").Select
    Selection.End(xlToRight).Select
    Range("AM13").Select
    ActiveWindow.SmallScroll Down:=22
    Range("AM35").Select
    ActiveWindow.SmallScroll Down:=-217
    ActiveWindow.SmallScroll ToRight:=4
    Application.WindowState = xlNormal
    Sheets("Proxy2_20200408_I").Select
    Sheets("Proxy2_20200408_I").move Before:=Sheets(1)
    Sheets("Proxy2_20200408_").Select
    Sheets("Proxy2_20200408_").move Before:=Sheets(4)
    Sheets(Array("Proxy2_20200408_", "SrcPivot_20200408_")).Select
    Sheets("SrcPivot_20200408_").Activate
    ActiveWindow.SelectedSheets.Delete
    Sheets("Proxy2_20200408_I").Select
    Selection.End(xlToRight).Select
    Range("AE12").Select
    Sheets("SrcPivot_20200408_I").Select
    ActiveWindow.SmallScroll ToRight:=25
    ActiveWindow.SmallScroll Down:=-84
    Range("AF1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    ActiveWindow.SmallScroll Down:=-8
    Range("AF1").Select
    ActiveWindow.SmallScroll Down:=17
    Range("AO33").Select
    ActiveCell.FormulaR1C1 = "=RC[-9]/RC[-3]"
    Range("AO33").Select
    ActiveCell.FormulaR1C1 = "=RC[-9]/RC[-3]"
    Range("AO33").Select
    ActiveCell.FormulaR1C1 = "=RC[-9]/RC[-3]*13.5"
    Range("AO33").Select
    Selection.ClearContents
    Sheets("Proxy2_20200408_I").Select
    ActiveWindow.SmallScroll Down:=-56
    Range("X1").Select
    Selection.AutoFilter
    ActiveWindow.SmallScroll ToRight:=16
    Range("AC13").Select
    ActiveWorkbook.Save
End Sub
