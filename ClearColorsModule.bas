Attribute VB_Name = "ClearColorsModule"
Option Explicit


'The MIT License (MIT)
'
'Copyright (c) 2020 FORREST
' Mateusz Milewski mateusz.milewski@mpsa.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
'
' THE EVO TOOL

' formatting module

Sub clearColors(ictrl As IRibbonControl)
    innerClearColors
End Sub


Sub innerClearColors()
Attribute innerClearColors.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TestForClearingColorsMacro Macro
'

'
    ' Windows("PICK_UP_SHEET_coforTest.xlsm").Activate
    Dim wrk As Workbook
    Set wrk = Nothing
    On Error Resume Next
    Set wrk = Workbooks(ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M1").Value)
    
    
    If Not wrk Is Nothing Then
    
        ' Range("A6").Select
        Dim rng As Range
        Dim lastRow As Long
        Set rng = wrk.Sheets("BASE").Range("a2")
        Set rng = rng.End(xlDown).End(xlDown).End(xlUp)
        lastRow = CLng(rng.row)
        
        Set rng = wrk.Sheets("BASE").Range("A3:AV" & CStr(lastRow))
        
        Set rng = rng.SpecialCells(xlCellTypeVisible)
        
        With rng.Font
            .ColorIndex = 1
            .TintAndShade = 0
        End With
        With rng.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Else
        MsgBox "No bind with PUS master workshhet!", vbCritical
    End If
End Sub



Sub removeAllComments(ictrl As IRibbonControl)
    innerRemoveAllComments
End Sub


Sub innerRemoveAllComments()
'
' TestForClearingColorsMacro Macro
'

'
    ' Windows("PICK_UP_SHEET_coforTest.xlsm").Activate
    Dim wrk As Workbook
    Set wrk = Nothing
    On Error Resume Next
    Set wrk = Workbooks(ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M1").Value)
    
    
    If Not wrk Is Nothing Then
    
        ' Range("A6").Select
        Dim rng As Range
        Dim lastRow As Long
        Set rng = wrk.Sheets("BASE").Range("a2")
        Set rng = rng.End(xlDown).End(xlDown).End(xlUp)
        lastRow = CLng(rng.row)
        
        
        Dim b As Worksheet
        Set b = wrk.Sheets("BASE")
        
        Dim c As Comment
        For Each c In b.Comments
            c.Delete
        Next c
    Else
        MsgBox "No bind with PUS master workshhet!", vbCritical
        
    End If
End Sub
