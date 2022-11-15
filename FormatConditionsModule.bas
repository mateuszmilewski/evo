Attribute VB_Name = "FormatConditionsModule"
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

Public Sub addFormatConditionsForReceptionReport()
    'addCondFrmt1 "T:T", "NOK", RGB(255, 0, 0)
    'addCondFrmt1 "W:W", "NOK", RGB(255, 0, 0)
    'addCondFrmt1 "W:W", "NO TANGO PRICE", RGB(255, 128, 0)
    
    addCondFrmt3939 "K:K", "3939"
    addCondFrmt1 "S:S", "NOK", RGB(255, 0, 0)
    addCondFrmt1 "S:S", "NO TANGO PRICE", RGB(255, 128, 0)
    addCondFrmt1 "S:S", "NO TANGO", RGB(255, 128, 0)
    addCondFrmt1 "S:S", "TP04 PRICE", RGB(240, 40, 0)
End Sub

Public Sub addFormatConditionsForGreenLightReport()
Attribute addFormatConditionsForGreenLightReport.VB_Description = "add a few dynamic colors"
Attribute addFormatConditionsForGreenLightReport.VB_ProcData.VB_Invoke_Func = "F\n14"

    addCondFrmt3939 "H:H", "3939"
    addCondFrmt1 "T:T", "NOK", RGB(255, 0, 0)
    addCondFrmt1 "W:W", "NOK", RGB(255, 0, 0)
    addCondFrmt1 "W:W", "NO TANGO PRICE", RGB(255, 128, 0)
End Sub

Private Sub addCondFrmt1(columnsStr As String, str As String, myColor As Long)
Attribute addCondFrmt1.VB_ProcData.VB_Invoke_Func = " \n14"

'
    columns(columnsStr).Select
    ' Application.CutCopyMode = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""" & str & """"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = myColor
        .TintAndShade = 0
    End With
    ' Selection.FormatConditions(1).StopIfTrue = False
End Sub


Private Sub addCondFrmt3939(columnStr As String, begStr As String)

'
    columns(columnStr).Select
    Selection.FormatConditions.Add Type:=xlTextString, String:=begStr, _
        TextOperator:=xlBeginsWith
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    ' Selection.FormatConditions(1).StopIfTrue = False
End Sub
