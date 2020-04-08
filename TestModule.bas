Attribute VB_Name = "TestModule"
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

' testing overall status
Private Sub testMain()


    Dim sh As StatusHandler, x As Integer
    Set sh = New StatusHandler
    sh.init_statusbar 10
    sh.show
    For x = 1 To 10
        Sleep 1000
        sh.progress_increase
    Next x
    
    sh.hide
    
    Set sh = Nothing
    
End Sub



' PIVOT tests
Private Sub testOnPivot()
    
    Dim p As PivotHandler
    Set p = New PivotHandler
    p.initPivotSource
    p.initPivotSheet
    Set p = Nothing

End Sub


Public Sub testForArr()
    
    Dim txt As String
    txt = " fefh iefj e"
    
    Dim arr As Variant
    
    arr = Split(txt, " ")
    Dim x As Integer
    For x = LBound(arr) To UBound(arr)
        Debug.Print arr(x)
    Next x
End Sub




Sub kopiaSourcePivot()
'
' kopiaSourcePivot Macro
'

'
    Sheets("SrcPivot_20200408_II").Select
    Sheets("SrcPivot_20200408_II").Copy Before:=Sheets(1)
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AM$4214").RemoveDuplicates columns:=1, Header:= _
        xlYes

    columns("U:X").Select
    Selection.Delete Shift:=xlToLeft
    columns("U:U").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    columns("R:S").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
End Sub


