Attribute VB_Name = "DeleteModule"
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


Public Sub deleteThisSheet(ictrl As IRibbonControl)

    Dim sh As Worksheet
    Set sh = ThisWorkbook.ActiveSheet
    
    
    If sh.name <> "register" Then
        
        If sh.name = "forValidation" Then
            MsgBox "You can not remove forValidation sheet!", vbExclamation
        Else
            Application.DisplayAlerts = False
            sh.Delete
            Application.DisplayAlerts = True
        End If
    Else
        MsgBox "Critical! You can not remove register sheet!", vbExclamation
    End If
End Sub


Public Sub deleteAllDataSheets(ictrl As IRibbonControl)


    Dim ret As Variant
    Dim x As Integer
    
    ret = MsgBox("Delete all?", vbQuestion + vbYesNo)
    If ret = vbYes Then
        Application.DisplayAlerts = False
        
        
        Dim sh As Worksheet
        x = 1
        Do
            If checkIfYouCanDelete(Sheets(x)) Then
                x = x + 1
            Else
                Sheets(x).Delete
            End If
        Loop Until x > Sheets.Count
        Application.DisplayAlerts = True
    End If
End Sub

Private Function checkIfYouCanDelete(sh As Worksheet) As Boolean
    checkIfYouCanDelete = True
    
    If sh.name = "forValidation" Then
        Exit Function
    End If
    
    If sh.name = "register" Then
        Exit Function
    End If
    
    
    checkIfYouCanDelete = False
End Function
