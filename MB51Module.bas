Attribute VB_Name = "MB51Module"
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

Public Sub getDataFromMB51(ictrl As IRibbonControl)
    
    MB51Form.TextBoxDu01.Value = Format((Date - 14), "dd.mm.yyyy")
    MB51Form.TextBoxAu01.Value = Format(Date, "dd.mm.yyyy")
    MB51Form.TextBoxMvt1_01.Value = "101"
    MB51Form.TextBoxMvt2_01.Value = "102"
    MB51Form.show
End Sub




Public Function validMb51data(sh1 As Worksheet, validationRefRef As Range) As Boolean
    validMb51data = False
    
    Dim validationRef As Range, labelsRef As Range
    
    ' Set validationRef = ThisWorkbook.Sheets("forValidation").Range("D35")
    Set validationRef = validationRefRef
    
    Set labelsRef = ActiveSheet.Cells(1, 1)
    
    Do
    
        If validationRef.Value = labelsRef.Value Then
            validMb51data = True
        Else
            validMb51data = False
            Exit Do
        End If
    
        Set validationRef = validationRef.offset(0, 1)
        Set labelsRef = labelsRef.offset(0, 1)
    Loop Until Trim(labelsRef.Value) = ""
    
    
    
    
    
    If validMb51data Then
        Debug.Print "activesheet is in std!"
    End If
    
    
End Function
