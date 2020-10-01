Attribute VB_Name = "InterrocomModule"
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


Public Function notValidInterrocomFile(strValue As String) As Boolean
    notValidInterrocomFile = True
    
    
    If Len(strValue) > 0 Then
    
        Dim validationReference As Range
        Set validationReference = ThisWorkbook.Sheets("forValidation").Range("D20")
        
        Dim potInterrSh As Worksheet, potInterrRng As Range
        Set potInterrSh = Workbooks(strValue).ActiveSheet
        Set potInterrRng = potInterrSh.Cells(1, 1)
        
        Dim tmpR As Range
        Set tmpR = validationReference
        Do
            If CStr(tmpR.Value) = CStr(potInterrRng.Value) Then
                notValidInterrocomFile = False
            Else
                notValidInterrocomFile = True
                Exit Function
            End If
            
            Set tmpR = tmpR.Offset(0, 1)
            Set potInterrRng = potInterrRng.Offset(0, 1)
        Loop Until Trim(tmpR.Value) = ""
    End If
End Function


Public Sub getDataFromInterrocom(ictrl As IRibbonControl)
    
    Debug.Print "get data from interrocom local file (export from tango) started!"
    
    startInterrocomForm
End Sub

Private Sub startInterrocomForm()

    
    With InterrocomForm
        .ComboBox1.Clear
        Dim w As Workbook
        For Each w In Workbooks
            .ComboBox1.addItem w.name
        Next w
        .show
    End With
End Sub


Public Sub ok_runInterrocomAdjustment(sh1 As Worksheet, Optional lean As Worksheet, Optional alias As String)



    Dim ih As New InterrocomHandler
    ih.setSh1 sh1
    ih.setLeanSheet lean
    ih.makeLoop
    ih.putCollectionIntoEvo True, alias
    
    
    ' huge side effect, but im too lazy to change this
    Set lean = ih.getLean()
End Sub
