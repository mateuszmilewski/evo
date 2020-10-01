Attribute VB_Name = "SQ01ToTP04Module"
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


Public Sub adjustSq01Data(ictrl As IRibbonControl)
    independentRunOfAdjuster
End Sub

Public Sub independentRunOfAdjuster()
    runAdjusterForDataFromSq01 ActiveSheet
End Sub

Public Sub runAdjusterForDataFromSq01(osh As Worksheet)


    If checkIfLabelsInOutputAreInlineWithStd(osh) Then
    
        Debug.Print "initial output sq01 table is in std, go with the logic then!"
        innerSq01
    Else
        MsgBox "Wrong standard!", vbCritical
    End If

End Sub


Public Sub innerSq01()

    Dim w As Workbook, tmpCaption As String, sh1 As Worksheet
    
    With FileChooser
    
        tmpCaption = .LabelForSecFile.Caption
        .LabelForSecFile.Caption = "SQ01 data"
    
        .scenarioType = E_FORM_SCENATIO_PRICE_MATCHING_FOR_SQ01
        .BtnCopy.Enabled = False
        .BtnValid.Enabled = True
        
        .ComboBoxFeed.Clear
        .ComboBoxMaster.Clear
        
        
        For Each w In Application.Workbooks
            ' .ComboBoxFeed.AddItem w.name
            .ComboBoxMaster.addItem w.name
        Next w
        
        For Each sh1 In ThisWorkbook.Sheets
            .ComboBoxFeed.addItem sh1.name
        Next sh1
        
        .ComboBoxFeed.Value = ActiveSheet.name
        
        
        .show
    End With
End Sub



Private Function checkIfLabelsInOutputAreInlineWithStd(sh As Worksheet) As Boolean
    
    
    checkIfLabelsInOutputAreInlineWithStd = True
    

    Dim rng As Range, fvr As Range
    Set rng = sh.Cells(1, 1)
    Set fvr = ThisWorkbook.Sheets("forValidation").Range(G_REF_MOUNT_SQ1_OUT)
    
    Do
        If fvr.Value <> rng Then
            checkIfLabelsInOutputAreInlineWithStd = False
            Exit Do
        End If
        
        Set fvr = fvr.Offset(0, 1)
        Set rng = rng.Offset(0, 1)
        
        
    Loop Until CStr(fvr.Value) = ""
    
End Function
