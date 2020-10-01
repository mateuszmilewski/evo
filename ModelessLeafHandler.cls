VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelessLeafHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Private greenLightMode As Boolean
Private receptionMode As Boolean

Private frm1 As ModelessLeaf

Private currentColumnToFill As Integer


Public Sub setMode(gl As Boolean, rc As Boolean, mfrm1 As ModelessLeaf)
    
    greenLightMode = gl
    receptionMode = rc
    
    Set frm1 = mfrm1
End Sub

Private Sub Class_Initialize()
    greenLightMode = False
    receptionMode = False
    
    Set frm1 = Nothing
End Sub

Private Sub Class_Terminate()
    greenLightMode = False
    receptionMode = False
    
    Set frm1 = Nothing
End Sub


Public Sub createLeaf()
    
    If greenLightMode Then
        innerCreateLeafForGreenLight
    End If
    
    If receptionMode Then
        innerCreateLeafForReception
    End If
End Sub


Private Sub innerCreateLeafForGreenLight()
    ' ------------------------------------------
    
    ' MAIN LOGIC FOR GREEN LIGHT
    ' ------------------------------------------
    Dim r As Range
    Set r = makeGreenLightScaffold()
    
    ' ------------------------------------------
End Sub

Private Sub innerCreateLeafForReception()
    ' ------------------------------------------
    
    ' MAIN LOGIC FOR RECEPTION
    ' ------------------------------------------
    Dim r As Range
    Set r = makeReceptionScaffold()
    ' ------------------------------------------
End Sub


Private Function makeGreenLightScaffold() As Range
    Set makeGreenLightScaffold = Nothing
    
    Dim sh As Worksheet
    Set sh = justMakeWorkbookAndSheetInIt("GREEN_LIGHT")
    
    
    ' fill green light labels
    fillGreenLightLabels sh
    fillWithDataFromForm sh
End Function


Private Function makeReceptionScaffold() As Range
    Set makeReceptionScaffold = Nothing
    
    Dim sh As Worksheet
    Set sh = justMakeWorkbookAndSheetInIt("RECEPTION")
    
    
    ' fill reception labels
    fillReceptionLabels sh
    fillWithDataFromForm sh
End Function

Private Function justMakeWorkbookAndSheetInIt(shNm As String) As Worksheet

    Set justMakeWorkbookAndSheetInIt = Nothing
    
    Dim wrk As Workbook
    Dim sh As Worksheet
    
    If frm1.outWrk Is Nothing Then
        Set wrk = Application.Workbooks.Add
        Set sh = wrk.Sheets.Add
        sh.name = CStr(shNm)
        Set frm1.outWrk = wrk
        currentColumnToFill = 3
    Else
        Set wrk = frm1.outWrk
        
        Set sh = Nothing
        On Error Resume Next
        Set sh = wrk.Sheets(shNm)
        
        If sh Is Nothing Then
            Set wrk = Application.Workbooks.Add
            Set sh = wrk.Sheets.Add
            sh.name = CStr(shNm)
            
            currentColumnToFill = 3
        Else
            currentColumnToFill = sh.Cells(3, 2).End(xlToRight).Column + 1
        End If
    End If
    

    
    Set justMakeWorkbookAndSheetInIt = sh
    
    
End Function


Private Sub fillGreenLightLabels(sh1 As Worksheet)

    With sh1
        .Cells(3, 2).Value = "Mass € Internal part (out of perimeter price * 1)"
        .Cells(4, 2).Value = "Number of Internal parts"
        .Cells(5, 2).Value = "Mass € Forcast without Tango price"
        .Cells(6, 2).Value = "% mass €  without Tango price"
        .Cells(7, 2).Value = "number ref without Tango price"
        .Cells(8, 2).Value = "rate  : Mass Sigapp pre-serial price / initial price (without Tango price)"
        
        .Cells(10, 2).Value = "Mass € Sigapp forcast (With Tango price)"
        .Cells(11, 2).Value = "number ref with Tango price"
        .Cells(12, 2).Value = "Including number no price * 1"

        .Cells(14, 2).Value = "Mass € target Price * 1"
        .Cells(15, 2).Value = "rate  : Mass forcast Sigapp price /target price * 1 (with Tango price)"
        .Cells(16, 2).Value = "Mass € gap Sigapp price / price *1"
    End With

End Sub


Private Sub fillReceptionLabels(sh1 As Worksheet)

    With sh1
        .Cells(3, 2).Value = "Mass € Internal part (out of perimeter price * 1)"
        .Cells(4, 2).Value = "Number of Internal parts"
        .Cells(5, 2).Value = "Mass € Reception without Tango price"
        .Cells(6, 2).Value = "% mass €  without Tango price"
        .Cells(7, 2).Value = "number ref without Tango price"
        .Cells(8, 2).Value = "rate  : Mass Sigapp pre-serial price / initial price (without Tango price)"
        
        .Cells(10, 2).Value = "Mass € Sigapp forcast (With Tango price)"
        .Cells(11, 2).Value = "number ref with Tango price"
        .Cells(12, 2).Value = "Including number no price * 1"

        .Cells(14, 2).Value = "Mass € target Price * 1"
        .Cells(15, 2).Value = "rate  : Mass forcast Sigapp price /target price * 1 (with Tango price)"
        .Cells(16, 2).Value = "Mass € gap Sigapp price / price *1"
    End With
    
    sh1.Range("B1:B1").columns.AutoFit

End Sub

Private Sub fillWithDataFromForm(sh1 As Worksheet)
    
    Dim r As Range
    With sh1
        .Cells(2, currentColumnToFill).Value = frm1.ListBoxScope.Value
        .Cells(3, currentColumnToFill).Value = frm1.TextBox_CostInternal.Value
        .Cells(4, currentColumnToFill).Value = frm1.TextBox_CountInternal.Value
        .Cells(5, currentColumnToFill).Value = frm1.TextBox_CostNoTango.Value
        .Cells(6, currentColumnToFill).FormulaR1C1Local = "=R[-1]C/(R[-1]C + R[4]C)"
        .Cells(7, currentColumnToFill).Value = frm1.TextBox_CountNoTango.Value
        ' row 8 empty for now - rate for parts without tango GL
        .Cells(10, currentColumnToFill).Value = frm1.TextBox_CostTango.Value
        On Error Resume Next
        .Cells(11, currentColumnToFill).Value = CLng(frm1.TextBox_CountTango.Value) + CLng(frm1.TextBox_CountTangoNOK.Value)
        .Cells(12, currentColumnToFill).Value = frm1.TextBox_CountTangoNOK.Value
        
        .Cells(14, currentColumnToFill).Value = frm1.TextBox_CostTarget.Value
        .Cells(15, currentColumnToFill).FormulaR1C1Local = "=R[-5]C/R[-1]C"
        .Cells(16, currentColumnToFill).FormulaR1C1Local = "=R[-6]C - R[-2]C"
    End With
End Sub
