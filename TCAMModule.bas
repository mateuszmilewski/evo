Attribute VB_Name = "TCAMModule"
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



Public Sub showDetails()


Public Sub createSourceForPivot(ictrl As IRibbonControl)

    
    Dim w As Workbook
    
    With FileChooser
        .scenarioType = E_FORM_SCENARIO_CREATE_PIVOT_SCENARIO
        .BtnCopy.Enabled = False
        .BtnValid.Enabled = True
        
        .ComboBoxFeed.Clear
        .ComboBoxMaster.Clear
        
        
        For Each w In Application.Workbooks
            .ComboBoxFeed.AddItem w.Name
            .ComboBoxMaster.AddItem w.Name
        Next w
        
        .ComboBoxFeed.Value = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("m2").Value
        .ComboBoxMaster.Value = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("m1").Value
        
        .show
    End With
    
    MsgBox "GOTOWE!", vbInformation
    
End Sub

Public Sub innerCreateSourceForPivot(masterFileName, feedFileName, Optional sh As StatusHandler)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    
    If IsMissing(sh) Or (sh Is Nothing) Then
        Set sh = New StatusHandler
    End If
    
    ' master and feed worksheets
    Dim m As Worksheet, f As Worksheet
    ' starting from most impotant sheets!
    On Error Resume Next
    Set m = Workbooks(masterFileName).Sheets(MAIN_SH_BASE)
    On Error Resume Next
    Set f = Workbooks(feedFileName).Sheets(G_FEED_SH_MAIN)

    Dim ch As CopyHandler, ph As PivotHandler
    Set ch = New CopyHandler
    Set ph = New PivotHandler
    
    ch.init m, f, E_COPY_HANDLER_FOR_PIVOT_CREATION
    
    ' MsgBox "implementation under way!", vbInformation
    ch.copyForSourcePivot ph, sh
    
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Public Sub makePivot(ictrl As IRibbonControl)
    
    innerMakePivot
    
    MsgBox "READY!"
End Sub

Public Sub makeTPivot(ictrl As IRibbonControl)
    
    innerMakeTPivot
    
    MsgBox "READY!"
End Sub


Private Sub innerMakePivot()


    If checkActiveSheetIfItIsProxy2() Then
    
        Dim ph As PivotHandler, oh As OperationsHandler
        Set ph = New PivotHandler
        Set oh = New OperationsHandler
        
        Set ph.proxy2 = ActiveSheet
        ph.initPivotSheet
        oh.makePivot ph
        
    Else
        MsgBox "Proxy2 sheet need to be active to perform action!", vbInformation
    End If
    
End Sub

Private Sub innerMakeTPivot()


    If checkActiveSheetIfItIsProxy2() Then
    
        Dim ph As PivotHandler, oh As OperationsHandler
        Set ph = New PivotHandler
        Set oh = New OperationsHandler
        
        Set ph.proxy2 = ActiveSheet
        ph.initPivotSheet
        oh.makeTPivot ph
        
    Else
        MsgBox "Proxy2 sheet need to be active to perform action!", vbInformation
    End If
    
End Sub

Private Function checkActiveSheetIfItIsProxy2() As Boolean
    checkActiveSheetIfItIsProxy2 = False
    
    If ActiveSheet.Name Like "Proxy2_*" Then
        If ActiveSheet.Cells(1, 1).Value = "ID" Then
        
            If ActiveSheet.Cells(1, 2).Value = "WIERSZ" Then
            
                If ActiveSheet.Cells(1, 3).Value = "REF" Then
                
                
                    checkActiveSheetIfItIsProxy2 = True
                    
                    Dim ans As Variant
                    ans = MsgBox("Do you want to create PIVOT for: " & CStr(ActiveSheet.Name) & " ? ", vbYesNo + vbQuestion)
                        
                    If ans = vbYes Then
                        checkActiveSheetIfItIsProxy2 = True
                    Else
                        checkActiveSheetIfItIsProxy2 = False
                        MsgBox "Nothing to do!"
                    End If
                End If
            End If
        End If
    End If

End Function
