Attribute VB_Name = "TP04Module"
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

Public Sub tp04Match(ictrl As IRibbonControl)


    PriceMatchForm.show
    
End Sub


Public Sub innerAfterSQ01Logic(masterFileName As String, feedFileName As String, Optional sh As StatusHandler)


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    
    If IsMissing(sh) Or (sh Is Nothing) Then
        Set sh = New StatusHandler
    End If
    
    ' master and feed worksheets
    Dim m As Worksheet, f As Worksheet
    
    ' safe init
    ' ----------------
    Set m = Nothing
    Set f = Nothing
    ' ----------------
    
    ' starting from most impotant sheets!
    On Error Resume Next
    Set m = Workbooks(masterFileName).Sheets(MAIN_SH_BASE)
    On Error Resume Next
    Set f = ThisWorkbook.Sheets(feedFileName)
    
    
    ' ====================================================================
    
    Dim instance_of_tp04 As TP04
    Set instance_of_tp04 = New TP04
    
    Dim ans As Variant, byUnit As Boolean
    ' ans = MsgBox("Choose YES if you want to see price by UNIT", vbInformation + vbYesNo)
    ' by default
    ans = vbNo ' so we calc later by packaging size!
    
    If ans = vbYes Then
        byUnit = True
    Else
        byUnit = False
    End If
    
    With instance_of_tp04
        .setStatusHandler sh
        .init m, f, byUnit
    End With
    ' ====================================================================
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Public Sub innerAfterTP04Logic(masterFileName As String, feedFileName As String, Optional sh As StatusHandler)


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    
    If IsMissing(sh) Or (sh Is Nothing) Then
        Set sh = New StatusHandler
    End If
    
    ' master and feed worksheets
    Dim m As Worksheet, f As Worksheet
    
    ' safe init
    ' ----------------
    Set m = Nothing
    Set f = Nothing
    ' ----------------
    
    ' starting from most impotant sheets!
    On Error Resume Next
    Set m = Workbooks(masterFileName).Sheets(MAIN_SH_BASE)
    On Error Resume Next
    Set f = Workbooks(feedFileName).Sheets(G_TP04_TP04_01)
    
    If f Is Nothing Then
        On Error Resume Next
        Set f = Workbooks(feedFileName).ActiveSheet
    End If
    
    
    
    ' ====================================================================
    
    Dim instance_of_tp04 As TP04
    Set instance_of_tp04 = New TP04
    
    
    Dim ans As Variant, byUnit As Boolean
    ans = MsgBox("Choose YES if you want to see price by UNIT", vbInformation + vbYesNo)
    If ans = vbYes Then
        byUnit = True
    Else
        byUnit = False
    End If
    
    With instance_of_tp04
        .setStatusHandler sh
        .init m, f, byUnit
    End With
    ' ====================================================================
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub
