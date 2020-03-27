Attribute VB_Name = "CopyAndOptimiseDataModule"
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


Public Sub copyData(ictrl As IRibbonControl)

    ' this obsolete from version 006
    ' innerCopyData
    
    ' prepare form
    
    Dim w As Workbook
    
    With FileChooser
        .BtnCopy.Enabled = False
        .BtnValid.Enabled = True
        
        .ComboBoxFeed.Clear
        .ComboBoxMaster.Clear
        
        
        For Each w In Application.Workbooks
            .ComboBoxFeed.AddItem w.Name
            .ComboBoxMaster.AddItem w.Name
        Next w
        
        
        .show
    End With
    
    MsgBox "GOTOWE!", vbInformation
End Sub

Public Sub innerCopyData(masterFileName, feedFileName, Optional sh As StatusHandler)


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    
    If IsMissing(sh) Or (sh Is Nothing) Then
        Set sh = New StatusHandler
    End If
    
    ' master and feed worksheets
    Dim m As Worksheet, f As Worksheet
    ' starting from most impotant sheets!
    Set m = Workbooks(masterFileName).Sheets("BASE")
    Set f = Workbooks(feedFileName).Sheets("BASE CPL")
    


    Dim copy_h As CopyHandler
    Set copy_h = New CopyHandler
    copy_h.init m, f
    
    copy_h.workWithData sh
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    
    
    

End Sub




Public Sub optimiseDatesByTMC(ictrl As IRibbonControl)

    innerOptimiseData
    
    MsgBox "GOTOWE!", vbInformation
End Sub

Public Sub innerOptimiseData(Optional sh As StatusHandler)


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    
    If IsMissing(sh) Or (sh Is Nothing) Then
        Set sh = New StatusHandler
    End If
    
    Dim m As Worksheet, f As Worksheet
    'ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M1").Value
    Dim masterFileName As String
    Dim feedFileName As String
    masterFileName = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M1").Value
    feedFileName = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M2").Value
    Set m = Workbooks(masterFileName).Sheets("BASE")
    Set f = Workbooks(feedFileName).Sheets("BASE CPL")

    Dim copy_h As CopyHandler
    Set copy_h = New CopyHandler
    copy_h.init m, f
    
    copy_h.optimise sh
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    
    
    

End Sub
