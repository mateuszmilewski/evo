Attribute VB_Name = "ComboReceptionModule"
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


' binding helpers subs - for coordination of steps

Public Sub receptionOneClickCombo(ictrl As IRibbonControl)
    showComboFormForReceptionReport
End Sub

Public Sub showComboFormForReceptionReport()

    Dim sh As Worksheet, c As Control

    ' check interrocom availability inside evo already
    ' +
    ' check internal suppliers data inside evo already
    With ComboFormReceptionReport
    
        .ComboBoxTangoSource.Clear
        .ComboBoxInternalSupplier.Clear
        .ComboBoxManagersDA.Clear
    
        For Each sh In ThisWorkbook.Sheets
            If sh.name Like "INTERROCOM_*" Then
                .ComboBoxTangoSource.addItem sh.name
                .ComboBoxTangoSource.Value = sh.name
            End If
            
            If sh.name Like "N_*" Then
                .ComboBoxInternalSupplier.addItem sh.name
                .ComboBoxInternalSupplier.Value = sh.name
            End If
            
            If sh.name Like "MANAGERS_DA_*" Then
                .ComboBoxManagersDA.addItem sh.name
                .ComboBoxManagersDA.Value = sh.name
            End If
        Next sh
        
    End With
    
    
    

    
    With ComboFormReceptionReport
    
        'scope
        Dim str_Y_CW As String
        ' a little bit over-complicated
        ' for CW I temp add 100, becuase when we have CW7 I want to have CW07 - so + 100 having 107 and then right("107", 2)
        ' so every time just taking last two digits and making string from it.
        str_Y_CW = CStr(CLng(Year(Date))) & " CW" & Right(CStr(100 + CLng(Application.WorksheetFunction.IsoWeekNum(CDbl(CDate(Date - 7))))), 2)
        
        .TextBoxYYYYCW.Value = str_Y_CW
    
        .TextBoxAu01.Value = Format(Date, "dd.mm.yyyy")
        .TextBoxDu01.Value = Format(CDate(Date - 30), "dd.mm.yyyy")
        .TextBoxMvt1_01.Value = "101"
        .TextBoxMvt2_01.Value = "102"
        
        Set c = Nothing
        On Error Resume Next
        Set c = .Controls("TextBoxMag02")
        
        If c Is Nothing Then .innerAddLine
        
        
    End With
    
    With ComboFormReceptionReport
        ' predef
        
        .ComboBoxPRE_DEF.Clear
        
        Dim rr As Range
        Set rr = ThisWorkbook.Sheets("register").Range("AD2")
        Do
            If rr.Value = "R" Then
                .ComboBoxPRE_DEF.addItem rr.Offset(0, 1).Value
            End If
            Set rr = rr.Offset(1, 0)
        Loop Until Trim(rr.Value) = ""
    End With
    
    
    With ComboFormReceptionReport
        .show
    End With
    
    
    
End Sub
