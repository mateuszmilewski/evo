Attribute VB_Name = "ComboGreenLightModule"
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

Public Sub greenLightOneClickCombo(ictrl As IRibbonControl)
    showComboFormForReceptionReport
End Sub

Public Sub showComboFormForReceptionReport()

    Dim sh As Worksheet
    
    Dim preDef As Range, preDefCurrFill As Range
    Set preDef = ThisWorkbook.Sheets("register").Range("AD2")
    
    With ComboFormGreenLightReport
        With .ComboBoxPRE_DEF
            .Clear
            
            Do
                If preDef.Value = "F" Then
                    .addItem preDef.offset(0, 1).Value
                End If
                Set preDef = preDef.offset(1, 0)
            Loop Until Trim(preDef) = ""
        End With
        
    End With
    
    
    ' fill already with pre-def
    ' PRE_DEF_RUN_FOR_SQ01
    Set preDefCurrFill = ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01")
    Dim refReg As Range, x As Variant
    Set refReg = preDefCurrFill
    With ComboFormGreenLightReport
        For x = 1 To 5
            On Error Resume Next
            .Controls("TextBox1" & CStr(x)).Value = CStr(refReg.offset(0, x - 1).Value)
            
            On Error Resume Next
            .Controls("TextBox2" & CStr(x)).Value = CStr(refReg.offset(1, x - 1).Value)
        Next x
    End With
    
    ' filling data from combo boxes
    With ComboFormGreenLightReport
    
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
            
            If sh.name Like "CONCAT_*" Then
                .ComboBoxTP04Ready.addItem sh.name
                .ComboBoxTP04Ready.Value = sh.name
            End If
        Next sh
        
    End With
    
    ' fill with opened workbook
    With ComboFormGreenLightReport
        Dim w As Workbook
        .ComboBoxPUS.Clear
        For Each w In Workbooks
            .ComboBoxPUS.addItem w.name
        Next w
    End With
    
    
    With ComboFormGreenLightReport
        .SubmitBtn.Enabled = False
        .show
    End With
    
End Sub
