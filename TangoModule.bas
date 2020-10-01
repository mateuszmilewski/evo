Attribute VB_Name = "TangoModule"
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



Public Sub matchSq01WithTango(ictrl As IRibbonControl)
    Debug.Print "matchSq01WithTango"
    
    fillForm
End Sub


Public Sub matchMb51WithTango(ictrl As IRibbonControl)
    Debug.Print "matchMb51WithTango"
    
    fillForm
End Sub

Private Sub fillForm()
    
    FindTangoOrIntSData.ComboBox1.Clear
    FindTangoOrIntSData.Caption = "Match with TANGO"
    
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.name Like "INTERROCOM_*" Then
            FindTangoOrIntSData.ComboBox1.addItem sh.name
        End If
    Next sh
    
    FindTangoOrIntSData.show
End Sub


Public Sub runMatchingLogicOnTango(sh As Worksheet, interrocomData As Worksheet, Optional auto1 As Boolean, Optional divStr As String)
    Debug.Print "main matching with tango logic!"
    
    
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    
    
    Dim divPattern As String
    divPattern = ""
    
    
    If auto1 Then
        If divStr <> "" Then
            divPattern = divStr
        End If
    Else
        On Error Resume Next
        divPattern = InputBox("Please put DIV pattern for price recogniction!", "Please put DIV pattern for price recogniction!")
    End If
    
    
    
    Dim rng As Range, bottomRng As Range, area As Range
    
    Dim e As E_MB51_AUTO_DECISION_LAYOUT
    
    
    If sh.name Like "TP04*" Then
        Set rng = sh.Range("A2").Offset(0, EVO.E_ADJUSTED_SQ01_Reference - 1)
    Else
        Set rng = sh.Range("A2")
        
        If sh.Range("A1").Value = "Article" Then
            e = E_MB51_AUTO_DECISION_LAYOUT_0
        Else
            e = E_MB51_AUTO_DECISION_LAYOUT_NEW
        End If
    End If
    
    Set bottomRng = rng.End(xlDown)
    Set area = sh.Range(rng, bottomRng)
    
    Dim status_h As New StatusHandler
    status_h.init_statusbar (area.Count / 10)
    
    status_h.show
    
    
    ' and N data will be all the time the same
    
    Dim irng As Range, bottom_irng As Range, i_area As Range
    Set irng = interrocomData.Range("A2").Offset(0, EVO.E_OUT_INTERROCOM_A - 1)
    Set bottom_irng = irng.End(xlDown)
    Set i_area = interrocomData.Range(irng, bottom_irng)
    
    If sh.name Like "TP04*" Then
    
        For Each rng In area
        
        
            rng.Offset(0, EVO.E_ADJUSTED_SQ01_IS_IN_TANGO - EVO.E_ADJUSTED_SQ01_Reference).Value = "NO TANGO"
        
            For Each irng In i_area
            
                If Trim(rng.Value) Like Trim(irng.Value) & "*" Then
                
                    If Trim(irng.Offset(0, EVO.E_OUT_INTERROCOM_REP - 1).Value) = "100" Then
                        
                        If Trim(irng.Offset(0, EVO.E_OUT_INTERROCOM_DIV - 1).Value) Like "*" & CStr(divPattern) & "*" Then
                
                            rng.Offset(0, EVO.E_ADJUSTED_SQ01_IS_IN_TANGO - EVO.E_ADJUSTED_SQ01_Reference).Value = ""
                            rng.Offset(0, EVO.E_ADJUSTED_SQ01_TANGO_PCS_PRICE - EVO.E_ADJUSTED_SQ01_Reference).Value = _
                                irng.Offset(0, EVO.E_OUT_INTERROCOM_FINAL_PRIX - EVO.E_OUT_INTERROCOM_A).Value
                        End If
                    End If
                End If
            Next irng
            
            
            If rng.row Mod 10 = 0 Then
                status_h.progress_increase
            End If
        Next rng
        
    Else
        
        
        For Each rng In area
        
        
            If e = E_MB51_AUTO_DECISION_LAYOUT_NEW Then
        
        
                rng.Offset(0, EVO.E_MB51_NEW_IS_IN_TANGO - EVO.E_MB51_NEW_MVT).Value = "NO TANGO"
            
                For Each irng In i_area
                
                    If Trim(rng.Offset(0, EVO.E_MB51_NEW_ARTICLE - EVO.E_MB51_NEW_MVT)) Like Trim(irng.Value) & "*" Then
                    
                    
                        If Trim(irng.Offset(0, EVO.E_OUT_INTERROCOM_REP - 1).Value) = "100" Then
                            
                            If Trim(irng.Offset(0, EVO.E_OUT_INTERROCOM_DIV - 1).Value) Like "*" & CStr(divPattern) & "*" Then
                    
                    
                                rng.Offset(0, EVO.E_MB51_NEW_IS_IN_TANGO - EVO.E_MB51_NEW_MVT).Value = ""
                                rng.Offset(0, EVO.E_MB51_NEW_TANGO_PCS_PRICE - EVO.E_MB51_NEW_MVT).Value = _
                                    irng.Offset(0, EVO.E_OUT_INTERROCOM_FINAL_PRIX - EVO.E_OUT_INTERROCOM_A).Value
                            End If
                        End If
                    End If
                Next irng
                
                
            ElseIf e = E_MB51_AUTO_DECISION_LAYOUT_0 Then
            
                rng.Offset(0, EVO.E_MB51_0_IS_IN_TANGO - 1).Value = "NO TANGO"
            
                For Each irng In i_area
                
                    If Trim(rng.Offset(0, EVO.E_MB51_0_ARTICLE - 1)) Like Trim(irng.Value) & "*" Then
                    
                    
                        If Trim(irng.Offset(0, EVO.E_OUT_INTERROCOM_REP - 1).Value) = "100" Then
                            
                            If Trim(irng.Offset(0, EVO.E_OUT_INTERROCOM_DIV - 1).Value) Like "*" & CStr(divPattern) & "*" Then
                    
                    
                                rng.Offset(0, EVO.E_MB51_0_IS_IN_TANGO - 1).Value = ""
                                rng.Offset(0, EVO.E_MB51_0_TANGO_PCS_PRICE - 1).Value = _
                                    irng.Offset(0, EVO.E_OUT_INTERROCOM_FINAL_PRIX - EVO.E_OUT_INTERROCOM_A).Value
                            End If
                        End If
                    End If
                Next irng
            
            
            Else
                MsgBox "Critical stop on tango matching with reception!", vbCritical
                End
            End If
            
            
            If rng.row Mod 10 = 0 Then
                status_h.progress_increase
            End If
        Next rng
        
    End If
    
    status_h.hide
    Set status_h = Nothing
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    If auto1 Then
    Else
        MsgBox "ready!"
    End If
End Sub


