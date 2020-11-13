VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatusHandler"
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
'
'
' THE EVO TOOL


Public szerokosc_paska As Double
' bardzo wazne zeby nie przekroczyc zakresu mozliwosci tych dwoch zmiennych
Public act As Double

' ile jest czesci w sumie do sprawdzenia
Public max As Double
Public dx As Double
Public jeden_procent As Double
' 360/255 <-  pod gradientowa zmiane koloru
Public grad_color As Double
Private increment_decrement_flag As Boolean

Private est_time As Double
Private avr_on_one_pn As Double
Private macro_starts As Double
Private time_flag As Double

Private collection_of_avr_on_pn As Collection

Public Sub simplifyStatusBar()
    
    StatusForm.LabelAverageTime.Caption = ""
    StatusForm.LabelEstimatedTime.Caption = ""
    
End Sub

Public Sub init_statusbar(m_max, Optional m_min)

    Dim x As Integer

    Set collection_of_avr_on_pn = New Collection

    max = m_max
    If Not IsMissing(m_min) Then
        act = m_min
    End If
    dx = CDbl(szerokosc_paska) / CDbl(max)
    
    macro_starts = CDbl(Now)
    avr_on_one_pn = CDbl(INITIAL_TIMING_FOR_ONE_PN)
    
    For x = 1 To max
        collection_of_avr_on_pn.Add CDbl(INITIAL_TIMING_FOR_ONE_PN)
    Next x
    
    time_flag = macro_starts
    est_time = CDbl(max) * CDbl(avr_on_one_pn)
    
    
    Dim mitem As Variant
    est_time = 0
    For Each mitem In collection_of_avr_on_pn
        est_time = est_time + CDbl(mitem)
    Next mitem
    
    StatusForm.LabelAverageTime.Caption = CStr(CLng(avr_on_one_pn)) & " s"
    StatusForm.LabelEstimatedTime.Caption = CStr(convert_and_show(CLng(est_time)))
    
End Sub

Public Sub progress_increase()
    If act < max Then
        act = act + 1
        
        
        StatusForm.PasekPostepu.Width = StatusForm.PasekPostepu.Width + dx
        If StatusForm.PasekPostepu.Width >= 360 Then
            StatusForm.PasekPostepu.Width = 360
        End If
        DoEvents
        StatusForm.FrameProgress.Caption = CStr(Int((CDbl(act) / CDbl(max)) * 100)) & " %"
        
        
        time_flag = CDbl(Now)
        avr_on_one_pn = CDbl(time_flag - macro_starts) ' co by dostac format sekundowy
        avr_on_one_pn = avr_on_one_pn * 24
        avr_on_one_pn = avr_on_one_pn * 60
        avr_on_one_pn = avr_on_one_pn * 60
        
        If collection_of_avr_on_pn.count > 1 Then
            collection_of_avr_on_pn.Remove 1
            collection_of_avr_on_pn.Remove 1
        End If
        ' ' Debug.Print collection_of_avr_on_pn.COUNT
        collection_of_avr_on_pn.Add CLng(avr_on_one_pn)
        ' ' Debug.Print collection_of_avr_on_pn.COUNT
        
        
        
        ' to bedzie taki warunek brzegowy czy w ogole zajmowac sie
        'dokladnym przeliczaniem makra
        If collection_of_avr_on_pn.count > 10 Then
            Dim x As Integer
            For x = act To collection_of_avr_on_pn.count
                '' Debug.Print collection_of_avr_on_pn.COUNT
                collection_of_avr_on_pn.Remove x
                '' Debug.Print collection_of_avr_on_pn.COUNT
                collection_of_avr_on_pn.Add CLng(avr_on_one_pn)
                '' Debug.Print collection_of_avr_on_pn.COUNT
            Next x
        End If

        est_time = 0
        Dim mitem As Variant
        For Each mitem In collection_of_avr_on_pn
            est_time = est_time + CDbl(mitem)
        Next mitem
        
        
        StatusForm.LabelAverageTime.Caption = CStr(CLng(avr_on_one_pn)) & " s"
        StatusForm.LabelEstimatedTime.Caption = CStr(convert_and_show(CLng(est_time)))
        
        macro_starts = time_flag
        
        StatusForm.Repaint
    Else
        act = 0
        StatusForm.PasekPostepu.Width = 0
    End If
End Sub

Public Function convert_and_show(et As Long) As String
    convert_and_show = " 00h:00m:00s"
    
    Dim godziny As Integer
    Dim minuty As Integer
    Dim pozostalosc_po_minutach As Integer
    
    godziny = Int(et / 3600)
    minuty = Int(et / 60)
    pozostalosc_po_minutach = Int(et Mod 60)
    
    convert_and_show = CStr(godziny) & "h:" & CStr(minuty) & "m:" & str(pozostalosc_po_minutach) & "s"
End Function

Public Sub show()
    StatusForm.PasekPostepu.BackColor = RGB(254, 190, 16)
    StatusForm.TloPaskaPostepu.BackColor = RGB(0, 0, 0)
    StatusForm.BtnPrzerwij.BackColor = RGB(0, 0, 0)
    StatusForm.BtnPrzerwij.ForeColor = RGB(254, 190, 16)
    
    StatusForm.show vbModeless
End Sub

Public Sub hide()
    StatusForm.hide
End Sub



Private Sub Class_Initialize()


    StatusForm.LabelEstimatedTime.Caption = 0
    StatusForm.LabelAverageTime.Caption = 0


    szerokosc_paska = 360
    grad_color = 155 ' od 155 do 255 kolor 100 razy zmieniany
    jeden_procent = CDbl(szerokosc_paska) / CDbl(100)
    act = 0
    StatusForm.PasekPostepu.Width = 0
    StatusForm.PasekPostepu.BackColor = RGB(255, 0, 0)
    If max <> 0 Then
        StatusForm.FrameProgress.Caption = CStr(Int((CDbl(act) / CDbl(max)) * 100)) & " %"
    Else
        StatusForm.FrameProgress.Caption = "0 %"
    End If
    
    
    
End Sub

Private Sub Class_Terminate()
    Set collection_of_avr_on_pn = Nothing
End Sub
