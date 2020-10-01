Attribute VB_Name = "InternalSuppliersModule"
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

Public Sub getInternalSuppliers(ictrl As IRibbonControl)

    Dim sh As Worksheet
    isolatedLogicForInternalSuppliers sh
End Sub




Public Sub isolatedLogicForInternalSuppliers(ByRef shInternal As Worksheet)


    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    
    Dim numHandler As NumberHandler
    Set numHandler = New NumberHandler


    Dim st_h As StatusHandler, xStHelper As Integer
    Set st_h = New StatusHandler
    st_h.init_statusbar 20
    st_h.show
    
    
    delegacjaDlaProgresu st_h, xStHelper, 20
    
    
    ' inter4sh stands for internal suppliers list worksheet
    Dim inter4Sh As Worksheet
    Set inter4Sh = ThisWorkbook.Sheets.Add
    inter4Sh.name = EVO.TryToRenameModule.tryToRenameWorksheet(inter4Sh, "N_" & "_")
    
    Set shInternal = inter4Sh
    
    
    ' LABLES ------------------------------------------------
    
    fillLabels inter4Sh.Range("A1")
    
    ' -------------------------------------------------------
    
    Dim sap__handler As New SAP_Handler
    
    ' automatisation on sigapp for internal suppliers - no matter what!
    sap__handler.runMainLogicFor__Y_PI1_80000391 inter4Sh, st_h, xStHelper
    
    
    
    
    
    st_h.hide
    Set st_h = Nothing
    
    
    
    Set numHandler = Nothing
    
    
    
    
'    Dim answer As Variant
'    answer = MsgBox("Raw output from SQ01 (quasi TP04) ready! Do you want to continue?", vbYesNo + vbQuestion)
'
'    If answer = vbYes Then
'        runAdjusterForDataFromSq01 osh
'    End If
    
    
    ' MsgBox "GOTOWE!", vbInformation
    
    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
End Sub




Public Sub matchSq01WithInternalSuppliers(ictrl As IRibbonControl)
    Debug.Print "matchSq01WithIS"
    
    fillForm
End Sub


Public Sub matchMb51WithInternalSuppliers(ictrl As IRibbonControl)
    Debug.Print "matchMb51WithIS"
    
    fillForm
End Sub

Private Sub fillForm()
    
    FindTangoOrIntSData.ComboBox1.Clear
    FindTangoOrIntSData.Caption = "Match with Internal Suppliers"
    
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.name Like "N_*" Then
            FindTangoOrIntSData.ComboBox1.addItem sh.name
        End If
    Next sh
    
    FindTangoOrIntSData.show
End Sub


Public Sub runMatchingLogicOnInternalSuppliers(sh As Worksheet, nData As Worksheet)
    
    
    Dim rng As Range, bottomRng As Range, area As Range
    
    Dim e As E_MB51_AUTO_DECISION_LAYOUT
    
    If sh.name Like "TP04*" Then
        Set rng = sh.Range("A2").Offset(0, EVO.E_ADJUSTED_SQ01_COFOR - 1)
        Set bottomRng = sh.Range("A2").End(xlDown).Offset(0, EVO.E_ADJUSTED_SQ01_COFOR - 1)
    Else
    
        If sh.Range("A1").Value = "Article" Then
            Set rng = sh.Range("A2").Offset(0, EVO.E_MB51_0_FOUR - 1)
            e = E_MB51_AUTO_DECISION_LAYOUT_0
        Else
            Set rng = sh.Range("A2").Offset(0, EVO.E_MB51_NEW_FOUR - 1)
            e = E_MB51_AUTO_DECISION_LAYOUT_NEW
        End If
        
        
        If e = E_MB51_AUTO_DECISION_LAYOUT_0 Then
            Set bottomRng = sh.Range("A2").End(xlDown).Offset(0, EVO.E_MB51_0_FOUR - 1)
        Else
            Set bottomRng = sh.Range("A2").End(xlDown).Offset(0, EVO.E_MB51_NEW_FOUR - 1)
        End If
    End If
    
    
    
    Set area = sh.Range(rng, bottomRng)
    
    Debug.Print "area for internal suppliers calc: " & area.Address
    
    ' and N data will be all the time the same
    
    Dim nrng As Range, bottom_nrng As Range, n_area As Range
    Set nrng = nData.Range("A2").Offset(0, EVO.E_N_SUPPLIERS_COFOR - 1)
    Set bottom_nrng = nrng.End(xlDown)
    Set n_area = nData.Range(nrng, bottom_nrng)
    
    If sh.name Like "TP04*" Then
    
        For Each rng In area
        
        
            rng.Offset(0, EVO.E_ADJUSTED_SQ01_IS_INTERNAL - EVO.E_ADJUSTED_SQ01_COFOR).Value = ""
        
            For Each nrng In n_area
            
                If Trim(nrng.Value) = Trim(rng.Value) Then
                    rng.Offset(0, EVO.E_ADJUSTED_SQ01_IS_INTERNAL - EVO.E_ADJUSTED_SQ01_COFOR).Value = "internal"
                End If
            Next nrng
        Next rng
        
    Else
        
        
        For Each rng In area
        
        
            If e = E_MB51_AUTO_DECISION_LAYOUT_NEW Then
        
            
                rng.Offset(0, EVO.E_MB51_NEW_IS_INTERNAL - EVO.E_MB51_NEW_FOUR).Value = ""
                
                If rng.Value <> "" Then
            
                    For Each nrng In n_area
                    
                        ' Debug.Print nrng.row & " " & rng.row
                    
                        If Trim(nrng.Value) = Trim(rng.Value) Then
                            rng.Offset(0, EVO.E_MB51_NEW_IS_INTERNAL - EVO.E_MB51_NEW_FOUR).Value = "internal"
                        End If
                    Next nrng
                Else
                    ' Exit For
                End If
                
                
            ElseIf e = E_MB51_AUTO_DECISION_LAYOUT_0 Then
            
                rng.Offset(0, EVO.E_MB51_0_IS_INTERNAL - EVO.E_MB51_0_FOUR).Value = ""
                
                If rng.Value <> "" Then
            
                    For Each nrng In n_area
                    
                        ' Debug.Print nrng.row & " " & rng.row
                    
                        If Trim(nrng.Value) = Trim(rng.Value) Then
                            rng.Offset(0, EVO.E_MB51_0_IS_INTERNAL - EVO.E_MB51_0_FOUR).Value = "internal"
                        End If
                    Next nrng
                Else
                    ' Exit For
                End If
            
            Else
                MsgBox "Critical stop - no scenario for reception!", vbCritical
                End
            End If
        Next rng
        
    End If
    
End Sub



