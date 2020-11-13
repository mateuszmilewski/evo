Attribute VB_Name = "FinalTouchOnReceptionModule"
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


Public Sub finalTouchOnReceptionReport(ictrl As IRibbonControl)
    innerFinalTouchOnReceptionReport ActiveSheet
    
    addFormatConditionsForReceptionReport
    
    startModelessLeaf
End Sub


Public Sub innerFinalTouchOnReceptionReport(Optional sh As Worksheet, Optional auto1 As Boolean, Optional ycw As String)


    ' before you do anything check if active sheet is in proper standard
    Dim vr1 As Range, vr2 As Range
    Set vr1 = ThisWorkbook.Sheets("forValidation").Range("D35")
    Set vr2 = ThisWorkbook.Sheets("forValidation").Range("D38")
    
    
    If validMb51data(sh, vr1) Or validMb51data(sh, vr2) Then
        
        Dim coll As Collection
        Set coll = Nothing
        Set coll = defineTiming(sh, ycw)
        
        Dim sh1 As Worksheet
        Set sh1 = sh
        
        
        Dim finalOut As Worksheet

        Set finalOut = ThisWorkbook.Sheets.Add
        finalOut.name = EVO.TryToRenameModule.tryToRenameWorksheet(finalOut, "RECEPTION_")

        fillLabelsForReceptions finalOut
        
        goThroughMb51Data sh1, finalOut, coll
        
        
        
        addFormatConditionsForReceptionReport
        
        ' some hiding and font changing with column width adjustment
        beautifyReceptionList
        
        startModelessLeaf
    Else
        MsgBox "Active worksheet is in wrong standard!"
    End If
    
End Sub

Private Function defineTiming(sh1 As Worksheet, Optional ycw As String) As Collection
    Set defineTiming = Nothing
    
    
    Dim dic As Dictionary
    Set dic = New Dictionary
    Dim kij As Variant
    
    Dim ir As Range
    Set ir = sh1.Cells(2, 1)
    
    Dim e As E_MB51_AUTO_DECISION_LAYOUT
    
    If sh1.Cells(1, 1).Value = "Article" Then
        e = E_MB51_AUTO_DECISION_LAYOUT_0
    Else
        e = E_MB51_AUTO_DECISION_LAYOUT_NEW
    End If
    
    
    If e = E_MB51_AUTO_DECISION_LAYOUT_NEW Then
    
    
        If ycw = "*" Then
            
            ' ------------------------------------------
            
            ' ------------------------------------------
        ElseIf ycw = "" Then
        
            Do
                If ir.Offset(0, EVO.E_MB51_NEW_CW - 1).Value Like "*CW*" Then
                    If Not dic.Exists(CStr(ir.Offset(0, EVO.E_MB51_NEW_CW - 1).Value)) Then
                        dic.Add CStr(ir.Offset(0, EVO.E_MB51_NEW_CW - 1).Value), 1
                    Else
                        'dic(CStr(ir.Offset(0, EVO.E_MB51_CW - 1).Value)) = _
                        '    dic(CStr(ir.Offset(0, EVO.E_MB51_CW - 1).Value)) + 1
                    End If
                End If
                Set ir = ir.Offset(1, 0)
            Loop Until Trim(ir.Value) = ""
            
        
            With FinalScope.ListBox1
                .Clear
                
                
                For Each kij In dic.Keys
                    .addItem CStr(kij)
                Next
            End With
            
            FinalScope.show
            
        Else
    
            
            ' going directly to timing collection
            Do
            
                If ir.Offset(0, EVO.E_MB51_NEW_CW - 1).Value Like "*CW*" Then
                
                    If ir.Offset(0, EVO.E_MB51_NEW_CW - 1).Value = CStr(ycw) Then
                
                        If Not dic.Exists(CStr(ir.Offset(0, EVO.E_MB51_NEW_CW - 1).Value)) Then
                            dic.Add CStr(ir.Offset(0, EVO.E_MB51_NEW_CW - 1).Value), 1
                        Else
                            'dic(CStr(ir.Offset(0, EVO.E_MB51_CW - 1).Value)) = _
                            '    dic(CStr(ir.Offset(0, EVO.E_MB51_CW - 1).Value)) + 1
                        End If
                    End If
                End If
                
                Set ir = ir.Offset(1, 0)
                
            Loop Until Trim(ir.Value) = ""
            
            Set FinalScope.c = Nothing
            Set FinalScope.c = New Collection
            
            For Each kij In dic
                FinalScope.c.Add CStr(kij)
            Next
            
        End If
        
    ElseIf e = E_MB51_AUTO_DECISION_LAYOUT_0 Then
    
        If ycw = "*" Then
    
    
        ElseIf ycw = "" Then
        
            Do
                If ir.Offset(0, EVO.E_MB51_0_CW - 1).Value Like "*CW*" Then
                    If Not dic.Exists(CStr(ir.Offset(0, EVO.E_MB51_0_CW - 1).Value)) Then
                        dic.Add CStr(ir.Offset(0, EVO.E_MB51_0_CW - 1).Value), 1
                    Else
                        'dic(CStr(ir.Offset(0, EVO.E_MB51_CW - 1).Value)) = _
                        '    dic(CStr(ir.Offset(0, EVO.E_MB51_CW - 1).Value)) + 1
                    End If
                End If
                Set ir = ir.Offset(1, 0)
            Loop Until Trim(ir.Value) = ""
            
        
            With FinalScope.ListBox1
                .Clear
                
                
                For Each kij In dic.Keys
                    .addItem CStr(kij)
                Next
            End With
            
            FinalScope.show
            
        Else
    
            
            ' going directly to timing collection
            Do
            
                If ir.Offset(0, EVO.E_MB51_0_CW - 1).Value Like "*CW*" Then
                
                    If ir.Offset(0, EVO.E_MB51_0_CW - 1).Value = CStr(ycw) Then
                
                        If Not dic.Exists(CStr(ir.Offset(0, EVO.E_MB51_0_CW - 1).Value)) Then
                            dic.Add CStr(ir.Offset(0, EVO.E_MB51_0_CW - 1).Value), 1
                        Else
                            'dic(CStr(ir.Offset(0, EVO.E_MB51_CW - 1).Value)) = _
                            '    dic(CStr(ir.Offset(0, EVO.E_MB51_CW - 1).Value)) + 1
                        End If
                    End If
                End If
                
                Set ir = ir.Offset(1, 0)
                
            Loop Until Trim(ir.Value) = ""
            
            Set FinalScope.c = Nothing
            Set FinalScope.c = New Collection
            
            For Each kij In dic
                FinalScope.c.Add CStr(kij)
            Next
            
        End If
    
    End If
    
    
    If Not FinalScope.c Is Nothing Then
        If FinalScope.c.count > 0 Then
            Debug.Print "FinalScope.c.Count: " & FinalScope.c.count
            Set defineTiming = FinalScope.c
        End If
    End If
End Function

Private Sub goThroughMb51Data(sh1 As Worksheet, out As Worksheet, coll As Collection)



    Application.Calculation = xlCalculationManual


    ' r1 as input
    ' r2 for ouput
    
    Dim r1 As Range, r2 As Range
    Set r1 = sh1.Cells(2, 1)
    Set r2 = out.Cells(2, 1)
    
    
    Dim e As E_MB51_AUTO_DECISION_LAYOUT
    
    If sh1.Cells(1, 1) = "Article" Then
        e = E_MB51_AUTO_DECISION_LAYOUT_0
    Else
        e = E_MB51_AUTO_DECISION_LAYOUT_NEW
    End If
    
    
    Do
        ' -----------------------------------------
        
        If isThisLineGoes(sh1, CLng(r1.row), coll, e) Then
        
        
            ' out out out
            ' -------------------------------------
            
            If e = E_MB51_AUTO_DECISION_LAYOUT_NEW Then
            
                'E_FINAL_TOUCH_RECEPTION_Mag = 1
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Mag - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_MAG - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Mvt
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Mvt - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_MVT - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Designation_article
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Designation_article - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_DESC - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Qte_UAc
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Qte_UAc - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_QTY - 1).Value
                'E_FINAL_TOUCH_RECEPTION_UAc
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_UAc - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_UNX - 1).Value
                
                'E_FINAL_TOUCH_RECEPTION_Montant_di
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Montant_di - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_MONTANT_DI - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Dev
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Dev - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_DEVISE - 1).Value
                'E_FINAL_TOUCH_RECEPTION_article
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_article - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_ARTICLE - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Date_cpt
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Date_cpt - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_DATE_PIECE - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Fourn
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Fourn - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_FOUR - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Cde_achat
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Cde_achat - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_CDE_ACHAT - 1).Value
                'E_FINAL_TOUCH_RECEPTION_prix_Sigapp
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_Sigapp - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_EXT_PCS_PRICE_IN_EUR - 1).Value
                'E_FINAL_TOUCH_RECEPTION_prix_Tango
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_Tango - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_TANGO_PCS_PRICE - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Ecart
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Ecart - 1).FormulaR1C1Local = _
                    ThisWorkbook.Sheets("register").Range("Z11").FormulaR1C1Local
                'E_FINAL_TOUCH_RECEPTION_RU
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_RU - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_RU - 1).Value
                
                'E_FINAL_TOUCH_RECEPTION_prix_cible
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_cible - 1).FormulaR1C1Local = _
                    ThisWorkbook.Sheets("register").Range("Z12").FormulaR1C1Local
                'E_FINAL_TOUCH_RECEPTION_Sigapp
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sigapp - 1).FormulaR1C1Local = _
                    ThisWorkbook.Sheets("register").Range("Z13").FormulaR1C1Local
                    
                'E_FINAL_TOUCH_RECEPTION_GAP
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_GAP - 1).FormulaR1C1Local = _
                    ThisWorkbook.Sheets("register").Range("Z21").FormulaR1C1Local
                    
                'E_FINAL_TOUCH_RECEPTION_OKNOK
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_OKNOK - 1).FormulaR1C1Local = _
                    ThisWorkbook.Sheets("register").Range("Z22").FormulaR1C1Local
                    
                'E_FINAL_TOUCH_RECEPTION_Interne
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Interne - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_IS_INTERNAL - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Sem
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sem - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_NEW_CW - 1).Value
                ' manager
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Manager_DA - 1).Value = "tbd"
            
            ElseIf e = E_MB51_AUTO_DECISION_LAYOUT_0 Then
            
            
                'E_FINAL_TOUCH_RECEPTION_Mag = 1
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Mag - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_MAG - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Mvt
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Mvt - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_MVT - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Designation_article
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Designation_article - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_DESC - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Qte_UAc
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Qte_UAc - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_QTY - 1).Value
                'E_FINAL_TOUCH_RECEPTION_UAc
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_UAc - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_UN - 1).Value
                
                'E_FINAL_TOUCH_RECEPTION_Montant_di
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Montant_di - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_MONTANT_DI - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Dev
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Dev - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_DEVISE - 1).Value
                'E_FINAL_TOUCH_RECEPTION_article
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_article - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_ARTICLE - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Date_cpt
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Date_cpt - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_DATE_PIECE - 1).Value
                    
                    
                ' Application.Calculation = xlCalculationManual
                    
                'E_FINAL_TOUCH_RECEPTION_Fourn
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Fourn - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_FOUR - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Cde_achat
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Cde_achat - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_CDE_ACHAT - 1).Value
                'E_FINAL_TOUCH_RECEPTION_prix_Sigapp
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_Sigapp - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_EXT_PCS_PRICE_IN_EUR - 1).Value
                'E_FINAL_TOUCH_RECEPTION_prix_Tango
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_Tango - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_TANGO_PCS_PRICE - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Ecart
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Ecart - 1).FormulaR1C1Local = _
                    ThisWorkbook.Sheets("register").Range("Z11").FormulaR1C1Local
                'E_FINAL_TOUCH_RECEPTION_RU
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_RU - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_RU - 1).Value
                
                'E_FINAL_TOUCH_RECEPTION_prix_cible
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_cible - 1).FormulaR1C1Local = _
                    ThisWorkbook.Sheets("register").Range("Z12").FormulaR1C1Local
                'E_FINAL_TOUCH_RECEPTION_Sigapp
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sigapp - 1).FormulaR1C1Local = _
                    ThisWorkbook.Sheets("register").Range("Z13").FormulaR1C1Local
                    
                    
                'E_FINAL_TOUCH_RECEPTION_GAP
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_GAP - 1).FormulaR1C1Local = _
                    ThisWorkbook.Sheets("register").Range("Z21").FormulaR1C1Local
                    
                'E_FINAL_TOUCH_RECEPTION_OKNOK
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_OKNOK - 1).FormulaR1C1Local = _
                    ThisWorkbook.Sheets("register").Range("Z22").FormulaR1C1Local
                    
                    
                'E_FINAL_TOUCH_RECEPTION_Interne
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Interne - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_IS_INTERNAL - 1).Value
                'E_FINAL_TOUCH_RECEPTION_Sem
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sem - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_CW - 1).Value
                ' manager
                r2.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Manager_DA - 1).Value = _
                    r1.Offset(0, EVO.E_MB51_0_MANAGER_DA - 1).Value
            Else
            
                MsgBox "Critical auto decision on reception output layout!", vbCritical
                End
            End If
            
            ' -------------------------------------
            
            Set r2 = r2.Offset(1, 0)
        End If
        
        ' -----------------------------------------
        
        Set r1 = r1.Offset(1, 0)
    Loop Until Trim(r1.Value) = ""
    
    
    Application.Calculation = xlCalculationAutomatic
    
    
End Sub

Private Function isThisLineGoes(sh1 As Worksheet, mrow As Long, scope As Collection, e As E_MB51_AUTO_DECISION_LAYOUT) As Boolean
    isThisLineGoes = False
    
    ' check CW
    
    If e = E_MB51_AUTO_DECISION_LAYOUT_NEW Then
    
        If Int(sh1.Cells(mrow, EVO.E_MB51_NEW_IS_CANCELLED).Value) = 0 Then
            If Int(sh1.Cells(mrow, EVO.E_MB51_NEW_IS_WITH_INDEX).Value) = 1 Then
                If isInScope(scope, sh1.Cells(mrow, EVO.E_MB51_NEW_CW).Value) Then
                    'If Trim(sh1.Cells(mrow, EVO.E_MB51_IS_INTERNAL).Value) = "" Then
                        isThisLineGoes = True
                    'End If
                End If
            End If
        End If
    
    ElseIf e = E_MB51_AUTO_DECISION_LAYOUT_0 Then
    
        If Int(sh1.Cells(mrow, EVO.E_MB51_0_IS_CANCELLED).Value) = 0 Then
            If Int(sh1.Cells(mrow, EVO.E_MB51_0_IS_WITH_INDEX).Value) = 1 Then
                If isInScope(scope, sh1.Cells(mrow, EVO.E_MB51_0_CW).Value) Then
                    'If Trim(sh1.Cells(mrow, EVO.E_MB51_IS_INTERNAL).Value) = "" Then
                        isThisLineGoes = True
                    'End If
                End If
            End If
        End If

    
    Else
        MsgBox "Critical! Not possible scenario in isThisLineGoesOnFinalTouchForReception!", vbCritical
        End
    End If
End Function

Private Function isInScope(scope As Collection, strKeyForScope As String) As Boolean
    isInScope = False
    
    If scope Is Nothing Then
        isInScope = True
    Else
    
        Dim el As Variant
        For Each el In scope
            If CStr(el) = strKeyForScope Then
                isInScope = True
                Exit For
            End If
        Next el
    End If
End Function

Private Sub fillLabelsForReceptions(sh2 As Worksheet)

    Dim finalReceptionValidationRef As Range
    Set finalReceptionValidationRef = ThisWorkbook.Sheets("forValidation").Range("D23")
    
    Dim iter As Integer
    iter = 1
    Do
    
        sh2.Cells(1, iter).Value = finalReceptionValidationRef.Value
        sh2.Cells(1, iter).Interior.Color = finalReceptionValidationRef.Interior.Color
        sh2.Cells(1, iter).Font.Color = finalReceptionValidationRef.Font.Color

        iter = iter + 1
        Set finalReceptionValidationRef = finalReceptionValidationRef.Offset(0, 1)
    Loop Until Trim(finalReceptionValidationRef.Value) = ""
End Sub



