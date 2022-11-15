Attribute VB_Name = "Sq01ToReceptionModule"
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

Public Sub adjustSq01DataForReception2_FOR_LIST_ALT_F8()
    innerAdjustSq01DataForReception ActiveSheet, Nothing
End Sub

Public Sub adjustSq01DataForReception(ictrl As IRibbonControl)
    innerAdjustSq01DataForReception ActiveSheet, Nothing
End Sub


Public Sub innerAdjustSq01DataForReception(Optional mb51sh As Worksheet, Optional sh1 As Worksheet)


    ' very first simple check!
    If Not mb51sh Is Nothing Then
        If mb51sh.name Like "MB51*" Then
        
            If inLineWithD38(mb51sh) Then
    
                If (sh1 Is Nothing) Or IsMissing(sh1) Then
                
                    Dim nm As String, priceFromTp04 As Double
                    nm = getTheConcatName()
                    
                    
                    If nm <> "" Then
                    
                        Set sh1 = ThisWorkbook.Sheets(nm)
                        
                        Dim ashr As Range
                        Set ashr = mb51sh.Cells(2, 1)
                        
                        Do
                            ' E_MB51_0_
                            ' E_MB51_0_IS_IN_TANGO
                            
                            If ashr.offset(0, EVO.E_MB51_0_IS_IN_TANGO - 1).Value = "NO TANGO" Then
                                ' in this case we need to verify if there is price in TP04
                                
                                priceFromTp04 = 0
                                
                                On Error Resume Next
                                priceFromTp04 = tryToFindAndAssignInitialPriceFromTp04(sh1, ashr.offset(0, EVO.E_MB51_0_ARTICLE - 1))
                                
                                If priceFromTp04 > 0 Then
                                
                                    ashr.offset(0, EVO.E_MB51_0_TANGO_PCS_PRICE - 1).Value = priceFromTp04
                                    ashr.offset(0, EVO.E_MB51_0_IS_IN_TANGO - 1).Value = "TP04 PRICE"
                                End If
                            End If
                            
                            Set ashr = ashr.offset(1, 0)
                        Loop Until Trim(ashr.Value) = ""
                    Else
                        Debug.Print "No CONCAT for initial price from TP04 for reception!"
                    End If
                Else
                End If
            End If
        Else
            MsgBox "Most probably you are in wrong worksheet right now!", vbInformation
        End If
    End If
End Sub

Private Function tryToFindAndAssignInitialPriceFromTp04(csh As Worksheet, ar As Range) As Double
    tryToFindAndAssignInitialPriceFromTp04 = 0
    
    Dim strArticleWithoutIndex As String
    strArticleWithoutIndex = ""
    
    strArticleWithoutIndex = Split(ar.Value, "-")(0)
    
    Dim cr As Range
    Set cr = csh.Cells(2, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_ARTICLE)
    
    Do
    
        If CStr(cr.Value) = CStr(strArticleWithoutIndex) Then
            tryToFindAndAssignInitialPriceFromTp04 = CDbl(cr.offset(0, EVO.E_FROM_SQ01_QUASI_TP04_II_RATE_CALCED_SUM - _
                EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value)
        End If
        
        
        Set cr = cr.offset(1, 0)
    Loop Until Trim(cr.Value) = ""
    
End Function


Private Function inLineWithD38(ash As Worksheet) As Boolean
    inLineWithD38 = False
    
    Dim r As Range, ashr As Range
    Set r = ThisWorkbook.Sheets("forValidation").Range("D38")
    Set ashr = ash.Range("A1")
    
    Do
    
        If r.Value = ashr.Value Then
            ' ok
            inLineWithD38 = True
        Else
            inLineWithD38 = False
            Exit Function
        End If
    
        Set r = r.offset(0, 1)
        Set ashr = ashr.offset(0, 1)
    Loop Until Trim(r) = ""
End Function


Private Function getTheConcatName() As String
    getTheConcatName = ""
    
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        
        If sh.name Like "CONCAT*" Then
            getTheConcatName = CStr(sh.name)
            Exit Function
        End If
    Next sh
End Function
