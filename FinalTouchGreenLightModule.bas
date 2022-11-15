Attribute VB_Name = "FinalTouchGreenLightModule"
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


Public Sub finalTouchOnGreenLightReport(ictrl As IRibbonControl)
    innerFinalTouchOnGreenLightReport ActiveSheet, False
    

End Sub


Public Sub innerFinalTouchOnGreenLightReport(sh1 As Worksheet, auto1 As Boolean)


    ' before you do anything check if active sheet is in proper standard
    If valid_TP04_data(sh1) Then
        
        Dim coll As Collection
        Set coll = Nothing
        Set coll = defineTimingInGreenLight(sh1, auto1)
        
        'Dim sh1 As Worksheet
        'Set sh1 = ActiveSheet
        
        
        If Not coll Is Nothing Then
            If coll.count > 0 Then
            
            
                Dim finalOut As Worksheet
                Set finalOut = ThisWorkbook.Sheets.Add
                finalOut.name = EVO.TryToRenameModule.tryToRenameWorksheet(finalOut, "GREEN_LIGHT_")
        
                fillLabelsForGreenLight finalOut
                
                goThroughTP04Data sh1, finalOut, coll
                
                
                
                
                
                addFormatConditionsForGreenLightReport
                
                ' some hiding and font changing with column width adjustment
                beautifyGreenLightList
                
                startModelessLeaf
                
            Else
                MsgBox "Scope is wrongly defined!", vbCritical
            End If
        Else
            MsgBox "Scope is wrongly defined!", vbCritical
        End If
    Else
        MsgBox "Active worksheet is in wrong standard!"
    End If
    
End Sub

Private Function defineTimingInGreenLight(sh1 As Worksheet, auto1 As Boolean) As Collection
    Set defineTimingInGreenLight = Nothing
    
    
    Dim dic As Dictionary
    Set dic = New Dictionary
    
    Dim ir As Range
    Set ir = sh1.Cells(2, 1)
    
    Do
        If ir.offset(0, EVO.E_ADJUSTED_SQ01_CW - 1).Value Like "*CW*" Or ir.offset(0, EVO.E_ADJUSTED_SQ01_CW - 1).Value Like "*S*/*" Then
            If Not dic.Exists(CStr(ir.offset(0, EVO.E_ADJUSTED_SQ01_CW - 1).Value)) Then
                dic.Add CStr(ir.offset(0, EVO.E_ADJUSTED_SQ01_CW - 1).Value), 1
            Else
                'dic(CStr(ir.Offset(0, EVO.E_MB51_CW - 1).Value)) = _
                '    dic(CStr(ir.Offset(0, EVO.E_MB51_CW - 1).Value)) + 1
            End If
        End If
        Set ir = ir.offset(1, 0)
    Loop Until Trim(ir.Value) = ""
    

    With FinalScope.ListBox1
        .Clear
        
        Dim kij As Variant
        For Each kij In dic.Keys
            .addItem CStr(kij)
        Next
    End With
    
    
    If auto1 = False Then
    
        FinalScope.show
        
        
        If Not FinalScope.c Is Nothing Then
            If FinalScope.c.count > 0 Then
                Debug.Print "FinalScope.c.Count: " & FinalScope.c.count
                Set defineTimingInGreenLight = FinalScope.c
            End If
        End If
    Else
    
    
        Dim ctmp As New Collection

        With FinalScope.ListBox1
        
            Dim x As Variant
            For x = 0 To .ListCount - 1
                ctmp.Add .list(x)
            Next x
        End With
        
        Set defineTimingInGreenLight = ctmp
    End If
End Function

Private Sub goThroughTP04Data(sh1 As Worksheet, out As Worksheet, coll As Collection)



    Application.Calculation = xlCalculationManual

    ' r1 as input
    ' r2 for ouput
    
    Dim r1 As Range, r2 As Range, x As Variant
    Set r1 = sh1.Cells(2, 1)
    Set r2 = out.Cells(2, 1)
    
    Do
        ' -----------------------------------------
        
        If isThisLineGoes(sh1, CLng(r1.row), coll) Then
        
        
            ' out out out
            ' -------------------------------------
            
            
            ' all inside loop
            ' \/\/\/\/\/\/

            
            For x = EVO.E_ADJUSTED_SQ01_Reference To EVO.E_ADJUSTED_SQ01_OKNOK
            
                r2.offset(0, x - 1).Value = r1.offset(0, x - 1).Value
            Next x
            
            'E_GREEN_LIGHT_TANGO_PRICE
            r2.offset(0, E_GREEN_LIGHT_TANGO_PRICE - 1).Value = _
                r1.offset(0, EVO.E_ADJUSTED_SQ01_TANGO_PCS_PRICE - 1).Value
            
            'E_GREEN_LIGHT_TANGO_RATE
            r2.offset(0, E_GREEN_LIGHT_TANGO_RATE - 1).FormulaR1C1Local = _
                ThisWorkbook.Sheets("register").Range("Z14").FormulaR1C1Local
                
            'E_GREEN_LIGHT_TANGO_OKNOK
            r2.offset(0, E_GREEN_LIGHT_TANGO_OKNOK - 1).FormulaR1C1Local = _
                ThisWorkbook.Sheets("register").Range("Z15").FormulaR1C1Local
                
            'E_GREEN_LIGHT_MANAGER
            r2.offset(0, E_GREEN_LIGHT_MANAGER - 1).Value = _
                r1.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_DA - 1).Value
            
            'E_GREEN_LIGHT_Spending_Sigapp
            r2.offset(0, E_GREEN_LIGHT_Spending_sigapp - 1).FormulaR1C1Local = _
                ThisWorkbook.Sheets("register").Range("Z16").FormulaR1C1Local
            
            'E_GREEN_LIGHT_Spending_Target
            r2.offset(0, E_GREEN_LIGHT_Spending_Target - 1).FormulaR1C1Local = _
                ThisWorkbook.Sheets("register").Range("Z17").FormulaR1C1Local
                
            'E_GREEN_LIGHT_Gap
            r2.offset(0, E_GREEN_LIGHT_Gap - 1).FormulaR1C1Local = _
                ThisWorkbook.Sheets("register").Range("Z18").FormulaR1C1Local
                
            'E_GREEN_LIGHT_rate
            r2.offset(0, E_GREEN_LIGHT_rate - 1).FormulaR1C1Local = _
                ThisWorkbook.Sheets("register").Range("Z19").FormulaR1C1Local
                
            'E_GREEN_LIGHT_INTERNAL
            r2.offset(0, EVO.E_GREEN_LIGHT_IS_INTERNAL - 1).Value = _
                r1.offset(0, EVO.E_ADJUSTED_SQ01_IS_INTERNAL - 1).Value
            
            'E_GREEN_LIGHT_WITH_INDEX
            r2.offset(0, EVO.E_GREEN_LIGHT_WITH_INDEX - 1).Value = _
                r1.offset(0, EVO.E_ADJUSTED_SQ01_WITH_INDEX - 1).Value
            
            'E_GREEN_LIGHT_IS_IN_TANGO
            r2.offset(0, E_GREEN_LIGHT_IS_IN_TANGO - 1).Value = _
                r1.offset(0, EVO.E_ADJUSTED_SQ01_IS_IN_TANGO - 1).Value
                
                
                
            ' new from 060
            ' =============================================================
            ' =============================================================
            r2.offset(0, EVO.E_GREEN_LIGHT_DOMAIN - 1).Value = _
                r1.offset(0, EVO.E_ADJUSTED_SQ01_DOMAIN - 1).Value
            r2.offset(0, EVO.E_GREEN_LIGHT_RU - 1).Value = _
                r1.offset(0, EVO.E_ADJUSTED_SQ01_RU - 1).Value
            r2.offset(0, EVO.E_GREEN_LIGHT_DIV - 1).Value = _
                r1.offset(0, EVO.E_ADJUSTED_SQ01_DIV - 1).Value
            ' =============================================================
            ' =============================================================
            
            ' new from 092
            ' =============================================================
            ' =============================================================
            r2.offset(0, EVO.E_GREEN_LIGHT_FAMILY - 1).Value = _
                r1.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_FAMILY - 1).Value
            r2.offset(0, EVO.E_GREEN_LIGHT_GROUP - 1).Value = _
                r1.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_GROUP - 1).Value
            ' =============================================================
            ' =============================================================

           
            ' -------------------------------------
            
            Set r2 = r2.offset(1, 0)
        End If
        
        ' -----------------------------------------
        
        Set r1 = r1.offset(1, 0)
    Loop Until Trim(r1.Value) = ""
    
    
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    
    
End Sub

Private Function isThisLineGoes(sh1 As Worksheet, mrow As Long, scope As Collection) As Boolean
    isThisLineGoes = False
    
    ' check CW
    If CStr(sh1.Cells(mrow, EVO.E_ADJUSTED_SQ01_Reference).Value) Like "*-??" Then
    
        sh1.Cells(mrow, EVO.E_ADJUSTED_SQ01_WITH_INDEX).Value = 1
    
        If isInScope(scope, sh1.Cells(mrow, EVO.E_ADJUSTED_SQ01_CW).Value) Then
            'If Trim(sh1.Cells(mrow, EVO.E_MB51_IS_INTERNAL).Value) = "" Then
                isThisLineGoes = True
            'End If
        End If
    End If
End Function

Private Function isInScope(scope As Collection, strKeyForScope As String) As Boolean
    isInScope = False
    
    Dim el As Variant
    For Each el In scope
        If CStr(el) = strKeyForScope Then
            isInScope = True
            Exit For
        End If
    Next el
End Function

Private Sub fillLabelsForGreenLight(sh2 As Worksheet)

    Dim finalValidationRef As Range
    Set finalValidationRef = ThisWorkbook.Sheets("forValidation").Range("D26")
    
    Dim iter As Integer
    iter = 1
    Do
    
        sh2.Cells(1, iter).Value = finalValidationRef.Value
        sh2.Cells(1, iter).Interior.Color = finalValidationRef.Interior.Color
        sh2.Cells(1, iter).Font.Bold = True

        iter = iter + 1
        Set finalValidationRef = finalValidationRef.offset(0, 1)
    Loop Until Trim(finalValidationRef.Value) = ""
End Sub



' common function used also for managers da logic!
Public Function valid_TP04_data(sh1 As Worksheet) As Boolean
    valid_TP04_data = False
    
    Dim validationRef As Range, labelsRef As Range
    Set validationRef = ThisWorkbook.Sheets("forValidation").Range("D32")
    Set labelsRef = ActiveSheet.Cells(1, 1)
    
    Do
    
        If UCase(validationRef.Value) = UCase(labelsRef.Value) Then
            valid_TP04_data = True
        Else
            valid_TP04_data = False
            Exit Do
        End If
    
        Set validationRef = validationRef.offset(0, 1)
        Set labelsRef = labelsRef.offset(0, 1)
    Loop Until Trim(labelsRef.Value) = ""
    
    
    If valid_TP04_data Then
        Debug.Print "activesheet is in std!"
    End If
    
    
End Function


