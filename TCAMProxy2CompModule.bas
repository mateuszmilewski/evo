Attribute VB_Name = "TCAMProxy2CompModule"
Option Explicit

'The MIT License (MIT)
'
'Copyright (c) 2021 FORREST
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

' proxy 2 comparison


Public Sub compareProxy2(ictrl As IRibbonControl)
    
    Debug.Print "start of comparison is here!"
    ' =====================================================================================
    
    Dim foundation As Worksheet
    
    Set foundation = ThisWorkbook.ActiveSheet
    
    If foudationMakeSense(foundation) Then
    
        FinalScope.ListBox1.Clear
        
        Dim sh As Worksheet
        For Each sh In ThisWorkbook.Sheets
            If (sh.name Like "Proxy2*") And (sh.name <> foundation.name) Then
                FinalScope.ListBox1.addItem sh.name
            End If
        Next sh
        
        FinalScope.show
        
        Dim c1 As Collection, str As Variant
        Set c1 = FinalScope.c
        
        If Not c1 Is Nothing Then
            If c1.count > 0 Then
                For Each str In c1
                
                    makeComparisonIterationForFoundation foundation, ThisWorkbook.Sheets(CStr(str))
                Next str
            End If
        End If
    Else
        MsgBox "Active sheet in Proxy2 standard required!"
    End If
    
    
    
    ' =====================================================================================
End Sub


Private Function foudationMakeSense(sh1 As Worksheet) As Boolean

    foudationMakeSense = False
    
    If sh1.name Like "Proxy2_*" Then
        If sh1.Cells(1, 1).Value = "ID" Then
            If sh1.Cells(1, 2).Value = "WIERSZ" Then
                foudationMakeSense = True
            End If
        End If
        
    End If
End Function


Private Sub makeComparisonIterationForFoundation(f As Worksheet, sh As Worksheet)
    
    ' this is an interation for pair of the sheets in EVO 120
    ' f is foundation and it is a target
    ' sh is source - checking differences
    
    Dim fRef As Range, shRef As Range, x As Variant
    Set fRef = f.Cells(2, EVO.E_PIVOT_PROXY2_ID)
    
    Do
        ' ============================================================================
        
        ' E_PIVOT_PROXY2
        Set shRef = sh.Cells(2, EVO.E_PIVOT_PROXY2_ID)
        
    
        Do
        
            If Trim(shRef.Value) = Trim(fRef.Value) Then
            
                ' f.Cells(fRef.row, x).Interior.Color = RGB(255, 255, 255)
            
            
                For x = (EVO.E_PIVOT_PROXY2.E_PIVOT_PROXY2_WIERSZ) To EVO.E_PIVOT_PROXY2.E_PIVOT_PROXY2__lb_INCOTERM
                
                    If CStr("" & sh.Cells(shRef.row, x).Value) <> CStr("" & f.Cells(fRef.row, x).Value) Then
                        
                        If f.Cells(fRef.row, x).Comment Is Nothing Then
                            f.Cells(fRef.row, x).AddComment CStr(sh.Cells(shRef.row, x).Value)
                            f.Cells(fRef.row, x).Interior.Color = CLng(f.Cells(fRef.row, x).Interior.Color) - 100
                        Else
                            f.Cells(fRef.row, x).Comment.Text CStr(sh.Cells(shRef.row, x).Value) & Chr(10), 1, False
                            f.Cells(fRef.row, x).Interior.Color = CLng(f.Cells(fRef.row, x).Interior.Color) - 100
                        End If
                    End If
                    
                Next x
                
                Exit Do
            End If
            
            Set shRef = shRef.offset(1, 0)
        
        Loop Until Trim(shRef.Value) = ""
        
        
        ' ============================================================================
        Set fRef = fRef.offset(1, 0)
    Loop Until Trim(fRef.Value) = ""
    
End Sub
