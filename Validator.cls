VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validator"
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




Private pusSh As Worksheet
Private cpl As Worksheet



Public Function getWorksheetForValidation(pusWrkName As String, fileType As E_FILE_TYPE) As Boolean


    Dim shName As String, newShName As String
    Dim rowLabel As Integer, newRowLabel  As Integer
    Dim rRef As Range, refR As Range, newRefR As Range
    
    If fileType = E_MASTER_PUS Then
        Set refR = ThisWorkbook.Sheets("forValidation").Range("D6")
        Set newRefR = ThisWorkbook.Sheets("forValidation").Range("D41")
        
        shName = CStr(ThisWorkbook.Sheets("forValidation").Range("A6").Value)
        newShName = CStr(ThisWorkbook.Sheets("forValidation").Range("A41").Value)
        
        rowLabel = Int(ThisWorkbook.Sheets("forValidation").Range("C6").Value)
        newRowLabel = Int(ThisWorkbook.Sheets("forValidation").Range("C41").Value)
        
        
    ElseIf fileType = E_FEED_CPL Then
        Set refR = ThisWorkbook.Sheets("forValidation").Range("D2")
        
        shName = CStr(ThisWorkbook.Sheets("forValidation").Range("A2").Value)
        rowLabel = Int(ThisWorkbook.Sheets("forValidation").Range("C2").Value)
    ElseIf fileType = E_TP04_01 Then
        Set refR = ThisWorkbook.Sheets("forValidation").Range("D10")
        
        shName = CStr(ThisWorkbook.Sheets("forValidation").Range("A10").Value)
        rowLabel = Int(ThisWorkbook.Sheets("forValidation").Range("C10").Value)
    ElseIf fileType = E_SQ01 Then
        Set refR = ThisWorkbook.Sheets("forValidation").Range("D14")
        
        shName = CStr(ThisWorkbook.Sheets("forValidation").Range("A14").Value)
        rowLabel = Int(ThisWorkbook.Sheets("forValidation").Range("C14").Value)
    Else
        shName = ""
        rowLabel = -1
        Set refR = Nothing
    End If
    

    getWorksheetForValidation = True

    Dim wrk As Workbook
    
    Dim mainSh As Worksheet, main As Worksheet
    Set main = Nothing
    
    If fileType <> E_SQ01 Then
        Set wrk = Workbooks(CStr(pusWrkName))
        On Error Resume Next
        Set main = wrk.Sheets(shName)
    ElseIf fileType = E_SQ01 Then
        
        Set main = ThisWorkbook.Sheets(CStr(pusWrkName))
    End If
    
    If main Is Nothing Then
        getWorksheetForValidation = False
        
    Else
    
        ' going further
        Dim vr As Range
        Dim srcr As Range
        
        
        If Not refR Is Nothing Then
        
            Set srcr = main.Cells(rowLabel, 1)
            Set vr = refR
            Do
            
                ' 100% must be
                If CStr(vr.Value) = CStr(srcr.Value) Then
                    ' nop required
                Else
                    getWorksheetForValidation = False
                    Exit Do
                End If
                
                Set vr = vr.offset(0, 1)
                Set srcr = srcr.offset(0, 1)
            Loop Until vr.Value = ""
            
            ' check 2nd possibility
            
            If getWorksheetForValidation = False Then
            
                ' start over again...
                getWorksheetForValidation = True
                
                Set srcr = main.Cells(newRowLabel, 1)
                Set vr = newRefR
                Do
                
                    ' 100% must be for this statement
                    If CStr(vr.Value) = CStr(srcr.Value) Then
                        ' nop required
                    Else
                    
                    
                        ' because TMC column changing the name all the time - it is req
                        ' to have extra logic - if for Validation trying to check TMC stuff
                        ' then just ignore what is inside - empty mask allow to name it
                        ' however you want!
                        If vr.Value = "TMC" Or vr.Value = "Delivery's Grouping" Then
                            ' nop required
                        Else
                    
                            getWorksheetForValidation = False
                            Exit Do
                        End If
                    End If
                    
                    
                    Set vr = vr.offset(0, 1)
                    Set srcr = srcr.offset(0, 1)
                Loop While vr.Column < (E_MASTER_2_ML + 1) ' 48 + 1
            End If
        
        Else
            getWorksheetForValidation = False
        End If
        
    End If
    
End Function


Public Function checkIfComboFormIsFilledProperly(frm1 As Variant, Optional strType As String) As Boolean
    checkIfComboFormIsFilledProperly = True
    
    Dim c As Control, cs As Controls
    
    
    Set cs = frm1.Controls
    
    
    
    If strType = "GREEN_LIGHT" Then
    
        For Each c In cs
        
            If c.name Like "TextBox?1" Then
            
                If c.Value = "" Then
                    checkIfComboFormIsFilledProperly = False
                    Exit For
                Else
                    checkIfComboFormIsFilledProperly = True
                End If
            
            End If
        
        Next c
    Else
    
        For Each c In cs
        
            If c.name Like "TextBoxMag*" Then
            
                If c.Value = "" Then
                    checkIfComboFormIsFilledProperly = False
                    Exit For
                Else
                    checkIfComboFormIsFilledProperly = True
                End If
            
            End If
        
        Next c
    End If
    
End Function




Public Function checkIsCodeOK(cec As String) As Boolean
    checkIsCodeOK = False
    
    
    If Len(CStr(cec)) = 3 Then
        If IsNumeric(CStr(cec)) Then
            If Left(cec, 2) = "11" Then
                checkIsCodeOK = True
            End If
        End If
    End If
    
End Function




Private Sub Class_Initialize()
    Set pusSh = Nothing
    Set cpl = Nothing
End Sub

Private Sub Class_Terminate()
    Set pusSh = Nothing
    Set cpl = Nothing
End Sub




