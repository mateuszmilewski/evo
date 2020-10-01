VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeedbackHandler"
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

' THE FEEDBACK HANDLER

' checking data from output of reception
' and verifying output from sq01


Private receptionSheet As Worksheet
Private greenLightSheet As Worksheet


Private dictionaryOfFeedbackItems As Dictionary



Public Function getElementFromDictionary(mkey As String) As FeedbackItem
    Set getElementFromDictionary = Nothing
    
    On Error Resume Next
    Set getElementFromDictionary = dictionaryOfFeedbackItems(mkey)
End Function


Public Sub tryToAddNewElement(fi As FeedbackItem)
    
    If dictionaryOfFeedbackItems.Exists(CStr(fi.rangeRefernceInGreen.row)) Then
    Else
        dictionaryOfFeedbackItems.Add CStr(fi.rangeRefernceInGreen.row), fi
    End If
End Sub



Public Sub setupSheets(sh1 As Worksheet, sh2 As Worksheet)
    
    Set receptionSheet = sh1
    Set greenLightSheet = sh2
End Sub


Public Function checkIfReallyGreenLight(sh1 As Worksheet) As Boolean
    checkIfReallyGreenLight = False
    
    
    ' it is all about checking labels!
    
    Dim ref As Range
    ' recv 0 only
    ' there was some fluctuation some time ago - there was 2 standards
    ' but for feedback logic Im taking only the RECV0 standard
    ' for me it is foundation!
    Set ref = ThisWorkbook.Sheets("forValidation").Range("D26")
    
    Dim r1 As Range
    Set r1 = sh1.Range("A1")
    
    Do
        If CStr(r1.Value) = CStr(ref.Value) Then
            checkIfReallyGreenLight = True
        Else
            checkIfReallyGreenLight = False
            Exit Function
        End If
        
        Set ref = ref.Offset(0, 1)
        Set r1 = r1.Offset(0, 1)
    Loop Until Trim(r1) = ""
End Function

Public Function checkIfReallyReception(sh1 As Worksheet) As Boolean
    checkIfReallyReception = False
    
    
    ' it is all about checking labels!
    
    Dim ref As Range
    ' recv 0 only
    ' there was some fluctuation some time ago - there was 2 standards
    ' but for feedback logic Im taking only the RECV0 standard
    ' for me it is foundation!
    Set ref = ThisWorkbook.Sheets("forValidation").Range("D23")
    
    Dim r1 As Range
    Set r1 = sh1.Range("A1")
    
    Do
        If CStr(r1.Value) = CStr(ref.Value) Then
            checkIfReallyReception = True
        Else
            checkIfReallyReception = False
            Exit Function
        End If
        
        Set ref = ref.Offset(0, 1)
        Set r1 = r1.Offset(0, 1)
    Loop Until Trim(r1) = ""
End Function

Private Sub Class_Initialize()
    Set receptionSheet = Nothing
    Set greenLightSheet = Nothing
    
    Set dictionaryOfFeedbackItems = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set receptionSheet = Nothing
    Set greenLightSheet = Nothing
    
    Set dictionaryOfFeedbackItems = Nothing
End Sub
