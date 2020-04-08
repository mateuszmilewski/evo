Attribute VB_Name = "DocInfoModule"
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



' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


' Hi Forrest from the future - this module is not working anymore - but it can
' Im leaving this implementation as legacy just in case!


' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


'
'Public Sub catchDocInfoFiles(ictrl As IRibbonControl)
'
'    innerCatchDocInfoFiles
'
'End Sub
'
'
'
'
'Public Sub verifyDocInfoFiles(ictrl As IRibbonControl)
'
'
'    innerVerifyDocInfoFiles
'
'End Sub



Public Sub innerCatchDocInfoFiles()




    Dim logMsg As String
    logMsg = ""
    Dim regSh As Worksheet
    Set regSh = ThisWorkbook.Sheets(EVO.REG_SH_NM)
    
    
    Application.DisplayAlerts = False
    
    
    Dim externalWorkbook As Workbook
    Set externalWorkbook = Nothing
    
    
    Dim r As Range
    
    If G_PROD Then
    
        logMsg = logMsg & "YOU ARE IN PROD! " & Chr(10) & Chr(10)
    
        Set r = regSh.Range("N2")
        Do
        
            r.Offset(0, -1).Value = ""
            r.Offset(0, -2).Value = ""
        
            If Trim(r.Value) Like "*docinfogroupe*" Then
            
                
                
                'clearly this is an file from docinfogroupe
                Set externalWorkbook = Workbooks.Open((CStr(r.Value)), True, True)
                
                Do
                    DoEvents
                Loop While externalWorkbook Is Nothing
                
                r.Offset(0, -2).Value = externalWorkbook.Path
                r.Offset(0, -1).Value = externalWorkbook.FullName
                
                logMsg = logMsg & externalWorkbook.FullName & ", " & Chr(10)
                
                
            End If
            
            Set r = r.Offset(1, 0)
            
        Loop Until Trim(r) = ""
    Else
    
        logMsg = logMsg & "YOU ARE IN PRE-PROD! " & Chr(10) & Chr(10)
    
        ' pre-prod !!!
        Set r = regSh.Range("T2")
        Do
        
            r.Offset(0, -1).Value = ""
            r.Offset(0, -2).Value = ""
        
            If Trim(r.Value) Like "C:\*" Then
            
                
                
                'clearly this is an file from docinfogroupe
                Set externalWorkbook = Workbooks.Open((CStr(r.Value)), True, True)
                
                Do
                    DoEvents
                Loop While externalWorkbook Is Nothing
                
                r.Offset(0, -2).Value = externalWorkbook.Path
                r.Offset(0, -1).Value = externalWorkbook.FullName
                
                logMsg = logMsg & externalWorkbook.FullName & ", " & Chr(10)
                
                
            End If
            
            Set r = r.Offset(1, 0)
            
        Loop Until Trim(r) = ""
    
    
    End If
    
    
    Set regSh = Nothing
    
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    
    Dim w As Workbook
    For Each w In Workbooks
        
        On Error Resume Next
        Debug.Print CStr(w.Name) & " " & Chr(10) & _
            CStr(w.FullName) & " " * Chr(10) & _
            CStr(w.Path) & _
            Chr(10) & Chr(10)
    Next w
    
    logMsg = logMsg & " imported and ready to work!"
    ' ---------------------------------------------------
    MsgBox logMsg
    ' ---------------------------------------------------


End Sub



Public Sub innerVerifyDocInfoFiles()


    Dim r As Range, wrk As Workbook, iterWrk As Workbook, logMsg As String
    logMsg = ""

    If EVO.G_PROD Then
        logMsg = logMsg & "YOU ARE IN PROD! " & Chr(10) & Chr(10) & Chr(10)
        Set r = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M2")
    Else
        
        logMsg = logMsg & "YOU ARE IN PRE-PROD! " & Chr(10) & Chr(10) & Chr(10)
        Set r = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("S2")
    End If
        
        
    Do
        Set wrk = Nothing
        On Error Resume Next
        
        For Each iterWrk In Workbooks
            
            If Trim(iterWrk.FullName) = Trim(r.Value) Then
                Set wrk = iterWrk
                Exit For
            End If
        Next iterWrk
        
        
        If Not wrk Is Nothing Then
        
            verifyThisFile wrk, CStr(r.Offset(0, 2)), logMsg
        Else
            MsgBox "Macro stops! sth wrong with the: " & r.Value & " please verify if it is available or process", vbCritical
            End
        End If
        
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    
    ' ---------------------------------------------------
    MsgBox logMsg, vbExclamation
    ' ---------------------------------------------------
        
End Sub


Private Sub verifyThisFile(w As Workbook, typeOfFile As String, ByRef logMsg As String)


    Dim sh As Worksheet, toBeTrue As Boolean
    
    Set sh = Nothing

    If typeOfFile = "feed" Then
        
        ' check if this file is really feeding file
        ' --------------------------------------------------
        
        toBeTrue = True
        
        ' check availability of main worksheet
        Set sh = Nothing
        On Error Resume Next
        Set sh = w.Sheets("FICHERO TRANSFER ONL-MON")
        
        ' such sheet is available
        If Not sh Is Nothing Then
        
            ' wacky check on 3 columns labels
            If Trim(sh.Cells(1, 1).Value) Like "Num?ro produit" Then
                If Trim(sh.Range("I1").Value) Like "D?signation longue" Then
                
                    ' and some from far away
                    If Trim(sh.Range("BX1").Value) = "DA COFOR VENDEDOR" Then
                    Else
                        toBeTrue = False
                    End If
                Else
                    toBeTrue = False
                End If
            Else
                toBeTrue = False
            End If

        
        Else
            toBeTrue = False
        End If
        
        ' --------------------------------------------------
        
    ElseIf typeOfFile = "master" Then
    
        ' check if this file is really master file
        ' --------------------------------------------------
        
        
        toBeTrue = True
        
        ' check availability of main worksheet
        Set sh = Nothing
        On Error Resume Next
        Set sh = w.Sheets("BASE")
        
        ' such sheet is available
        If Not sh Is Nothing Then
        
            ' wacky check on 3 columns labels
            If sh.Cells(1, 1).Value = "ONL" Then
                
                If sh.Cells(2, 1).Value = "REFERENCE" Then
                
                    ' and some from far away
                    If Trim(sh.Range("AR2").Value) = "DHEF" Then
                    Else
                        toBeTrue = False
                    End If
                Else
                    toBeTrue = False
                End If
            Else
                toBeTrue = False
            End If

        
        Else
            toBeTrue = False
        End If
        
        
        ' --------------------------------------------------
        
    End If
    
    
    If toBeTrue Then
        logMsg = logMsg & " initial validation OK! File: " & _
            w.FullName & " is inline with type of: " & typeOfFile & Chr(10) & Chr(10)
    Else
        logMsg = logMsg & " initial validation NOK! File: " & _
            w.FullName & " is not a std type of: " & typeOfFile & " please re-check your input data " & Chr(10) & Chr(10)
    End If

End Sub
