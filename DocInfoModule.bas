Attribute VB_Name = "DocInfoModule"
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



Public Sub catchDocInfoFiles(ictrl As IRibbonControl)


    Dim logMsg As String
    logMsg = ""
    Dim regSh As Worksheet
    Set regSh = ThisWorkbook.Sheets(EVO.REG_SH_NM)
    
    
    Application.DisplayAlerts = False
    
    
    Dim externalWorkbook As Workbook
    Set externalWorkbook = Nothing
    
    
    Dim r As Range
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
    
    
    Set regSh = Nothing
    
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    
    Dim w As Workbook
    For Each w In Workbooks
        
        Debug.Print w.Name & " " & Chr(10) & _
            w.FullName & " " * Chr(10) & _
            w.Path & _
            Chr(10) & Chr(10) & Chr(10)
    Next w
    
    logMsg = logMsg & " imported and ready to work!"
    ' ---------------------------------------------------
    MsgBox logMsg
    ' ---------------------------------------------------

End Sub

Public Sub verifyDocInfoFiles(ictrl As IRibbonControl)
    MsgBox "to be implemented!"
End Sub
