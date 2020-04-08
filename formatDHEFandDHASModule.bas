Attribute VB_Name = "formatDHEFandDHASModule"
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



Public Sub formatDHEFandDHAS(ictrl As IRibbonControl)
    
    innerformatDHEFandDHAS
    MsgBox "FORMATOWANIE GOTOWE!", vbInformation
End Sub



Public Sub innerformatDHEFandDHAS()


    ' -------------------------------------------------------------------------------
    ' -------------------------------------------------------------------------------
    ' should be after validation
    
    Dim strm1 As String, iterDate As Date
    strm1 = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M1").Value
    
    Dim wrk As Workbook
    
    Set wrk = Nothing
    
    On Error Resume Next
    Set wrk = Workbooks(strm1)
    
    
    If Not wrk Is Nothing Then
        
        Dim sh As Worksheet
        Set sh = Nothing
        
        On Error Resume Next
        Set sh = wrk.Sheets("BASE")
        
        If Not sh Is Nothing Then
            Dim r As Range, dh As Range
            Set r = sh.Cells(3, 1)
            Do
            
                Set dh = r.Offset(0, EVO.G_DHEF_COL - 1)
                
                If IsDate(dh) Then
                    
                    ' dh.Value = CStr(Trim(Format(dh, "dd/mm/yyyy  hh:mm:ss")))
                    iterDate = CDate(dh)
                    dh.Value = CStr(parseStdDateToStrangeFormat(iterDate))
                    
                End If
                
                If IsDate(dh.Offset(0, 1)) Then
                
                    ' dh.Offset(0, 1).Value = CStr(Trim(Format(dh.Offset(0, 1), "dd/mm/yyyy  hh:mm:ss")))
                    iterDate = CDate(dh.Offset(0, 1))
                    dh.Offset(0, 1).Value = CStr(parseStdDateToStrangeFormat(iterDate))
                End If
            
            
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
        Else
            MsgBox "Wrong file!", vbCritical
        End If
    Else
        MsgBox "Please re-check you pus master file!", vbCritical
    End If
    ' -------------------------------------------------------------------------------
    ' -------------------------------------------------------------------------------

End Sub


Private Function parseStdDateToStrangeFormat(iterDate As Date) As String

    parseStdDateToStrangeFormat = ""
    
    
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    Dim hh As String
    Dim mmm As String
    Dim sS As String
    
    yyyy = CStr(Year(iterDate))
    mm = CStr(Month(iterDate))
    
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
    
    dd = CStr(Day(iterDate))
    
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    
    hh = CStr(Hour(iterDate))
    
    If Len(hh) = 1 Then
        hh = "0" & hh
    End If
    
    
    mmm = CStr(Minute(iterDate))
    
    If Len(mmm) = 1 Then
        mmm = "0" & mmm
    End If
    
    sS = CStr(Second(iterDate))
    
    If Len(sS) = 1 Then
        sS = "0" & sS
    End If
    
    parseStdDateToStrangeFormat = "" & CStr(dd) & "/" & CStr(mm) & "/" & CStr(yyyy) & "  " & _
        CStr(hh) & ":" & CStr(mmm) & ":" & CStr(sS)
    
End Function
