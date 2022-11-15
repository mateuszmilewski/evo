VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MalHandler"
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


Private mi As Outlook.mailItem
Private doKogo As String

Private app As Outlook.Application


Private preBody As String



Private Sub Class_Initialize()



    Set app = New Outlook.Application

    Set mi = Nothing
    doKogo = ""
    
    
    doKogo = CStr(defineDoKoGoList())
    
    
    preBody = "Dear all, new NOKs charts are waiting for feedback, " & Chr(10) & _
        "please check attached excelfile." & Chr(10)

End Sub


Private Function defineDoKoGoList() As String
    
    defineDoKoGoList = ""
    
    
    Dim r As Range
    Set r = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("B3")
    Do
        defineDoKoGoList = defineDoKoGoList & CStr(r) & " ; "
        Set r = r.offset(1, 0)
    Loop Until CStr(r) = vbNullString
End Function


Public Sub attachFileWithDataFromActiveSheet()


    If Not mi Is Nothing Then
    
        Dim kpiSh As Worksheet
        Set kpiSh = ActiveSheet
        
        If kpiSh.name Like "KPI_*" And kpiSh.Cells(2, 2).Value = "all data lines: " Then
        
            
            kpiSh.Copy
            
            Set kpiSh = ActiveSheet
            
            
            Dim wb As Workbook
            Set wb = ActiveWorkbook
            wb.Save
            
            Debug.Print "attachment from: " & wb.FullName
            mi.Attachments.Add wb.FullName
            
            DoEvents
            wb.Close False
        Else
            MsgBox "You cant add this sheet as attachment in your email!", vbCritical
        End If
        
        
        
        
    
        
    End If

End Sub

Public Sub displayMail()
    
    If Not mi Is Nothing Then mi.Display
    
End Sub


Public Sub procOnMailItem(subject As String, mBody As String)


    innerProc mi, subject, mBody, doKogo
    
End Sub


Private Sub innerProc(mailItem As Outlook.mailItem, mSubject As String, mBody As String, mDoKogo As String)


        
    If mi Is Nothing Then
        Set mi = app.CreateItem(olMailItem)
        
        With mi
            .subject = mSubject
            .body = preBody & mBody
            .Recipients.Add mDoKogo
            
            
        End With
    
    End If

End Sub

Private Sub Class_Terminate()
    
    Set mi = Nothing
    Set app = Nothing
    
End Sub
