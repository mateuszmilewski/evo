VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MalHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mi As Outlook.mailItem
Private doKogo As String

Private app As Outlook.Application


Private preBody As String



Private Sub Class_Initialize()



    Set app = New Outlook.Application

    Set mi = Nothing
    doKogo = ""
    
    
    doKogo = CStr(defineDoKoGoList())
    
    
    preBody = "Dear XXX team, new NOKs are waiting for feedback, " & Chr(10) & _
        "please visit: URL" & Chr(10)

End Sub


Private Function defineDoKoGoList() As String
    
    defineDoKoGoList = ""
    
    
    Dim r As Range
    Set r = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("B3")
    Do
        defineDoKoGoList = defineDoKoGoList & CStr(r) & " ; "
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = vbNullString
End Function

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
