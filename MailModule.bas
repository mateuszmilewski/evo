Attribute VB_Name = "MailModule"
Public Sub runMail(ictrl As IRibbonControl)
    
    ' MsgBox "run mail - to be implemented!"
    
    Dim mh As MalHandler
    Set mh = New MalHandler
    
    mh.procOnMailItem "TEST", ""
    mh.displayMail
    
    Set mh = Nothing
End Sub
