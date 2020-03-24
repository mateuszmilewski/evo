Attribute VB_Name = "TestModule"
' testing overall status
Private Sub testMain()


    Dim sh As StatusHandler
    Set sh = New StatusHandler
    sh.init_statusbar 10
    sh.show
    For x = 1 To 10
        Sleep 1000
        sh.progress_increase
    Next x
    
    sh.hide
    
    Set sh = Nothing
    
End Sub
