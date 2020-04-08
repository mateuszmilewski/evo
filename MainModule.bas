Attribute VB_Name = "MainModule"
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

Public Sub reset()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub



Public Sub runMain(ictrl As IRibbonControl)
    
    ' ---------------------------------------------
    
    Debug.Print "start -> run main"
    main
    
    ' ---------------------------------------------
    
End Sub





' here is the main logic for whole tool
Private Sub main()


    Dim sh As StatusHandler
    Set sh = New StatusHandler
    
    ' steps
    ' -------------------------------------
    ' obsolete
    'innerCatchDocInfoFiles
    ' obsolete
    'innerVerifyDocInfoFiles
    'innerCopyData
    
    MsgBox "to be implemented!"
    ' -------------------------------------
    
    Set sh = Nothing
    
End Sub
