Attribute VB_Name = "ReadMeButtonsModule"
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

' quick guide on green light
Sub Button1_Click()
    Debug.Print " Welcome in read me! "
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollRow = 100
End Sub

' quick guide on reception
Sub Button2_Click()
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollRow = 200
End Sub

Sub goToDoc()
    
    Dim ie As InternetExplorer
    Set ie = New InternetExplorer
    ie.Visible = True
    ie.Navigate "http://docinfogroupe.inetpsa.com/ead/doc/ref.01817_20_01308/v.vc/fiche"
End Sub


Public Sub goBackToTheTop()
    ActiveWindow.ScrollRow = 1
End Sub
