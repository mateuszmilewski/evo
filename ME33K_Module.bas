Attribute VB_Name = "ME33K_Module"
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


Public Sub getDataFromME33K(ictrl As IRibbonControl)
    innerGetDataFromME33K
End Sub


Public Sub innerGetDataFromME33K()
Attribute innerGetDataFromME33K.VB_ProcData.VB_Invoke_Func = "M\n14"
    
    Dim sapHandler As New SAP_Handler
    
    sapHandler.runMainLogicFor__ME33K Nothing, Nothing, 0
    
End Sub




Public Sub innerGetDataFromME33K_2()
    
    Dim sapHandler As New SAP_Handler
    
    
    Dim outputFromInputBox As String
    outputFromInputBox = Trim(InputBox("DIV:"))
    
    sapHandler.runMainLogicFor__ME33K_mass Selection, outputFromInputBox
    
End Sub
