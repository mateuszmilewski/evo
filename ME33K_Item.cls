VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ME33K_Item"
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


' TANGO_TP04
Public pos As String
Public article As String
Public nm As String
Public qty As String

Public ua As String
Public up As String

Public strPrice As String
Public price As Double
Public curr As String

Public Sub makePriceByUnit()
    
    If up <> "" Then
        
        If EVO.calcUnSpecial(up) > 0 Then
            price = price / EVO.calcUnSpecial(up)
        End If
    End If
End Sub

Public Sub parsePrice()

    price = 0#
    
    On Error Resume Next
    price = CDbl(strPrice)
    
    If price = 0 Then
        
        If IsNumeric(strPrice) Then
            Debug.Print "is numeric, but it is not working!"
            
            
        Else
            
            Debug.Print "is not numeric!"
            
            If CStr(strPrice) Like "*.???,??" Then
                
                strPrice = Replace(strPrice, ".", "")
                On Error Resume Next
                price = CDbl(strPrice)
                
            End If
        End If
    End If
End Sub
