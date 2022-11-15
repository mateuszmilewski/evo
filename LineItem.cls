VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LineItem"
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

Public masterRng As Range
Public wiersz As Long
Public qty As Long

Public yyyycw As Long
Public monday As Date
Public isCodEntregaAvailable As Boolean
Public CodEntrega As String
Public dateAfterOffset As Date

Public delDate As Date
Public delTime As Date

Public puDate As Date
Public puTime As Date

Public tmc As String
Public tmc2 As String ' the last column -> Echéanciers regroupés par TMC - new implementation
Public bb As String ' should match with tmc2

Public pn As String
Public cofor As String
Public COFOR_COFOR As String

' from 0.16 for pivot source!
' ---------------------------
Public qty2 As Long
Public confirmedQty As Long
' ---------------------------

Public condi As String


Public sourceRow As Long
Public sourceRng As Range

Public found As Boolean
Public isPast As Boolean

' new fields connect directly with logic for creating
' pivot on data connected with
' TRANSPORT CALCULATION AND MONITORING
Public validationForPivotSource As Boolean
Public supplementForPivotSource As LineItemPivotSrouceSupplement


Private log As String



' ========================================================
Public Function getFullLog() As String
    
    getFullLog = CStr(log)
End Function


Public Sub addRawInfoToLog(info As String)
    
    log = log & CStr(info) & Chr(10)
End Sub
' ========================================================


Private Sub Class_Initialize()
    isCodEntregaAvailable = False
    log = ""
End Sub

Private Sub Class_Terminate()
    isCodEntregaAvailable = False
End Sub
