Attribute VB_Name = "GlobalModule"
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



' IS_PRODUCTION
Global Const G_PROD = True

' DH COLUMNS !!!
Global Const G_DHEF_COL = 44
Global Const G_DHAS_COL = 45

Global Const G_COD_TRANSPORT_COLUMN = 36
Global Const G_PU_TIME_COLUMN = 37
Global Const G_T_TIME_COLUMN = 38
Global Const G_DEL_TIME_COLUMN = 39

Global Const G_TXT_IN_CELL = "transport by supplier (serial logistic in DAP)"
Global Const G_DAP = "DAP"
Global Const G_NON_TMC = "non"




' delay time
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



#If Win64 Then
  Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                      (ByVal IpBuffer As String, nSize As Long) As Long
  Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
                      (ByVal lpBuffer As String, nSize As Long) As Long

#Else
  Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                      (ByVal IpBuffer As String, nSize As Long) As Long
  Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
                      (ByVal lpBuffer As String, nSize As Long) As Long
#End If



Global Const REG_SH_NM = "register"
Global Const INPUT_SH_NM = "input"

Global Const MAIN_SH_FEED = "FICHERO TRANSFER ONL-MON"
Global Const MAIN_SH_BASE = "BASE"


