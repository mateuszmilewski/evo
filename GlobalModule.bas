Attribute VB_Name = "GlobalModule"
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



' for status handler
Global Const INITIAL_TIMING_FOR_ONE_PN = 1


' IS_PRODUCTION
Global Const G_PROD = True


' QTY and CQTY
Global Const G_QTY_COL = 3
Global Const G_CONF_QTY_COL = 4


' DH COLUMNS !!!
Global Const G_DHEF_COL = 44
Global Const G_DHAS_COL = 45

' COFOR
Global Const G_COFOR_VENDEUR_COL = 17
Global Const G_COFOR_EXPEDITEUR_COL = 18

' pack and UC
Global Const G_CONDI_COL = 24
Global Const G_UC_COL = 25

Global Const G_COD_TRANSPORT_COLUMN = 36
Global Const G_PU_TIME_COLUMN = 37
Global Const G_T_TIME_COLUMN = 38
Global Const G_DEL_TIME_COLUMN = 39
Global Const G_SRC_DHEF_COL = 30
Global Const G_SRC_DHAS_COL = 31

Global Const G_PLE_SUB_FOR_ORDER_COLUMN = 37

Global Const G_TXT_IN_CELL = "transport by supplier (serial logistic in DAP)"
Global Const G_TXT_IN_CELL_II = "transport by supplier"
Global Const G_DAP = "DAP"
Global Const G_DDP = "DDP"
Global Const G_NON_TMC = "non"



' column "G"
Global Const G_PLE_VEN_COFOR = 6
Global Const G_PLE_SHIPPER_COFOR = 7
Global Const G_CLOE_SHIPPER_COFOR = 1
Global Const G_CLOE_COFORS = 2
Global Const G_UA_KEY = 1
Global Const G_UA_MAX_CAPACITY_COLUMN = 6
Global Const G_FEED_MAIN_SH_VEN_COFOR = 3
Global Const G_FEED_MAIN_SH_SHIPPER_COFOR = 4
Global Const G_FEED_MAIN_SH_CONDI = 10



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

' Global Const MAIN_SH_FEED = "FICHERO TRANSFER ONL-MON"
Global Const MAIN_SH_BASE = "BASE"


Global Const G_FEED_SH_MAIN = "MAIN"
Global Const G_FEED_SH_PLE = "PLE"
Global Const G_FEED_SH_CLOE = "CLOE"
Global Const G_FEED_SH_UA = "UA"


' TP04
Global Const G_TP04_TP04_01 = "TP04_"

' SQ01
Global Const G_REF_MOUNT_SQ1_OUT = "D14"
Global Const G_REF_MOUNT_N_SUPPLIERS_OUT = "D17"


Global Const G_COL_IS_INTERNAL_GREEN_LIGHT = 21
Global Const G_COL_IS_TANGO_GREEN_LIGHT = 22




Public Function calcUnSpecial(param As Variant) As Double
    calcUnSpecial = 1#
    
    Dim regRef As Range
    Set regRef = ThisWorkbook.Sheets("register").Range("UN_REF")
    
    Do
        If CStr(regRef.Value) = CStr(param) Then
            calcUnSpecial = CDbl(regRef.offset(0, 1).Value)
            Exit Do
        End If
        Set regRef = regRef.offset(1, 0)
    Loop Until Trim(regRef.Value) = ""
End Function

