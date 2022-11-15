VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CopyIterationForPivotSource"
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



' TYPE_DE_PIECE handler for offset
Private diffE As Integer
Private diff_for_cofor_hybrid As Integer
Private diff_for_incoterm As Integer

Public Sub D_copyIterationForPivotSource(master As Worksheet, sh As Worksheet, ByRef li As LineItem, ByRef ph As PivotHandler, linia As Long)




    ' uwaga!
    ' ta iteracja ma pomieszane kolejnosci kolumn totalnie
    ' wiec ostatnie zmiany niekoniecznie znajduja sie na koncu tego suba!

    

    
    'Debug.Print "li.supplementForPivotSource.refRangeForCloe.Address: " & _
    '    li.supplementForPivotSource.refRangeForCloe.Address
    
    'Debug.Print "li.supplementForPivotSource.refRangeForPLE.Address: " & _
    '    li.supplementForPivotSource.refRangeForPLE.Address
        
    'Debug.Print "li.sourceRng.Address: " & _
    '    li.sourceRng.Address
    
    
    ' ID LATER AFTER DEL DATE!
    
    sh.Cells(linia, EVO.E_PIVOT_SRC_WIERSZ).Value = li.wiersz
    sh.Cells(linia, EVO.E_PIVOT_SRC_REF).Value = li.pn
    sh.Cells(linia, EVO.E_PIVOT_SRC_COFOR_VENDEUR).Value = Split(li.COFOR_COFOR, "_")(0)
    sh.Cells(linia, EVO.E_PIVOT_SRC_COFOR_EXPEDITEUR).Value = li.cofor
    
    
    If Not li.supplementForPivotSource.refRangeForPLE Is Nothing Then
        sh.Cells(linia, EVO.E_PIVOT_SRC_SUPPLIER_NAME).Value = _
            li.supplementForPivotSource.refRangeForPLE.offset(0, -2).Value
    Else
        
        
    End If
        
        

    
    ' Debug.Print li.sourceRng.Offset(0, diffE).Value ' OK!
    ' -------------------------------------------------------------------------------
    ' -------------------------------------------------------------------------------
    
    ' this is some extra data based on sourceRng
    
    sh.Cells(linia, EVO.E_PIVOT_SRC____TYPE_DE_PIECE).Value = _
        li.sourceRng.offset(0, diffE).Value
        
    ' based on sourceRng starting from 0.8x
    sh.Cells(linia, EVO.E_PIVOT_SRC__lb_COFOR_HYBRID).Value = _
        li.sourceRng.offset(0, diff_for_cofor_hybrid).Value
        
    sh.Cells(linia, EVO.E_PIVOT_SRC__lb_INCOTERM).Value = _
        li.sourceRng.offset(0, diff_for_incoterm).Value
        
    ' -------------------------------------------------------------------------------
    ' -------------------------------------------------------------------------------
        
    ' DEL DATE
    If IsDate(li.delDate) Then
        If CLng(li.delDate) <> 0 Then
    
            sh.Cells(linia, EVO.E_PIVOT_SRC_DELIVERY_DATE).Value = li.delDate
            sh.Cells(linia, EVO.E_PIVOT_SRC_DELIVERY_YEAR).Value = Year(li.delDate)
            sh.Cells(linia, EVO.E_PIVOT_SRC_DELIVERY_MONTH).Value = Month(li.delDate)
            sh.Cells(linia, EVO.E_PIVOT_SRC_DELIVERY_WEEK).Value = _
                Application.WorksheetFunction.IsoWeekNum(li.delDate)
                
            sh.Cells(linia, EVO.E_PIVOT_SRC_DELIVERY_YYYYCW).Value = CLng(Year(li.delDate)) * 10000 + _
                CLng(sh.Cells(linia, EVO.E_PIVOT_SRC_DELIVERY_WEEK).Value) * 100 + CLng(Weekday(li.delDate, vbMonday))
                
                
            ' here quickly short date
            sh.Cells(linia, EVO.E_PIVOT_SRC___SHORT_DEL_DATE).Value = Format(li.delDate, "dd/mm/yy")
            ' new in 0.8x
            sh.Cells(linia, EVO.E_PIVOT_SRC__lb_DHAS).Value = li.delDate
                
        End If
    End If
    
    ' PU DATE
    If IsDate(li.puDate) Then
        If CLng(li.puDate) <> 0 Then
            sh.Cells(linia, EVO.E_PIVOT_SRC__lb_DHEF).Value = li.puDate
        End If
    End If
    
    ' for order date calc
    ' -------------------------------------------------------------------------------------------------
    sh.Cells(linia, EVO.E_PIVOT_SRC_AK_COL_FROM_PLE).Value = li.supplementForPivotSource.heoDecalageJX
    ' -------------------------------------------------------------------------------------------------
    
    ' ID !!!
    ' --------------
    ' --------------
    ' -------------------------------------------------------------------------------------------------
    sh.Cells(linia, EVO.E_PIVOT_SRC_ID).Value = _
        CStr(li.pn) & _
        "-" & CStr(li.cofor) & _
        "-" & CStr(sh.Cells(linia, EVO.E_PIVOT_SRC_DELIVERY_YYYYCW).Value)
    ' -------------------------------------------------------------------------------------------------
    
    
    If IsDate(li.supplementForPivotSource.orderDate) Then
    
        If CLng(li.supplementForPivotSource.orderDate) <> 0 Then
    
            sh.Cells(linia, EVO.E_PIVOT_SRC_ORDER_DATE).Value = _
                li.supplementForPivotSource.orderDate
            sh.Cells(linia, EVO.E_PIVOT_SRC_ORDER_YEAR).Value = _
                Year(li.supplementForPivotSource.orderDate)
            sh.Cells(linia, EVO.E_PIVOT_SRC_ORDER_MONTH).Value = _
                Month(li.supplementForPivotSource.orderDate)
            sh.Cells(linia, EVO.E_PIVOT_SRC_ORDER_WEEK).Value = _
                Application.WorksheetFunction.IsoWeekNum(li.supplementForPivotSource.orderDate)
                
                
            sh.Cells(linia, EVO.E_PIVOT_SRC_ORDER_YYYYCW).Value = _
                CLng(Year(li.supplementForPivotSource.orderDate)) * 100 + _
                CLng(sh.Cells(linia, EVO.E_PIVOT_SRC_ORDER_WEEK).Value)
                
                
                
            ' here quickly short date
            sh.Cells(linia, EVO.E_PIVOT_SRC___SHORT_ORDER_DATE).Value = _
                Format(li.supplementForPivotSource.orderDate, "dd/mm/yy")
                
            
            ' new in 0.8x
            sh.Cells(linia, EVO.E_PIVOT_SRC__lb_DHEO).Value = _
                li.supplementForPivotSource.orderDate
                
        End If
            
            
            
    End If
    
    
    'E_PIVOT_SRC_ROUTE
    'E_PIVOT_SRC_PILOT
    'E_PIVOT_SRC_ROUTE_AND_PILOT
    sh.Cells(linia, EVO.E_PIVOT_SRC_ROUTE).Value = _
        li.supplementForPivotSource.refRangeForCloe.offset(0, 2).Value
    sh.Cells(linia, EVO.E_PIVOT_SRC_PILOT).Value = _
        li.supplementForPivotSource.refRangeForCloe.offset(0, 3).Value
        
    ' CLOE PRICE
    ' -------------------------------------------------------------------------------------------------------------
    ' -------------------------------------------------------------------------------------------------------------
    sh.Cells(linia, EVO.E_PIVOT_SRC____CLOE_COL_E_PRICE).Value = 0
    If Application.WorksheetFunction.IsErr(li.supplementForPivotSource.refRangeForCloe.offset(0, 4).Value) Then
        On Error Resume Next
        sh.Cells(linia, EVO.E_PIVOT_SRC____CLOE_COL_E_PRICE).Value = 0
    ElseIf IsNumeric(li.supplementForPivotSource.refRangeForCloe.offset(0, 4).Value) Then
        sh.Cells(linia, EVO.E_PIVOT_SRC____CLOE_COL_E_PRICE).Value = _
            li.supplementForPivotSource.refRangeForCloe.offset(0, 4).Value
    Else
        On Error Resume Next
        sh.Cells(linia, EVO.E_PIVOT_SRC____CLOE_COL_E_PRICE).Value = 0
    End If
    ' -------------------------------------------------------------------------------------------------------------
    ' -------------------------------------------------------------------------------------------------------------
        
        
    sh.Cells(linia, EVO.E_PIVOT_SRC_ROUTE_AND_PILOT).Value = "" & _
        li.supplementForPivotSource.refRangeForCloe.offset(0, 2).Value & _
        " - " & _
        li.supplementForPivotSource.refRangeForCloe.offset(0, 3).Value
        
    ' again splitted data - need to see it
    ' new in 0.8x
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    sh.Cells(linia, EVO.E_PIVOT_SRC__lb_PILOT).Value = _
        li.supplementForPivotSource.refRangeForCloe.offset(0, 3).Value
    sh.Cells(linia, EVO.E_PIVOT_SRC__lb_ROUTE).Value = _
        li.supplementForPivotSource.refRangeForCloe.offset(0, 2).Value
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        
    ' QTY
    ' -------------------------------------------------------------------
    sh.Cells(linia, EVO.E_PIVOT_SRC_QTY).Value = li.qty2
    sh.Cells(linia, EVO.E_PIVOT_SRC_CQTY).Value = li.confirmedQty
    ' -------------------------------------------------------------------
    
    
    ' HERE MASTER!
    ' ======================================================================================
    sh.Cells(linia, EVO.E_PIVOT_SRC_UC).Value = master.Cells(li.wiersz, EVO.G_UC_COL).Value
    sh.Cells(linia, EVO.E_PIVOT_SRC_SUMIF_UC).Value = master.Cells(li.wiersz, EVO.G_UC_COL).Value
    ' ======================================================================================
    
    
    ' OQ - QTY per UC
    ' -------------------------------------------------------------------
    If IsNumeric(master.Cells(li.wiersz, EVO.G_UC_COL).Value) Then
        If CLng(master.Cells(li.wiersz, EVO.G_UC_COL).Value) > 0 Then
        
        
            sh.Cells(linia, EVO.E_PIVOT_SRC_OQ).Value = _
                CDbl(li.qty2) / CDbl(CLng(master.Cells(li.wiersz, EVO.G_UC_COL).Value))
                
            sh.Cells(linia, EVO.E_PIVOT_SRC_COQ).Value = _
                CDbl(li.confirmedQty) / CDbl(CLng(master.Cells(li.wiersz, EVO.G_UC_COL).Value))
                
            sh.Cells(linia, EVO.E_PIVOT_SRC_ROUNDUP_OQ1).Value = _
                CDbl(Application.WorksheetFunction.RoundUp(CDbl(sh.Cells(linia, EVO.E_PIVOT_SRC_OQ).Value), 0))
                
            sh.Cells(linia, EVO.E_PIVOT_SRC_ROUNDUP_COQ1).Value = _
                CDbl(Application.WorksheetFunction.RoundUp(CDbl(sh.Cells(linia, EVO.E_PIVOT_SRC_COQ).Value), 0))
            
        End If
    End If
    ' -------------------------------------------------------------------
    
    ' Debug.Assert li.wiersz <> 2866
    
    If CStr(li.condi) <> "" Then
    
        sh.Cells(linia, EVO.E_PIVOT_SRC_CONDI).Value = li.condi
        
        'E_PIVOT_SRC_UA_PC_GV
        'E_PIVOT_SRC_UA_BPC
        'E_PIVOT_SRC_UA_MC
        'E_PIVOT_SRC_UA_MBU
        'E_PIVOT_SRC_UA_MAX_CAPACITY
        
        sh.Cells(linia, EVO.E_PIVOT_SRC_UA_PC_GV).Value = _
            li.supplementForPivotSource.refRaneForCondi.offset(0, 1).Value
        sh.Cells(linia, EVO.E_PIVOT_SRC_UA_BPC).Value = _
            li.supplementForPivotSource.refRaneForCondi.offset(0, 2).Value
        sh.Cells(linia, EVO.E_PIVOT_SRC_UA_MC).Value = _
            li.supplementForPivotSource.refRaneForCondi.offset(0, 3).Value
        sh.Cells(linia, EVO.E_PIVOT_SRC_UA_MBU).Value = _
            li.supplementForPivotSource.refRaneForCondi.offset(0, 4).Value
        sh.Cells(linia, EVO.E_PIVOT_SRC_UA_MAX_CAPACITY).Value = _
            li.supplementForPivotSource.refRaneForCondi.offset(0, 5).Value

    End If
    

        
        
End Sub

Private Sub Class_Initialize()
    ' TYPE_DE_PIECE
    diffE = EVO.E_MASTER_2_TYPE_DE_PIECE - EVO.E_MASTER_2_COFOR_EXPEDITEUR
    diff_for_cofor_hybrid = EVO.E_MASTER_2_COFOR_HIBRIDE - EVO.E_MASTER_2_COFOR_EXPEDITEUR
    diff_for_incoterm = EVO.E_MASTER_2_Incoterm - EVO.E_MASTER_2_COFOR_EXPEDITEUR
End Sub
