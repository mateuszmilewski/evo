Attribute VB_Name = "SQ01MassModule"
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


' concatenate domains (the output worksheet)
' ==============================================





Public Sub mass5__MainForSq01()



    ' 3 stands for 3rd version of the managers da - look for everything - stop making input list!


    KillMessageFilter
    
    Dim list As Range ' , listItem As Range
    Set list = ThisWorkbook.Sheets("register").Range("AD2")
    
    Dim out1 As Worksheet, out2 As Worksheet, resultSh As Worksheet
    
    Dim sap4out As SAP_Handler
    
    Set sap4out = Nothing
    
    
    Do
    
        If Left(list.Value, 1) = "F" Then
            innerMassItemForSq01_WO_IN_LIST list, 2, out1, sap4out
            innerMassItemForSq01_WO_IN_LIST list, 3, out2, sap4out
            
            ' be careful is from SQ01Module directly
            
            concatAndStd out1, out2, resultSh, sap4out
            
            
            ' managers DA
            resultSh.Activate
            innerGetManagersDa3 resultSh
        End If
    
        Set list = list.offset(1, 0)
    
    Loop Until Trim(list.Value) = ""
    
    
    EVO.RestoreMessageFilter
    
End Sub







' BUT STILL NOK TAKING ONLY ONE DIV - work for nothing
Public Sub mass4__MainForSq01_ONEITEMALL()



    ' 3 stands for 3rd version of the managers da - look for everything - stop making input list!


    KillMessageFilter
    
    Dim list As Range ' , listItem As Range
    Set list = ThisWorkbook.Sheets("register").Range("AD2")
    
    Dim out1 As Worksheet, out2 As Worksheet, resultSh As Worksheet
    
    Dim sap4out As SAP_Handler
    
    Set sap4out = Nothing
    
    
    Do
    
        If Left(list.Value, 1) = "F" Then
            innerMassItemForSq01_WO_IN_LIST_2 list, 2, out1, sap4out
            ' innerMassItemForSq01_WO_IN_LIST_2 list, 3, out2, sap4out
            
            ' be careful is from SQ01Module directly
            concatAndStd out1, out2, resultSh, sap4out
            
            
            ' managers DA
            'resultSh.Activate
            'innerGetManagersDa3 resultSh
            
            Exit Do
        End If
    
        Set list = list.offset(1, 0)
    
    Loop Until Trim(list.Value) = ""
    
    
    EVO.RestoreMessageFilter
    
End Sub






Public Sub mass3_testOnlyFor_managersDa3()
    
    innerGetManagersDa3 ActiveSheet
End Sub


Public Sub mass3__MainForSq01()



    ' 3 stands for 3rd version of the managers da - look for everything - stop making input list!


    KillMessageFilter
    
    Dim list As Range ' , listItem As Range
    Set list = ThisWorkbook.Sheets("register").Range("AD2")
    
    Dim out1 As Worksheet, out2 As Worksheet, resultSh As Worksheet
    
    Dim sap4out As SAP_Handler
    
    Set sap4out = Nothing
    
    
    Do
    
        If Left(list.Value, 1) = "F" Then
            innerMassItemForSq01_WO_IN_LIST list, 2, out1, sap4out
            innerMassItemForSq01_WO_IN_LIST list, 3, out2, sap4out
            
            ' be careful is from SQ01Module directly
            concatAndStd out1, out2, resultSh, sap4out
            
            
            ' managers DA
            resultSh.Activate
            innerGetManagersDa3 resultSh
        End If
    
        Set list = list.offset(1, 0)
    
    Loop Until Trim(list.Value) = ""
    
    
    EVO.RestoreMessageFilter
    
End Sub





Public Sub mass2__MainForSq01()


    KillMessageFilter
    
    Dim list As Range ' , listItem As Range
    Set list = ThisWorkbook.Sheets("register").Range("AD2")
    
    Dim out1 As Worksheet, out2 As Worksheet, resultSh As Worksheet
    
    Dim sap4out As SAP_Handler
    
    Set sap4out = Nothing
    
    
    Do
    
        If Left(list.Value, 1) = "F" Then
            innerMassItemForSq01_WO_IN_LIST list, 2, out1, sap4out
            innerMassItemForSq01_WO_IN_LIST list, 3, out2, sap4out
            
            ' be careful is from SQ01Module directly
            concatAndStd out1, out2, resultSh, sap4out
            
            
            ' managers DA
            resultSh.Activate
            innerGetManagersDa resultSh
        End If
    
        Set list = list.offset(1, 0)
    
    Loop Until Trim(list.Value) = ""
    
    
    EVO.RestoreMessageFilter
    
End Sub

    

Public Sub ONE_MainForSq01_forSelection()


    KillMessageFilter
    
    Dim list As Range
    Set list = Selection
    
    Dim out1 As Worksheet, out2 As Worksheet, resultSh As Worksheet
    
    Dim sap4out As SAP_Handler
    
    Set sap4out = Nothing
    

    If list.Value = "F" Then
        innerMassItemForSq01 list, 2, out1, sap4out
        innerMassItemForSq01 list, 3, out2, sap4out
        
        ' be careful is from SQ01Module directly
        concatAndStd out1, out2, resultSh, sap4out
        
        
        ' managers DA
        resultSh.Activate
        innerGetManagersDa resultSh
    End If
    
    
    EVO.RestoreMessageFilter
    
End Sub


Public Sub ONE_MainForSq01_forSelection_WO_IN_LIST()


    KillMessageFilter
    
    Dim list As Range
    Set list = Selection
    
    Dim out1 As Worksheet, out2 As Worksheet, resultSh As Worksheet
    

    If list.Value = "F" Then
        innerMassItemForSq01_WO_IN_LIST list, 2, out1
        innerMassItemForSq01_WO_IN_LIST list, 3, out2
        
        ' be careful is from SQ01Module directly
        concatAndStd out1, out2, resultSh
        
        
        ' managers DA
        resultSh.Activate
        innerGetManagersDa resultSh
    End If
    
    
    EVO.RestoreMessageFilter
    
End Sub

Public Sub ONE_MainForSq01()
    
    Dim list As Range
    Set list = ThisWorkbook.Sheets("register").Range("AD11")
    
    
    Dim out1 As Worksheet, out2 As Worksheet, resultSh As Worksheet
    

    If list.Value = "F" Then
        innerMassItemForSq01 list, 2, out1
        innerMassItemForSq01 list, 3, out2
        
        ' be careful is from SQ01Module directly
        concatAndStd out1, out2, resultSh
        
        
        ' managers DA
        resultSh.Activate
        innerGetManagersDa resultSh
    End If
    
End Sub

Public Sub massMainForSq01()
    
    Dim list As Range
    Set list = ThisWorkbook.Sheets("register").Range("AD2")
    
    
    Dim out1 As Worksheet, out2 As Worksheet, resultSh As Worksheet
    
    Do
        If list.Value = "F" Then
            innerMassItemForSq01 list, 2, out1
            innerMassItemForSq01 list, 3, out2
            
            ' be careful is from SQ01Module directly
            concatAndStd out1, out2, resultSh
            
            
            ' managers DA
            resultSh.Activate
            innerGetManagersDa resultSh
        End If
        Set list = list.offset(1, 0)
    Loop Until Trim(list.Value) = ""
End Sub





Public Sub innerMassItemForSq01(r As Range, offst As Integer, ByRef osh As Worksheet, Optional ByRef sap4out As SAP_Handler)


    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    ' Dim osh As Worksheet ' moved as param now
    Dim numHandler As NumberHandler
    Set numHandler = New NumberHandler


    Dim st_h As StatusHandler, xStHelper As Integer
    Set st_h = New StatusHandler
    st_h.init_statusbar 20
    st_h.show
    
    
    delegacjaDlaProgresu st_h, xStHelper, 20
    
    
    ' inter4sh stands for internal suppliers list worksheet
    Dim ish As Worksheet
    ' already as params to have possibility for combo logic
    ' Dim osh As Worksheet, osh2 As Worksheet
    ' dim inter4Sh As Worksheet,
    Dim irng As Range, orng As Range
    Set ish = ThisWorkbook.Sheets.Add
    Set osh = ThisWorkbook.Sheets.Add
    
    
    ish.name = EVO.TryToRenameModule.tryToRenameWorksheet(ish, "IN1_" & CStr(r.offset(0, offst).Value) & "_")
    osh.name = EVO.TryToRenameModule.tryToRenameWorksheet(osh, "OUT1_" & CStr(r.offset(0, offst).Value) & "_")
    
    
    ' inter4Sh.name = EVO.TryToRenameModule.tryToRenameWorksheet(inter4Sh, "N_" & CStr(tbx3_Str) & "_")
    
    
    ' LABLES ------------------------------------------------
    
    fillLabels osh.Range("A1")
    
    ' fillLabels inter4Sh.Range("A1")
    
    ' -------------------------------------------------------
    
    Dim sap__handler As New SAP_Handler
    sap__handler.runMainLogicForSQ01__forMassItem r.offset(0, offst), ish, osh, st_h, xStHelper, 1
    
    Set sap4out = sap__handler
    
    
    
    
    ' COPY AND PASTE AS VALUES ------------------------------
    
    ' ???
    ' copyAndPasteAsValues osh.Range("A1").Offset(0, E_FROM_SQ01_QUASI_TP04_ARTICLE - 1)
    
    ' -------------------------------------------------------
    
    ' data ready - change string price into normal num
    changePricesIntoDouble osh
    
    
    st_h.hide
    Set st_h = Nothing
    Set numHandler = Nothing
    
    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
End Sub





Public Sub innerMassItemForSq01_WO_IN_LIST(r As Range, offst As Integer, ByRef osh As Worksheet, Optional ByRef sap4out As SAP_Handler)


    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    ' Dim osh As Worksheet ' moved as param now
    Dim numHandler As NumberHandler
    Set numHandler = New NumberHandler


    Dim st_h As StatusHandler, xStHelper As Integer
    Set st_h = New StatusHandler
    st_h.init_statusbar 20
    st_h.show
    
    
    delegacjaDlaProgresu st_h, xStHelper, 20
    
    
    ' inter4sh stands for internal suppliers list worksheet
    Dim ish As Worksheet
    ' already as params to have possibility for combo logic
    ' Dim osh As Worksheet, osh2 As Worksheet
    ' dim inter4Sh As Worksheet,
    Dim irng As Range, orng As Range
    Set ish = ThisWorkbook.Sheets.Add
    Set osh = ThisWorkbook.Sheets.Add
    
    
    ish.name = EVO.TryToRenameModule.tryToRenameWorksheet(ish, "IN1_" & CStr(r.offset(0, offst).Value) & "_")
    osh.name = EVO.TryToRenameModule.tryToRenameWorksheet(osh, "OUT1_" & CStr(r.offset(0, offst).Value) & "_")
    
    
    ' inter4Sh.name = EVO.TryToRenameModule.tryToRenameWorksheet(inter4Sh, "N_" & CStr(tbx3_Str) & "_")
    
    
    ' LABLES ------------------------------------------------
    
    fillLabels osh.Range("A1")
    
    ' fillLabels inter4Sh.Range("A1")
    
    ' -------------------------------------------------------
    
    Dim sap__handler As New SAP_Handler
    On Error Resume Next
    sap__handler.runMainLogicForSQ01__forMassItem_WO_IN_LIST r.offset(0, offst), osh, st_h, xStHelper, 1
    
    Set sap4out = sap__handler
    
    
    
    
    ' COPY AND PASTE AS VALUES ------------------------------
    
    ' ???
    ' copyAndPasteAsValues osh.Range("A1").Offset(0, E_FROM_SQ01_QUASI_TP04_ARTICLE - 1)
    
    ' -------------------------------------------------------
    
    ' data ready - change string price into normal num
    changePricesIntoDouble osh
    
    
    st_h.hide
    Set st_h = Nothing
    Set numHandler = Nothing
    
    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
End Sub

Public Sub innerMassItemForSq01_WO_IN_LIST_2(r As Range, offst As Integer, ByRef osh As Worksheet, Optional ByRef sap4out As SAP_Handler)


    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    ' Dim osh As Worksheet ' moved as param now
    Dim numHandler As NumberHandler
    Set numHandler = New NumberHandler


    Dim st_h As StatusHandler, xStHelper As Integer
    Set st_h = New StatusHandler
    st_h.init_statusbar 20
    st_h.show
    
    
    delegacjaDlaProgresu st_h, xStHelper, 20
    
    
    ' inter4sh stands for internal suppliers list worksheet
    Dim ish As Worksheet
    ' already as params to have possibility for combo logic
    ' Dim osh As Worksheet, osh2 As Worksheet
    ' dim inter4Sh As Worksheet,
    Dim irng As Range, orng As Range
    Set ish = ThisWorkbook.Sheets.Add
    Set osh = ThisWorkbook.Sheets.Add
    
    
    ish.name = EVO.TryToRenameModule.tryToRenameWorksheet(ish, "IN1_" & CStr(r.offset(0, offst).Value) & "_")
    osh.name = EVO.TryToRenameModule.tryToRenameWorksheet(osh, "OUT1_" & CStr(r.offset(0, offst).Value) & "_")
    
    
    ' inter4Sh.name = EVO.TryToRenameModule.tryToRenameWorksheet(inter4Sh, "N_" & CStr(tbx3_Str) & "_")
    
    
    ' LABLES ------------------------------------------------
    
    fillLabels osh.Range("A1")
    
    ' fillLabels inter4Sh.Range("A1")
    
    ' -------------------------------------------------------
    
    Dim sap__handler As New SAP_Handler
    sap__handler.runMainLogicForSQ01__forMassItem_WO_IN_LIST_2 r.offset(0, offst), osh, st_h, xStHelper, 1
    
    Set sap4out = sap__handler
    
    
    
    
    ' COPY AND PASTE AS VALUES ------------------------------
    
    ' ???
    ' copyAndPasteAsValues osh.Range("A1").Offset(0, E_FROM_SQ01_QUASI_TP04_ARTICLE - 1)
    
    ' -------------------------------------------------------
    
    ' data ready - change string price into normal num
    changePricesIntoDouble osh
    
    
    st_h.hide
    Set st_h = Nothing
    Set numHandler = Nothing
    
    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
End Sub

