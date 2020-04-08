VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotHandler"
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

' PIVOT HANDLER


Private srcSheet As Worksheet
Private refRangeInSrcSheet As Range
Private pivotSheet As Worksheet
Private srcSheetName As String
Private pivotSheetName As String

' really special one!
Public proxy2 As Worksheet




Public Function getRefRangeInSrcSheet() As Range
    Set getRefRangeInSrcSheet = refRangeInSrcSheet
End Function

Public Sub moveRefRangeInSrcSheet()
    Set refRangeInSrcSheet = refRangeInSrcSheet.Offset(1, 0)
End Sub


Public Function getPivotSource() As Worksheet
    Set getPivotSource = srcSheet
End Function


Private Sub Class_Initialize()
    Set srcSheet = Nothing
    Set pivotSheet = Nothing
End Sub

Private Sub Class_Terminate()
    Set srcSheet = Nothing
    Set pivotSheet = Nothing
End Sub


Public Sub initPivotSource()
    Set srcSheet = ThisWorkbook.Sheets.Add
    srcSheet.Name = CStr(tryToRenameSourceWorksheet(srcSheet))
    srcSheetName = CStr(srcSheet.Name)
End Sub

Public Sub initPivotSheet()
    Set pivotSheet = ThisWorkbook.Sheets.Add
    pivotSheet.Name = CStr(tryToRenamePivotWorksheet(pivotSheet))
    pivotSheetName = CStr(pivotSheet.Name)
End Sub



Public Function tryToRenameProxy2Worksheet(psh As Worksheet) As String

    tryToRenameProxy2Worksheet = psh.Name
    
    Dim tmpNewName As String, mm As String, dd As String
    mm = CStr(Month(Date))
    dd = CStr(Day(Date))
    
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
    
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    
    ' On Error Resume Next
    tmpNewName = "Proxy2_" & CStr(Year(Date)) & mm & dd & "_"
    
    
    On Error Resume Next
    psh.Name = CStr(tmpNewName)
    
    innerJokeFoo psh, tmpNewName
    
    
    tryToRenameProxy2Worksheet = psh.Name

End Function



Private Function tryToRenameSourceWorksheet(psh As Worksheet) As String

    tryToRenameSourceWorksheet = psh.Name
    
    Dim tmpNewName As String, mm As String, dd As String
    mm = CStr(Month(Date))
    dd = CStr(Day(Date))
    
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
    
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    
    ' On Error Resume Next
    tmpNewName = "SrcPivot_" & CStr(Year(Date)) & mm & dd & "_"
    
    
    On Error Resume Next
    psh.Name = CStr(tmpNewName)
    
    innerJokeFoo psh, tmpNewName
    
    
    tryToRenameSourceWorksheet = psh.Name

End Function

Private Function tryToRenamePivotWorksheet(psh As Worksheet) As String

    tryToRenamePivotWorksheet = psh.Name
    
    Dim tmpNewName As String, mm As String, dd As String
    mm = CStr(Month(Date))
    dd = CStr(Day(Date))
    
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
    
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    
    tmpNewName = "Pivot_" & CStr(Year(Date)) & mm & dd & "_"
    
    
    On Error Resume Next
    psh.Name = CStr(tmpNewName)
    
    innerJokeFoo psh, tmpNewName
    
    
    tryToRenamePivotWorksheet = psh.Name

End Function


Private Sub innerJokeFoo(ByRef psh As Worksheet, newName As String)

    On Error Resume Next
    psh.Name = CStr(newName)
    
    If psh.Name = newName Then
        Exit Sub
    Else
    
        If Len(newName) < 30 Then
            innerJokeFoo psh, newName & "I"
        Else
            Exit Sub
        End If
    End If
End Sub

Public Sub fillPivotSourceWithLabels()

    '    ID = 1
    '    E_PIVOT_SRC_WIERSZ = 2
    '    E_PIVOT_SRC_REF
    '    E_PIVOT_SRC_COFOR_VENDEUR
    '    E_PIVOT_SRC_COFOR_EXPEDITEUR
    '    E_PIVOT_SRC_SUPPLIER_NAME
    '    E_PIVOT_SRC_DELIVERY_DATE
    '    E_PIVOT_SRC_DELIVERY_YEAR
    '    E_PIVOT_SRC_DELIVERY_MONTH
    '    E_PIVOT_SRC_DELIVERY_WEEK
    '    yyyycw
    '    ak col from ple
    '    E_PIVOT_SRC_ORDER_DATE
    '    E_PIVOT_SRC_ORDER_YEAR
    '    E_PIVOT_SRC_ORDER_MONTH
    '    E_PIVOT_SRC_ORDER_WEEK
    '    yyyycw
    '    E_PIVOT_SRC_ROUTE
    '    E_PIVOT_SRC_PILOT
    '    E_PIVOT_SRC_ROUTE_AND_PILOT
    '
    '    qty
    '    conf qty
    '    uc
    '    OQ
    '    Confirmed OQ
    
    'E_PIVOT_SRC_SUMIF_QTY
    'E_PIVOT_SRC_SUMIF_CQTY
    'E_PIVOT_SRC_SUMIF_UC
    'E_PIVOT_SRC_SUMIF_OQ
    'E_PIVOT_SRC_SUMIF_COQ
    '    condi
    ' E_PIVOT_SRC_UA_PC_GV
    ' E_PIVOT_SRC_UA_MAX_CAPACITY
    ' E_PIVOT_SRC_TN_ML


    With srcSheet
        .Cells(1, EVO.E_PIVOT_SRC_ID).Value = "ID"
        .Cells(1, EVO.E_PIVOT_SRC_WIERSZ).Value = "WIERSZ"
        .Cells(1, EVO.E_PIVOT_SRC_REF).Value = "REF"
        .Cells(1, EVO.E_PIVOT_SRC_COFOR_VENDEUR).Value = "COFOR_VENDEUR"
        .Cells(1, EVO.E_PIVOT_SRC_COFOR_EXPEDITEUR).Value = "COFOR_EXPEDITEUR"
        .Cells(1, EVO.E_PIVOT_SRC_SUPPLIER_NAME).Value = "NOM_FOURNISSEUR"
        
        .Cells(1, E_PIVOT_SRC_DELIVERY_DATE).Value = "DELIVERY DATE"
        .Cells(1, E_PIVOT_SRC_DELIVERY_YEAR).Value = "DELIVERY YEAR"
        .Cells(1, E_PIVOT_SRC_DELIVERY_MONTH).Value = "DELIVERY MONTH"
        .Cells(1, E_PIVOT_SRC_DELIVERY_WEEK).Value = "DELIVERY WEEK"
        .Cells(1, EVO.E_PIVOT_SRC_DELIVERY_YYYYCW).Value = "DELIVERY YYYYCW"
        
        .Cells(1, EVO.E_PIVOT_SRC_AK_COL_FROM_PLE).Value = "AK"
        
        .Cells(1, E_PIVOT_SRC_ORDER_DATE).Value = "ORDER DATE"
        .Cells(1, E_PIVOT_SRC_ORDER_YEAR).Value = "ORDER YEAR"
        .Cells(1, E_PIVOT_SRC_ORDER_MONTH).Value = "ORDER MONTH"
        .Cells(1, E_PIVOT_SRC_ORDER_WEEK).Value = "ORDER WEEK"
        .Cells(1, EVO.E_PIVOT_SRC_ORDER_YYYYCW).Value = "ORDER YYYYCW"
        
        .Cells(1, E_PIVOT_SRC_ROUTE).Value = "ROUTE NAME"
        .Cells(1, E_PIVOT_SRC_PILOT).Value = "PILOT"
        .Cells(1, E_PIVOT_SRC_ROUTE_AND_PILOT).Value = "ROUTE NAME AND PILOT"
        
        .Cells(1, EVO.E_PIVOT_SRC_QTY).Value = "QTY"
        .Cells(1, EVO.E_PIVOT_SRC_CQTY).Value = "CQTY"
        .Cells(1, EVO.E_PIVOT_SRC_UC).Value = "UC"
        .Cells(1, EVO.E_PIVOT_SRC_OQ).Value = "OQ"
        .Cells(1, EVO.E_PIVOT_SRC_COQ).Value = "Confirmed OQ"
        
        
        .Cells(1, EVO.E_PIVOT_SRC_ROUNDUP_OQ1).Value = "ROUNDUP OQ 1"
        .Cells(1, EVO.E_PIVOT_SRC_ROUNDUP_COQ1).Value = "ROUNDUP Confirmed OQ 1"
        
        
        .Cells(1, E_PIVOT_SRC_SUMIF_QTY).Value = "QTY 2"
        .Cells(1, E_PIVOT_SRC_SUMIF_CQTY).Value = "CQTY 2"
        .Cells(1, E_PIVOT_SRC_SUMIF_UC).Value = "UC 2"
        .Cells(1, E_PIVOT_SRC_SUMIF_OQ).Value = "OQ 2"
        .Cells(1, E_PIVOT_SRC_SUMIF_COQ).Value = "Confirmed OQ 2"
        
        .Cells(1, EVO.E_PIVOT_SRC_CONDI).Value = "CONDI"
        .Cells(1, E_PIVOT_SRC_UA_PC_GV).Value = "PC_GV"
        .Cells(1, E_PIVOT_SRC_UA_BPC).Value = "BPC"
        .Cells(1, E_PIVOT_SRC_UA_MC).Value = "MC"
        .Cells(1, E_PIVOT_SRC_UA_MBU).Value = "MBU"
        .Cells(1, E_PIVOT_SRC_UA_MAX_CAPACITY).Value = "MAX CAPACITY"
        
        With .Cells(1, E_PIVOT_SRC__TN_ML)
            .Value = "(TN)(mL)"
            .AddComment CStr("(TN)(mL) = sum(OQ/PA) * 13,5")
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        With .Cells(1, E_PIVOT_SRC__CONFIRMED_TN_ML)
            .Value = "Confirmed (TN)(mL)"
            .AddComment CStr("Confirmed (TN)(mL) = sum(OQ/PA) * 13,5")
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
    End With

End Sub

Public Sub mountStartingPoint()


    Set refRangeInSrcSheet = srcSheet.Cells(2, 1)
    
    
End Sub



' static code - do not use - just for refractorisation!
Private Sub TestTemplateForPivotMacro()
'
' TestTemplateForPivotMacro Macro
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "PIVOT_001"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "[PICK_UP_SHEET_coforTest.xlsm]BASE!R2C1:R370C48", Version:=6). _
        CreatePivotTable TableDestination:="PIVOT_001!R1C1", TableName:= _
        "PivotTable1", DefaultVersion:=6
    Sheets("PIVOT_001").Select
    Cells(1, 1).Select
    Windows("PICK_UP_SHEET_coforTest.xlsm").Activate
    Windows("EVO_013.xlsm").Activate
End Sub

