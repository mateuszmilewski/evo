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


Private pc As PivotCache
Private pT As PivotTable


Private srcSheet As Worksheet
Private refRangeInSrcSheet As Range
Private pivotSheet As Worksheet
Private srcSheetName As String
Private pivotSheetName As String
Private tcamSheetName As String

' really special one!
Public proxy2 As Worksheet
Private rngFromProxy2 As Range

' tcam final
Public tcamSheet As Worksheet
Private tcamRng As Range






Public Sub runMainLogicForCreationTheoreticPivotTable()



    If Not rngFromProxy2 Is Nothing Then
    
    
        Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, rngFromProxy2)
        Set pT = pc.CreatePivotTable(pivotSheet.Cells(1, 1), "THEORETIC_PIVOT_" & CStr(pivotSheet.name))
        
        ' test loop through fields
        Dim pfs As PivotFields, pf As PivotField, pi As PivotItem
        Set pfs = pT.PivotFields
        
        For Each pf In pfs
            Debug.Print pf.name
            ' Debug.Print pf.Orientation
            
            If CStr(pf.name) = "S_ORDER_DATE" Then
                pf.Orientation = xlRowField
                pf.Position = 1
                ' pf.AutoGroup
                
                pf.name = "ORDER"
                
            End If
            
            If CStr(pf.name) = "S_DELIVERY_DATE" Then
                pf.Orientation = xlColumnField
                pf.Position = 1
                ' pf.AutoGroup
                
                
                pf.name = "DELIVERY"
                
                'For Each pi In pf.PivotItems
                '
                '    If IsDate(pi.Name) Then
                '        ' Debug.Print pi.Name
                '        pi.Name = Format(pi.Name, "dd/mm/yyyy")
                '    End If
                'Next pi
            End If
        Next pf
        
        
        For Each pf In pfs
            
            If CStr(pf.name) = "ROUTE NAME AND PILOT" Then
                pf.Orientation = xlRowField
                pf.Position = 2
            End If
            
            
            If CStr(pf.name) = "TYPE_DE_PIECE" Then
                pf.Orientation = xlPageField
                pf.Position = 1
            End If
            
            
            If CStr(pf.name) = "DELIVERY YEAR" Then
                pf.Orientation = xlPageField
                pf.Position = 1
                On Error Resume Next
                pf.CurrentPage = Year(Date)
            End If
            
            If CStr(pf.name) = "DELIVERY WEEK" Then
                pf.Orientation = xlPageField
                pf.Position = 1
                On Error Resume Next
                pf.CurrentPage = CStr(Application.WorksheetFunction.IsoWeekNum(CDate(Date)))
                
            End If
            
            
            
            If CStr(pf.name) = "(RN)(mL)" Then
                pf.Orientation = xlDataField
                pf.Position = 1
                'pf.Caption = "C(RN)(mL)"
                pf.Function = xlSum
                pf.name = "Th(RN)(mL)"
            End If
            
            'If Cstr(pf.Name) = "ROUNDUP (RN)(mL)" Then
            '    pf.Orientation = xlDataField
            '    pf.Position = 2
            '    'pf.Caption = "RC(RN)(mL)"
            '    pf.Function = xlSum
            '    pf.Name = "ThR(RN)(mL)"
            'End If
        Next pf
        
        ' ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
        ' pT.RowAxisLayout xlTabularRow
        'pT.RepeatAllLabels xlRepeatLabels
    Else
    
        MsgBox "no data provided for runMainLogicForCreationTheoritcPivotTable", vbCritical
        
    End If

End Sub

Public Sub runMainLogicForCreationPivotTable()

    If Not rngFromProxy2 Is Nothing Then
    
    
        Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, rngFromProxy2)
        Set pT = pc.CreatePivotTable(pivotSheet.Cells(1, 1), "PIVOT_" & CStr(pivotSheet.name))
        
        ' test loop through fields
        Dim pfs As PivotFields, pf As PivotField, pi As PivotItem
        Set pfs = pT.PivotFields
        For Each pf In pfs
            Debug.Print pf.name
            ' Debug.Print pf.Orientation
            
            If CStr(pf.name) = "S_ORDER_DATE" Then
                pf.Orientation = xlRowField
                pf.Position = 1
                ' pf.AutoGroup
                
                pf.name = "ORDER"
                
            End If
            
            If CStr(pf.name) = "S_DELIVERY_DATE" Then
                pf.Orientation = xlColumnField
                pf.Position = 1
                ' pf.AutoGroup
                
                
                pf.name = "DELIVERY"
                
                'For Each pi In pf.PivotItems
                '
                '    If IsDate(pi.Name) Then
                '        ' Debug.Print pi.Name
                '        pi.Name = Format(pi.Name, "dd/mm/yyyy")
                '    End If
                'Next pi
            End If
        Next pf
        
        
        For Each pf In pfs
            
            If CStr(pf.name) = "ROUTE NAME AND PILOT" Then
                pf.Orientation = xlRowField
                pf.Position = 2
            End If
            
            
            If CStr(pf.name) = "TYPE_DE_PIECE" Then
                pf.Orientation = xlPageField
                pf.Position = 1
            End If
            
            
            If CStr(pf.name) = "DELIVERY YEAR" Then
                pf.Orientation = xlPageField
                pf.Position = 1
                On Error Resume Next
                pf.CurrentPage = Year(Date)
            End If
            
            If CStr(pf.name) = "DELIVERY WEEK" Then
                pf.Orientation = xlPageField
                pf.Position = 1
                On Error Resume Next
                pf.CurrentPage = CStr(Application.WorksheetFunction.IsoWeekNum(CDate(Date)))
                
            End If
            
            
            
            If CStr(pf.name) = "Confirmed (RN)(mL)" Then
                pf.Orientation = xlDataField
                pf.Position = 1
                'pf.Caption = "C(RN)(mL)"
                pf.Function = xlSum
                pf.name = "C(RN)(mL)"
            End If
            
            'If Cstr(pf.Name) = "ROUNDUP Confirmed (RN)(mL)" Then
            '    pf.Orientation = xlDataField
            '    pf.Position = 2
            '    'pf.Caption = "RC(RN)(mL)"
            '    pf.Function = xlSum
            '    pf.Name = "RC(RN)(mL)"
            'End If
        Next pf
        
        ' ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
        ' pT.RowAxisLayout xlTabularRow
        'pT.RepeatAllLabels xlRepeatLabels
    Else
    
        MsgBox "no data provided for runMainLogicForCreationPivotTable", vbCritical
        
    End If
End Sub



Public Sub setRangeDataFromProxy2ForPivot(mr As Range)
    Set rngFromProxy2 = mr
End Sub

Public Function getRangeDataFromProxy2() As Range
    
    Set getRangeDataFromProxy2 = rngFromProxy2
End Function


Public Function getRefRangeInSrcSheet() As Range
    Set getRefRangeInSrcSheet = refRangeInSrcSheet
End Function

Public Sub moveRefRangeInSrcSheet()
    Set refRangeInSrcSheet = refRangeInSrcSheet.Offset(1, 0)
End Sub


Public Function getPivotSource() As Worksheet
    Set getPivotSource = srcSheet
End Function

Public Function getPivot() As Worksheet
    Set getPivot = pivotSheet
End Function

Public Sub setTcamPivot(mp As PivotTable)
    Set pT = Nothing
    On Error Resume Next
    Set pT = mp
End Sub

Public Sub setPivotSheet(psh1 As Worksheet)
    Set pivotSheet = Nothing
    On Error Resume Next
    Set pivotSheet = psh1
End Sub



Private Sub Class_Initialize()
    Set srcSheet = Nothing
    Set pivotSheet = Nothing
    Set rngFromProxy2 = Nothing
End Sub

Private Sub Class_Terminate()
    Set srcSheet = Nothing
    Set pivotSheet = Nothing
    Set rngFromProxy2 = Nothing
End Sub


Public Sub initPivotSource()
    Set srcSheet = ThisWorkbook.Sheets.Add
    srcSheet.name = CStr(tryToRenameSourceWorksheet(srcSheet))
    srcSheetName = CStr(srcSheet.name)
End Sub

Public Sub initPivotSheet()
    Set pivotSheet = ThisWorkbook.Sheets.Add
    pivotSheet.name = CStr(tryToRenamePivotWorksheet(pivotSheet))
    pivotSheetName = CStr(pivotSheet.name)
End Sub

Public Sub initTcamRepSheet()
    
    Set tcamSheet = ThisWorkbook.Sheets.Add
    tcamSheet.name = CStr(tryToRenameTcamRepWorksheet(tcamSheet))
    tcamSheetName = CStr(tcamSheet.name)
    
End Sub



Public Function tryToRenameProxy2Worksheet(psh As Worksheet) As String

    tryToRenameProxy2Worksheet = psh.name
    
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
    psh.name = CStr(tmpNewName)
    
    innerJokeFoo psh, tmpNewName
    
    
    tryToRenameProxy2Worksheet = psh.name

End Function



Private Function tryToRenameSourceWorksheet(psh As Worksheet) As String

    tryToRenameSourceWorksheet = psh.name
    
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
    psh.name = CStr(tmpNewName)
    
    innerJokeFoo psh, tmpNewName
    
    
    tryToRenameSourceWorksheet = psh.name

End Function

Private Function tryToRenamePivotWorksheet(psh As Worksheet) As String

    tryToRenamePivotWorksheet = psh.name
    
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
    psh.name = CStr(tmpNewName)
    
    innerJokeFoo psh, tmpNewName
    
    
    tryToRenamePivotWorksheet = psh.name

End Function


Private Function tryToRenameTcamRepWorksheet(psh As Worksheet) As String

    tryToRenameTcamRepWorksheet = psh.name
    
    Dim tmpNewName As String, mm As String, dd As String
    mm = CStr(Month(Date))
    dd = CStr(Day(Date))
    
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
    
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    
    tmpNewName = "TCAM_" & CStr(Year(Date)) & mm & dd & "_"
    
    
    On Error Resume Next
    psh.name = CStr(tmpNewName)
    
    innerJokeFoo psh, tmpNewName
    
    
    tryToRenameTcamRepWorksheet = psh.name

End Function


Private Sub innerJokeFoo(ByRef psh As Worksheet, newName As String)

    On Error Resume Next
    psh.name = CStr(newName)
    
    If psh.name = newName Then
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
    ' E_PIVOT_SRC__TN_ML
    ' E_PIVOT_SRC__CONFIRMED_TN_ML
    ' E_PIVOT_SRC__LQ
    ' E_PIVOT_SRC__RP
    ' E_PIVOT_SRC__TN_box_ML
    ' E_PIVOT_SRC__CONFIRMED_TN_box_ML
    ' E_PIVOT_SRC__RN_ML
    ' E_PIVOT_SRC__CONFIRMED_RN_ML
    ' short order date
    ' short del date


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
        
        
        With .Cells(1, E_PIVOT_SRC__LQ)
            .Value = "LQ"
            .AddComment CStr("LQ = roundup(OQ / BPL)")
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        With .Cells(1, E_PIVOT_SRC__CLQ)
            .Value = "Confirmed LQ"
            .AddComment CStr("Confirmed LQ = roundup(COQ / BPL)")
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        With .Cells(1, EVO.E_PIVOT_SRC__RP)
            .Value = "RP"
            .AddComment CStr("RP = LQ / LPP")
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        
        With .Cells(1, EVO.E_PIVOT_SRC__CRP)
            .Value = "Confirmed RP"
            .AddComment CStr("Confirmed RP = CLQ / LPP")
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        
        With .Cells(1, EVO.E_PIVOT_SRC__TN_box_ML)
            .Value = "(TNbox)(mL)"
            .AddComment CStr("(TNbox)(mL) = ((RP * 1,16) / (90)) * 13,5")
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        
        With .Cells(1, EVO.E_PIVOT_SRC__CONFIRMED_TN_box_ML)
            .Value = "Confirmed (TNbox)(mL)"
            .AddComment CStr("Confirmed (TNbox)(mL) = (( CRP * 1,16) / (90)) * 13,5")
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        
        With .Cells(1, EVO.E_PIVOT_SRC__RN_ML)
            .Value = "(RN)(mL)"
            .AddComment "(RN)(mL) = TN + TN_box"
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        
        With .Cells(1, EVO.E_PIVOT_SRC__CONFIRMED_RN_ML)
            .Value = "Confirmed (RN)(mL)"
            .AddComment "Confirmed (RN)(mL) = TN + TN_box"
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        
        With .Cells(1, EVO.E_PIVOT_SRC__ROUNDUP_RN_ML)
            .Value = "ROUNDUP (RN)(mL)"
            .AddComment "ROUNDUP (RN)(mL) = TN + TN_box"
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        
        With .Cells(1, EVO.E_PIVOT_SRC__ROUNDUP_CONFIRMED_RN_ML)
            .Value = "ROUNDUP Confirmed (RN)(mL)"
            .AddComment "ROUNDUP Confirmed (RN)(mL) = TN + TN_box"
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        .Cells(1, EVO.E_PIVOT_SRC___SHORT_ORDER_DATE).Value = "S_ORDER_DATE"
        .Cells(1, EVO.E_PIVOT_SRC___SHORT_DEL_DATE).Value = "S_DELIVERY_DATE"
        
        .Cells(1, EVO.E_PIVOT_SRC____TYPE_DE_PIECE).Value = "TYPE_DE_PIECE"
        .Cells(1, EVO.E_PIVOT_SRC____CLOE_COL_E_PRICE).Value = "PRICE_FROM_CLOE_COL_E"
        
        
        .Cells(1, EVO.E_PIVOT_SRC__lb_COFOR_HYBRID).Value = "COFOR HYBRID"
        .Cells(1, EVO.E_PIVOT_SRC__lb_DHAS).Value = "DHAS"
        .Cells(1, EVO.E_PIVOT_SRC__lb_DHEO).Value = "DHEO"
        .Cells(1, EVO.E_PIVOT_SRC__lb_DHEF).Value = "DHEF"
        
        .Cells(1, EVO.E_PIVOT_SRC__lb_INCOTERM).Value = "INCOTERM"
        .Cells(1, EVO.E_PIVOT_SRC__lb_PILOT).Value = "PILOT"
        .Cells(1, EVO.E_PIVOT_SRC__lb_ROUTE).Value = "ROUTE"
        
        
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
    Sheets("Sheet1").name = "PIVOT_001"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "[PICK_UP_SHEET_coforTest.xlsm]BASE!R2C1:R370C48", Version:=6). _
        CreatePivotTable TableDestination:="PIVOT_001!R1C1", TableName:= _
        "PivotTable1", DefaultVersion:=6
    Sheets("PIVOT_001").Select
    Cells(1, 1).Select
    Windows("PICK_UP_SHEET_coforTest.xlsm").Activate
    Windows("EVO_013.xlsm").Activate
End Sub

