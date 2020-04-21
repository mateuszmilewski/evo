Attribute VB_Name = "TestModule"
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

' testing overall status
Private Sub testMain()


    Dim sh As StatusHandler, x As Integer
    Set sh = New StatusHandler
    sh.init_statusbar 10
    sh.show
    For x = 1 To 10
        Sleep 1000
        sh.progress_increase
    Next x
    
    sh.hide
    
    Set sh = Nothing
    
End Sub



' PIVOT tests
Private Sub testOnPivot()
    
    Dim p As PivotHandler
    Set p = New PivotHandler
    p.initPivotSource
    p.initPivotSheet
    Set p = Nothing

End Sub


Private Sub testForArr()
    
    Dim txt As String
    txt = " fefh iefj e"
    
    Dim arr As Variant
    
    arr = Split(txt, " ")
    Dim x As Integer
    For x = LBound(arr) To UBound(arr)
        Debug.Print arr(x)
    Next x
End Sub




Private Sub kopiaSourcePivot()
'
' kopiaSourcePivot Macro
'

'
    Sheets("SrcPivot_20200408_II").Select
    Sheets("SrcPivot_20200408_II").Copy Before:=Sheets(1)
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AM$4214").RemoveDuplicates columns:=1, Header:= _
        xlYes

    columns("U:X").Select
    Selection.Delete Shift:=xlToLeft
    columns("U:U").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    columns("R:S").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
End Sub











Private Sub newPivotFromProxy2()
'
' newPivotFromProxy2 Macro
'

'
    Range("A1:AO2364").Select
    Range("G6").Activate
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Proxy2_20200409_!R1C1:R2364C41", Version:=6).CreatePivotTable _
        TableDestination:="Sheet5!R3C1", TableName:="PivotTable2", DefaultVersion _
        :=6
    Sheets("Sheet5").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY YEAR")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY WEEK")
        .Orientation = xlPageField
        .Position = 1
    End With
    Range("B1").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY YEAR"). _
        CurrentPage = "2020"
    Range("B2").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY WEEK"). _
        CurrentPage = "16"
    Range("B2").Select
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY DATE")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY DATE").AutoGroup
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Months").Orientation = _
        xlHidden
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("ORDER DATE")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ORDER DATE").AutoGroup
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Months2").Orientation = _
        xlHidden
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("ROUTE NAME AND PILOT")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Confirmed (RN)(mL)"), "Sum of Confirmed (RN)(mL)", _
        xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("ROUNDUP Confirmed (RN)(mL)"), _
        "Sum of ROUNDUP Confirmed (RN)(mL)", xlSum
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ID").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("WIERSZ").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("REF").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("COFOR_VENDEUR").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("COFOR_EXPEDITEUR"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("NOM_FOURNISSEUR"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY DATE").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY YEAR").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY MONTH").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY WEEK").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY YYYYCW"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("AK").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ORDER DATE").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ORDER YEAR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ORDER MONTH").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ORDER WEEK").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ORDER YYYYCW").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ROUTE NAME AND PILOT"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("QTY 2").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("CQTY 2").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("UC 2").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("OQ 2").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Confirmed OQ 2").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("CONDI").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("PC_GV").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("BPC").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MC").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MBU").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MAX CAPACITY").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("(TN)(mL)").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Confirmed (TN)(mL)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("LQ").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Confirmed LQ").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("RP").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Confirmed RP").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("(TNbox)(mL)").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Confirmed (TNbox)(mL)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("(RN)(mL)").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Confirmed (RN)(mL)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ROUNDUP (RN)(mL)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ROUNDUP Confirmed (RN)(mL)" _
        ).Subtotals = Array(False, False, False, False, False, False, False, False, False, False _
        , False, False)
    ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
    Range("C6").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of Confirmed (RN)(mL)").Caption = "C (RN)(mL)"
    Range("D6").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of ROUNDUP Confirmed (RN)(mL)").Caption = "RC (RN)(mL)"
    Range("C4").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY DATE").Caption = _
        "DD"
    Range("D4").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.Caption = "V"
    Range("B2").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY WEEK"). _
        CurrentPage = "17"
    Range("B2").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DELIVERY WEEK"). _
        CurrentPage = "16"
    Range("B3").Select
    ActiveWindow.SmallScroll ToRight:=0
    Range("B6").Select
    With ActiveSheet.PivotTables("PivotTable2")
        .ColumnGrand = False
        .RowGrand = False
    End With
    ActiveWindow.SmallScroll Down:=-9
    Range("A3").Select
    columns("A:A").ColumnWidth = 16.29
    columns("B:B").EntireColumn.AutoFit
End Sub





Private Sub testRemoveTotals()
'
' testRemoveTotals Macro
'

'
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("ID").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("WIERSZ"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("REF").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("COFOR_VENDEUR"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields( _
        "COFOR_EXPEDITEUR").Subtotals = Array(False, False, False, False, False, False, False _
        , False, False, False, False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("NOM_FOURNISSEUR" _
        ).Subtotals = Array(False, False, False, False, False, False, False, False, False, False _
        , False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("DELIVERY DATE"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("DELIVERY YEAR"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("DELIVERY MONTH") _
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("DELIVERY WEEK"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("DELIVERY YYYYCW" _
        ).Subtotals = Array(False, False, False, False, False, False, False, False, False, False _
        , False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("AK").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("ORDER DATE"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("ORDER YEAR"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("ORDER MONTH"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("ORDER WEEK"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("ORDER YYYYCW"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields( _
        "ROUTE NAME AND PILOT").Subtotals = Array(False, False, False, False, False, False, _
        False, False, False, False, False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("QTY 2"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("CQTY 2"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("UC 2"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("OQ 2"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("Confirmed OQ 2") _
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("CONDI"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("PC_GV"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("BPC").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("MC").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("MBU").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("MAX CAPACITY"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("(TN)(mL)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields( _
        "Confirmed (TN)(mL)").Subtotals = Array(False, False, False, False, False, False, _
        False, False, False, False, False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("LQ").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("Confirmed LQ"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("RP").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("Confirmed RP"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("(TNbox)(mL)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields( _
        "Confirmed (TNbox)(mL)").Subtotals = Array(False, False, False, False, False, False, _
        False, False, False, False, False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("(RN)(mL)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields( _
        "Confirmed (RN)(mL)").Subtotals = Array(False, False, False, False, False, False, _
        False, False, False, False, False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields( _
        "ROUNDUP (RN)(mL)").Subtotals = Array(False, False, False, False, False, False, False _
        , False, False, False, False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields( _
        "ROUNDUP Confirmed (RN)(mL)").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("S_ORDER_DATE"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PIVOT_Pivot_20200414_I").PivotFields("S_DELIVERY_DATE" _
        ).Subtotals = Array(False, False, False, False, False, False, False, False, False, False _
        , False, False)
End Sub


