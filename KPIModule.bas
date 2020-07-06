Attribute VB_Name = "KPIModule"
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
' KPI MODULE


Public Sub createKpi(ictrl As IRibbonControl)
    innerCreateKPI
End Sub


Public Sub innerCreateKPI()

    
    Dim w As Workbook
    
    With FileChooser
        .scenarioType = E_FORM_SCENARIO_KPI
        .BtnCopy.Enabled = False
        .BtnValid.Enabled = True
        
        .ComboBoxFeed.Clear
        .ComboBoxMaster.Clear
        
        
        For Each w In Application.Workbooks
            .ComboBoxFeed.AddItem w.name
            .ComboBoxMaster.AddItem w.name
        Next w
        
        .ComboBoxFeed.Value = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("m2").Value
        .ComboBoxMaster.Value = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("m1").Value
        
        .show
    End With
    
    MsgBox "GOTOWE!", vbInformation
    
End Sub


Public Sub innerAfterFormCreateKPI(masterFileName, feedFileName, Optional sh As StatusHandler)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    
    If IsMissing(sh) Or (sh Is Nothing) Then
        Set sh = New StatusHandler
    End If
    
    ' master and feed worksheets
    Dim m As Worksheet, f As Worksheet
    ' starting from most impotant sheets!
    On Error Resume Next
    Set m = Workbooks(masterFileName).Sheets(MAIN_SH_BASE)
    On Error Resume Next
    Set f = Workbooks(feedFileName).Sheets(G_FEED_SH_MAIN)
    
    
    
    Dim kpi_h As KpiHandler
    Set kpi_h = New KpiHandler
    kpi_h.set_master m
    kpi_h.setStatusHandler sh
    kpi_h.fillPartNumberDictionary
    
    kpi_h.makeRepFromPns
    
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub




' KPI small buttons!
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------

Public Sub kpiFilterOnUA(ictrl As IRibbonControl)
    kpiFilterOn ""
    kpiFilterOn "ua"
End Sub

Public Sub kpiFilterOnPLE(ictrl As IRibbonControl)
    kpiFilterOn ""
    kpiFilterOn "ple"
End Sub

Public Sub clearKpiFilter(ictrl As IRibbonControl)
    kpiFilterOn ""
End Sub


' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------


' side subs and functions for kpi small buttons
' ---------------------------------------------------------------------

Public Sub kpiFilterOn(str As String)



'    E_KPI_SRC_REF = 1
'    E_KPI_SRC_V_COFOR
'    E_KPI_SRC_E_COFOR
'    E_KPI_SRC_SUPPLIER_NAME
'    E_KPI_SRC_DAP
'    E_KPI_SRC_NO_DATA
'    E_KPI_SRC_TYPE_DE_PIECE
'    E_KPI_SRC_UA
'    E_KPI_SRC_PLE
'    E_KPI_SRC_GREEN
'    E_KPI_SRC_BLUE
'    E_KPI_SRC_YELLOW


    ' check if activesheet is in fact kpi src
    If ActiveSheet.name Like "*KPI_SRC_*" Then
    
        Dim tb As Range
        Set tb = ActiveSheet.UsedRange
        
        ' Debug.Print tb.Address
        
        innerKpiFilter str, tb
        
    
    Else
        MsgBox "to make filtering stuff you need to have kpi raw table activated and in front of you!"
    End If
    
    

    

End Sub


Private Sub innerKpiFilter(str As String, ByRef tb As Range)

    If str = "UA" Or str = "ua" Then
        With tb
            .AutoFilter Field:=E_KPI_SRC_DAP, Criteria1:="0"
            .AutoFilter Field:=E_KPI_SRC_NO_DATA, Criteria1:="0"
            .AutoFilter Field:=E_KPI_SRC_UA, Criteria1:="0"
        End With
    ElseIf str = "PLE" Or str = "ple" Then
        With tb
            .AutoFilter Field:=E_KPI_SRC_DAP, Criteria1:="0"
            .AutoFilter Field:=E_KPI_SRC_NO_DATA, Criteria1:="0"
            .AutoFilter Field:=E_KPI_SRC_PLE, Criteria1:="0"
        End With
    Else
        ' clear filter
        With tb
            On Error Resume Next
            .Parent.ShowAllData
        End With
    End If
End Sub


' ---------------------------------------------------------------------







' logic for lean tables in kpi logic
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------


Public Sub initLeanTablesAndCharts(ictrl As IRibbonControl)
    makeSomeLeanTablesAndChartsForActiveSheet
End Sub


Private Sub makeSomeLeanTablesAndChartsForActiveSheet()



'    E_KPI_SRC_REF = 1
'    E_KPI_SRC_V_COFOR
'    E_KPI_SRC_E_COFOR
'    E_KPI_SRC_SUPPLIER_NAME
'    E_KPI_SRC_DAP
'    E_KPI_SRC_NO_DATA
'    E_KPI_SRC_TYPE_DE_PIECE
'    E_KPI_SRC_UA
'    E_KPI_SRC_PLE
'    E_KPI_SRC_GREEN
'    E_KPI_SRC_BLUE
'    E_KPI_SRC_YELLOW


    ' check if activesheet is in fact kpi src
    If ActiveSheet.name Like "*KPI_SRC_*" Then
    
        Dim tb As Range, dapRng As Range, noDataRng As Range, srcSh As Worksheet
        
        Set srcSh = ActiveSheet
        Set tb = ActiveSheet.UsedRange
        Dim countRows As Long
        countRows = tb.rows.Count - 1
        
        With ActiveSheet
            Set dapRng = .Range(.Cells(2, E_KPI_SRC_DAP), .Cells(countRows, E_KPI_SRC_DAP))
            Set noDataRng = .Range(.Cells(2, E_KPI_SRC_NO_DATA), .Cells(countRows, E_KPI_SRC_NO_DATA))
        End With
        

        
        Dim newRange As Range, sh As Worksheet
        Set sh = ThisWorkbook.Sheets.Add
        sh.name = CStr(tryToRenameWorksheet3(sh))
        Set newRange = sh.Cells(1, 1)
        
        
        ' all data
        sh.Cells(2, 2).Value = "all data lines: "
        sh.Cells(2, 3).Value = countRows
        
        ' ile dapow
        sh.Cells(4, 2).Value = "DAPs"
        sh.Cells(4, 3).Value = "DAP"
        sh.Cells(4, 4).Value = "NO DAP"
        
        
        
        ' -------------------------------------------------------------------------------------------------
        sh.Cells(5, 2).Value = "montage"
        sh.Cells(6, 2).Value = "ferrage"
        
        ' dap for montage
        sh.Cells(5, 3).Value = CLng(innerCalcKpi(srcSh, "montage", E_KPI_SRC_TYPE_DE_PIECE, "1", E_KPI_SRC_DAP))
        sh.Cells(6, 3).Value = CLng(innerCalcKpi(srcSh, "ferrage", E_KPI_SRC_TYPE_DE_PIECE, "1", E_KPI_SRC_DAP))
        ' NO DAP
        sh.Cells(5, 4).Value = CLng(innerCalcKpi(srcSh, "montage", E_KPI_SRC_TYPE_DE_PIECE, "0", E_KPI_SRC_DAP))
        sh.Cells(6, 4).Value = CLng(innerCalcKpi(srcSh, "ferrage", E_KPI_SRC_TYPE_DE_PIECE, "0", E_KPI_SRC_DAP))
        
        ' -------------------------------------------------------------------------------------------------
        
        ' -------------------------------------------------------------------------------------------------
        'PLE
        sh.Cells(9, 2).Value = "PLE"
        sh.Cells(9, 3).Value = "PLE OK"
        sh.Cells(9, 4).Value = "PLE NOK"
        
        sh.Cells(10, 2).Value = "montage"
        sh.Cells(11, 2).Value = "ferrage"
        
        ' ple for montage
        sh.Cells(10, 3).Value = _
            CLng(innerCalcKpi2ForNoDapOnly(srcSh, "montage", E_KPI_SRC_TYPE_DE_PIECE, "1", E_KPI_SRC_PLE))
        sh.Cells(11, 3).Value = _
            CLng(innerCalcKpi2ForNoDapOnly(srcSh, "ferrage", E_KPI_SRC_TYPE_DE_PIECE, "1", E_KPI_SRC_PLE))
        ' no ple
        sh.Cells(10, 4).Value = _
            CLng(innerCalcKpi2ForNoDapOnly(srcSh, "montage", E_KPI_SRC_TYPE_DE_PIECE, "0", E_KPI_SRC_PLE))
        sh.Cells(11, 4).Value = _
            CLng(innerCalcKpi2ForNoDapOnly(srcSh, "ferrage", E_KPI_SRC_TYPE_DE_PIECE, "0", E_KPI_SRC_PLE))
        ' -------------------------------------------------------------------------------------------------
        
        ' -------------------------------------------------------------------------------------------------
        'UA
        sh.Cells(13, 2).Value = "UA"
        sh.Cells(13, 3).Value = "UA OK"
        sh.Cells(13, 4).Value = "UA NOK"
        
        sh.Cells(14, 2).Value = "montage"
        sh.Cells(15, 2).Value = "ferrage"
        
        ' ple for montage
        sh.Cells(14, 3).Value = _
            CLng(innerCalcKpi2ForNoDapOnly(srcSh, "montage", E_KPI_SRC_TYPE_DE_PIECE, "1", E_KPI_SRC_UA))
        sh.Cells(15, 3).Value = _
            CLng(innerCalcKpi2ForNoDapOnly(srcSh, "ferrage", E_KPI_SRC_TYPE_DE_PIECE, "1", E_KPI_SRC_UA))
        ' no ple
        sh.Cells(14, 4).Value = _
            CLng(innerCalcKpi2ForNoDapOnly(srcSh, "montage", E_KPI_SRC_TYPE_DE_PIECE, "0", E_KPI_SRC_UA))
        sh.Cells(15, 4).Value = _
            CLng(innerCalcKpi2ForNoDapOnly(srcSh, "ferrage", E_KPI_SRC_TYPE_DE_PIECE, "0", E_KPI_SRC_UA))
        ' -------------------------------------------------------------------------------------------------
        
        
        
        
        ' from diff module - chart module
        addCharts sh, sh.Range("b4")
        
        
    
    Else
        MsgBox "to make filtering stuff you need to have kpi raw table activated and in front of you!"
    End If
    
    

    

End Sub

Private Function innerCalcKpi(srcSh As Worksheet, strTypeDePiece As String, e1 As E_KPI_SRC, strDap As String, e2 As E_KPI_SRC) As Long
    innerCalcKpi = 0
    
    Dim wiersz As Long
    wiersz = 2
    
    Dim rng As Range
    Set rng = srcSh.Cells(wiersz, 1)
    
    Do
    
        If srcSh.Cells(wiersz, e1).Value Like "*" & strTypeDePiece & "*" Then
            
            
            If CStr(srcSh.Cells(wiersz, e2).Value) = CStr(strDap) Then
                innerCalcKpi = innerCalcKpi + 1
            End If
        End If
    
        Set rng = rng.Offset(1, 0)
        wiersz = wiersz + 1
    Loop Until CStr(rng) = ""
End Function

Private Function innerCalcKpi2ForNoDapOnly(srcSh As Worksheet, strTypeDePiece As String, e1 As E_KPI_SRC, strVal As String, e2 As E_KPI_SRC) As Long
    
    
    innerCalcKpi2ForNoDapOnly = 0
    
    Dim wiersz As Long, eDap As E_KPI_SRC, eNoData As E_KPI_SRC
    wiersz = 2
    eDap = E_KPI_SRC_DAP
    eNoData = E_KPI_SRC_NO_DATA
    
    
    Dim rng As Range
    Set rng = srcSh.Cells(wiersz, 1)
    
    Do
    
        If srcSh.Cells(wiersz, e1).Value Like "*" & strTypeDePiece & "*" Then
            
            ' no dap! mandatory!
            If CStr(srcSh.Cells(wiersz, eDap).Value) = "0" Then
                If CStr(srcSh.Cells(wiersz, eNoData).Value) = "0" Then
                    If CStr(srcSh.Cells(wiersz, e2).Value) = CStr(strVal) Then
                        innerCalcKpi2ForNoDapOnly = innerCalcKpi2ForNoDapOnly + 1
                    End If
                End If
            End If
        End If
    
        Set rng = rng.Offset(1, 0)
        wiersz = wiersz + 1
    Loop Until CStr(rng) = ""
End Function


' try rename to juz 3 taka sama funkcja :/
Private Function tryToRenameWorksheet3(s1 As Worksheet) As String

    tryToRenameWorksheet3 = s1.name
    
    Dim tmpNewName As String, mm As String, dd As String
    mm = CStr(Month(Date))
    dd = CStr(Day(Date))
    
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
    
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    
    tmpNewName = "KPI_" & CStr(Year(Date)) & mm & dd & "_"
    
    
    On Error Resume Next
    s1.name = CStr(tmpNewName)
    
    noFukinDryBecauseFukOff3 s1, tmpNewName
    tryToRenameWorksheet3 = s1.name

End Function

Private Sub noFukinDryBecauseFukOff3(ByRef psh As Worksheet, newName As String)

    On Error Resume Next
    psh.name = CStr(newName)
    
    If psh.name = newName Then
        Exit Sub
    Else
    
        If Len(newName) < 30 Then
            noFukinDryBecauseFukOff3 psh, newName & "I"
        Else
            Exit Sub
        End If
    End If
End Sub

' ---------------------------------------------------------------------
' ---------------------------------------------------------------------

