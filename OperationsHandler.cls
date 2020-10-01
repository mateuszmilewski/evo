VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OperationsHandler"
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


Public Function simpleByCellValidation(sh1 As Worksheet) As E_PIVOT_STD
    simpleByCellValidation = E_PIVOT_ANGEL_STD
    
    Dim tmpB As Boolean
    tmpB = True
    If sh1.Cells(1, 1).Value = "DELIVERY YEAR" Then
        If sh1.Cells(2, 1).Value = "DELIVERY WEEK" Then
            If sh1.Cells(3, 1).Value = "TYPE_DE_PIECE" Then
            
                If sh1.Cells(5, 1).Value = "OQ ( OQ 2 )" Then
                
                    simpleByCellValidation = E_PIVOT_CAROLINE_STD
                    
                ElseIf sh1.Cells(5, 1).Value = "Th(RN)(mL)" Then
                
                    simpleByCellValidation = E_PIVOT_ANGEL_STD
                    
                Else
                    tmpB = True
                End If
            Else
                tmpB = False
            End If
        Else
            tmpB = False
        End If
    Else
        tmpB = False
    End If
    
    If Not tmpB Then
        MsgBox "This active sheet is not pivot in proper stanard!"
        End
    End If
End Function


Public Sub recalculDatesIfSatOrSun(ByRef li As LineItem, Optional e As E_ECHANCIER_ONL_semaine_SCENARIO)


    Dim wd As Variant
    
    
    
    If e = E_ECHANCIER_ONL_semaine_SCENARIO_DEL Then
    
        wd = Application.WorksheetFunction.Weekday(li.puDate, 2)
        
        If Int(wd) = 7 Then
            li.puDate = CDate(li.puDate - 2)
        End If
        
        If Int(wd) = 6 Then
            li.puDate = CDate(li.puDate - 1)
        End If
        
    ElseIf e = E_ECHANCIER_ONL_semaine_SCENARIO_PU Then
    
    
        wd = Application.WorksheetFunction.Weekday(li.delDate, 2)
        
        If Int(wd) = 7 Then
            li.delDate = CDate(li.delDate + 1)
        End If
        
        If Int(wd) = 6 Then
            li.delDate = CDate(li.delDate + 2)
        End If
        
        
    End If
End Sub


Public Function validateSearching(pleRng As Range, cloeRng As Range, mainRng As Range, condiRng As Range) As Boolean
    validateSearching = False
    
    If Not pleRng Is Nothing Then
        If Not cloeRng Is Nothing Then
            If Not mainRng Is Nothing Then
                If Not condiRng Is Nothing Then
                    validateSearching = True
                Else
                    'On Error Resume Next
                    'MsgBox "condi data missing for cofor exp: " & mainRng.Value, vbCritical
                    'MsgBox "Macro stops on validateSearching not able to continue - wrong data in ctrl tower db!", vbCritical
                    'End
                    
                    Debug.Print "condi data missing for cofor exp: " & mainRng.Value & mainRng.row
                End If
            Else
                'On Error Resume Next
                'MsgBox "main data missing for cofor exp: " & mainRng.Value, vbCritical
                'MsgBox "Macro stops on validateSearching not able to continue - wrong data in ctrl tower db!", vbCritical
                'End
                
                Debug.Print "main data missing for cofor exp: " & mainRng.Value & mainRng.row
            End If
        Else
            'On Error Resume Next
            'MsgBox "cloe data missing for cofor exp: " & mainRng.Value, vbCritical
            'MsgBox "Macro stops on validateSearching not able to continue - wrong data in ctrl tower db!", vbCritical
            'End
            
            Debug.Print "cloe data missing for cofor exp: " & mainRng.Value & mainRng.row
        End If
    Else
        'On Error Resume Next
        'MsgBox "PLE missing for cofor exp: " & mainRng.Value, vbCritical
        'MsgBox "Macro stops on validateSearching not able to continue - wrong data in ctrl tower db!", vbCritical
        'End
        
        Debug.Print "ple data missing for cofor exp: " & mainRng.Value & mainRng.row
    End If
End Function


Public Function removeIndice(fullPn As String) As String


    Dim arr As Variant
    
    arr = Split(fullPn, "-")
    
    removeIndice = arr(LBound(arr))
    
End Function


Public Function getMondayFromYYYYCW(yyyycw As Long) As Date
    
    getMondayFromYYYYCW = CDate(Format("1900-01-01", "yyyy-mm-dd"))
    
    
    Dim tmpdate As Date
    Dim currYear As Long
    Dim currCW As Long
    tmpdate = Date
    
    Dim currWeekday As Double
    currWeekday = Application.WorksheetFunction.Weekday(tmpdate, 2)
    tmpdate = tmpdate - currWeekday + 1
    
    currYear = CLng(Year(tmpdate))
    currCW = CLng(Application.WorksheetFunction.IsoWeekNum(tmpdate))
    
    Dim curr_yyyycw As Long
    curr_yyyycw = currYear * 100 + CLng(currCW)
    
    Dim diff As Long
    diff = yyyycw - curr_yyyycw
    
    ' to jest current monday: tmpDate
    tmpdate = tmpdate + 7 * diff
    
    getMondayFromYYYYCW = tmpdate
    
    
    
End Function


Public Function adjustStringCofor(sc As String) As String

    If sc Like "*  *" And Len(sc) > 0 Then
        adjustStringCofor = CStr(sc)
    Else
        adjustStringCofor = "X"
    End If
End Function


Public Function isThisLineIsPast(ycw As Long) As Boolean
    
    Dim currYcw As Long
    currYcw = Year(Date) * 100 + CLng(Application.WorksheetFunction.IsoWeekNum(Date))
    
    If ycw < currYcw Then
        isThisLineIsPast = True
        
    Else
        isThisLineIsPast = False
    
    End If
End Function

Public Function isThisLineIsPast2(puDate As Date) As Boolean
    
    If puDate < Date Then
        isThisLineIsPast2 = True
    Else
        isThisLineIsPast2 = False
    End If
End Function


Public Function calculateYYYYCW(s As String)


    Dim arr As Variant
    Dim cw As String
    Dim yy As String
    
    Dim yyyycw As Long
    Dim yyyy As Long
    
    
    If VBA.Left(s, 1) = "S" Then
    
        s = VBA.Mid(s, 2)
    
        arr = Split(s, "/")
        

        ' wazne na 1000 lat :)
        yyyycw = 200000
        
        cw = CStr(arr(LBound(arr)))
        yy = CStr(arr(UBound(arr)))
        
        
        yyyycw = yyyycw + CLng(cw) + (CLng(yy) * 100)
        
        calculateYYYYCW = yyyycw
    
    
    ElseIf VBA.Left(s, 2) = "CW" Then
    
        
        yyyy = Year(Date)
        yyyy = yyyy * 100
        
        s = VBA.Mid(s, 3)
        
        If IsNumeric(s) Then
            
            cw = CStr(s)
        Else
            cw = CStr(Application.WorksheetFunction.IsoWeekNum(Date))
        End If
        
        
        yyyycw = CLng(yyyy) + CLng(cw)
        
        calculateYYYYCW = yyyycw
        
    ElseIf VBA.Mid(s, 3, 3) = "-CW" Then
    
        yyyy = 2000
        
        yyyy = yyyy + CLng(CStr(VBA.Left(s, 2)))
        
        arr = Split(s, "CW")
        cw = CStr(arr(1))
        If Len(cw) = 1 Then
            cw = "0" & CStr(cw)
        End If
        
        If IsNumeric(cw) Then
            cw = CStr(cw)
        Else
            cw = CStr(Application.WorksheetFunction.IsoWeekNum(Date))
        End If
        
        
        yyyycw = CStr(yyyy) + CStr(cw)
        calculateYYYYCW = yyyycw
        
    Else
        MsgBox "Wrong format of the weeks in ECHANCIER ONL (semaine)!", vbCritical
        End
    End If
End Function







Public Sub makeMyFormulas(sh As Worksheet)
    
    Dim r As Range, lastRow As Long, wynikParsu As String
    lastRow = calcLastRowInThisPivotSource(sh)
    
    ThisWorkbook.Activate
    sh.Activate
    
    makeSpecificSumIf sh, r, CLng(lastRow), E_PIVOT_SRC_QTY, E_PIVOT_SRC_SUMIF_QTY
    makeSpecificSumIf sh, r, CLng(lastRow), E_PIVOT_SRC_CQTY, E_PIVOT_SRC_SUMIF_CQTY
    
    makeSpecificDiv sh, r, CLng(lastRow), EVO.E_PIVOT_SRC_SUMIF_QTY, E_PIVOT_SRC_SUMIF_UC, E_PIVOT_SRC_SUMIF_OQ
    makeSpecificDiv sh, r, CLng(lastRow), EVO.E_PIVOT_SRC_SUMIF_CQTY, E_PIVOT_SRC_SUMIF_UC, E_PIVOT_SRC_SUMIF_COQ
    
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC_SUMIF_QTY
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC_SUMIF_CQTY
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC_SUMIF_OQ
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC_SUMIF_COQ
    
    
    makeSpecificDivForZ3 sh, r, CLng(lastRow), _
        E_PIVOT_SRC_SUMIF_OQ, _
        E_PIVOT_SRC_UA_MAX_CAPACITY, _
        E_PIVOT_SRC_UA_PC_GV, _
        E_PIVOT_SRC__TN_ML
        
    makeSpecificDivForZ3 sh, r, CLng(lastRow), _
        E_PIVOT_SRC_SUMIF_COQ, _
        E_PIVOT_SRC_UA_MAX_CAPACITY, _
        E_PIVOT_SRC_UA_PC_GV, _
        E_PIVOT_SRC__CONFIRMED_TN_ML
        
        
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__TN_ML
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__CONFIRMED_TN_ML
    
    ' BPL  =  BPC
    makeSpecificDivForZ4 sh, r, CLng(lastRow), _
        E_PIVOT_SRC_SUMIF_OQ, _
        E_PIVOT_SRC_UA_BPC, _
        E_PIVOT_SRC_UA_PC_GV, _
        E_PIVOT_SRC__LQ
        
    makeSpecificDivForZ4 sh, r, CLng(lastRow), _
        E_PIVOT_SRC_SUMIF_COQ, _
        E_PIVOT_SRC_UA_BPC, _
        E_PIVOT_SRC_UA_PC_GV, _
        E_PIVOT_SRC__CLQ
        
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__LQ
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__CLQ
    
    
    
    makeSpecificDivForZ5 sh, r, CLng(lastRow), E_PIVOT_SRC__LQ, E_PIVOT_SRC_UA_MC, E_PIVOT_SRC__RP
    makeSpecificDivForZ5 sh, r, CLng(lastRow), E_PIVOT_SRC__CLQ, E_PIVOT_SRC_UA_MC, E_PIVOT_SRC__CRP
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__RP
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__CRP

        
    makeSpecificDivForZ6 sh, r, CLng(lastRow), E_PIVOT_SRC__RP, E_PIVOT_SRC__TN_box_ML
    makeSpecificDivForZ6 sh, r, CLng(lastRow), E_PIVOT_SRC__CRP, E_PIVOT_SRC__CONFIRMED_TN_box_ML
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__TN_box_ML
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__CONFIRMED_TN_box_ML
    
    
    
    makeSpecificDivForZ7 sh, r, CLng(lastRow), _
        E_PIVOT_SRC__TN_box_ML, _
        E_PIVOT_SRC__TN_ML, _
        E_PIVOT_SRC__RN_ML
        
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__RN_ML
    
    
    makeSpecificDivForZ7 sh, r, CLng(lastRow), _
        E_PIVOT_SRC__CONFIRMED_TN_ML, _
        E_PIVOT_SRC__CONFIRMED_TN_box_ML, _
        E_PIVOT_SRC__CONFIRMED_RN_ML
        
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__CONFIRMED_RN_ML
    
    
    
    makeSpecificDivForZ8 sh, r, CLng(lastRow), E_PIVOT_SRC__RN_ML, E_PIVOT_SRC__ROUNDUP_RN_ML
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__ROUNDUP_RN_ML
    
    makeSpecificDivForZ8 sh, r, CLng(lastRow), E_PIVOT_SRC__CONFIRMED_RN_ML, E_PIVOT_SRC__ROUNDUP_CONFIRMED_RN_ML
    changeFormulaIntoValue sh, r, CLng(lastRow), E_PIVOT_SRC__ROUNDUP_CONFIRMED_RN_ML
    
End Sub

Private Sub changeFormulaIntoValue(sh As Worksheet, r As Range, lastRow As Long, e As E_PIVOT_SRC)
    Set r = sh.Range("A2:A" & lastRow)
    Set r = r.Offset(0, e - 1)
    r.Copy
    r.PasteSpecial xlPasteValues
    sh.Range("A2").Select
    Application.CutCopyMode = False
End Sub


Private Sub makeSpecificDivForZ8(sh As Worksheet, r As Range, lastRow As Long, _
    eforX As E_PIVOT_SRC, e As E_PIVOT_SRC)
    
    
        Dim wynikParsu As String
    
        ' column A as a matchy :) by hard
        Set r = sh.Range("A2:A" & lastRow)
        wynikParsu = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("Z8").Formula
        
        wynikParsu = Replace(wynikParsu, "X", "R2C" & CStr(eforX) & ":R" & CStr(lastRow) & "C" & CStr(eforX))
        
        
        Set r = r.Offset(0, e - 1)
        r.FormulaArray = wynikParsu
    
End Sub


Private Sub makeSpecificDivForZ7(sh As Worksheet, r As Range, lastRow As Long, _
    eforX As E_PIVOT_SRC, eforY As E_PIVOT_SRC, e As E_PIVOT_SRC)
    
    
        Dim wynikParsu As String
    
        ' column A as a matchy :) by hard
        Set r = sh.Range("A2:A" & lastRow)
        wynikParsu = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("Z7").Formula
        
        wynikParsu = Replace(wynikParsu, "X", "R2C" & CStr(eforX) & ":R" & CStr(lastRow) & "C" & CStr(eforX))
        wynikParsu = Replace(wynikParsu, "Y", "R2C" & CStr(eforY) & ":R" & CStr(lastRow) & "C" & CStr(eforY))
        
        
        Set r = r.Offset(0, e - 1)
        r.FormulaArray = wynikParsu
    
End Sub



Private Sub makeSpecificDivForZ6(sh As Worksheet, r As Range, lastRow As Long, _
    eforX As E_PIVOT_SRC, e As E_PIVOT_SRC)
    
    
        Dim wynikParsu As String
    
        ' column A as a matchy :) by hard
        Set r = sh.Range("A2:A" & lastRow)
        wynikParsu = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("Z6").Formula
        
        wynikParsu = Replace(wynikParsu, "X", "R2C" & CStr(eforX) & ":R" & CStr(lastRow) & "C" & CStr(eforX))
        
        
        Set r = r.Offset(0, e - 1)
        r.FormulaArray = wynikParsu
    
End Sub



Private Sub makeSpecificDivForZ5(sh As Worksheet, r As Range, lastRow As Long, _
    eforX As E_PIVOT_SRC, eforY As E_PIVOT_SRC, _
    e As E_PIVOT_SRC)
    
    
        Dim wynikParsu As String
    
        ' column A as a matchy :) by hard
        Set r = sh.Range("A2:A" & lastRow)
        wynikParsu = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("Z5").Formula
        
        wynikParsu = Replace(wynikParsu, "X", "R2C" & CStr(eforX) & ":R" & CStr(lastRow) & "C" & CStr(eforX))
        wynikParsu = Replace(wynikParsu, "Y", "R2C" & CStr(eforY) & ":R" & CStr(lastRow) & "C" & CStr(eforY))
        
        Set r = r.Offset(0, e - 1)
        r.FormulaArray = wynikParsu
    
End Sub


Private Sub makeSpecificDivForZ4(sh As Worksheet, r As Range, lastRow As Long, _
    eforX As E_PIVOT_SRC, eforY As E_PIVOT_SRC, eforZ As E_PIVOT_SRC, _
    e As E_PIVOT_SRC)
    
    
        Dim wynikParsu As String
    
        ' column A as a matchy :) by hard
        Set r = sh.Range("A2:A" & lastRow)
        wynikParsu = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("Z4").Formula
        
        wynikParsu = Replace(wynikParsu, "X", "R2C" & CStr(eforX) & ":R" & CStr(lastRow) & "C" & CStr(eforX))
        wynikParsu = Replace(wynikParsu, "Y", "R2C" & CStr(eforY) & ":R" & CStr(lastRow) & "C" & CStr(eforY))
        wynikParsu = Replace(wynikParsu, "Z", "R2C" & CStr(eforZ) & ":R" & CStr(lastRow) & "C" & CStr(eforZ))
        
        Set r = r.Offset(0, e - 1)
        r.FormulaArray = wynikParsu
    
End Sub


Private Sub makeSpecificDivForZ3(sh As Worksheet, r As Range, lastRow As Long, _
    eforX As E_PIVOT_SRC, eforY As E_PIVOT_SRC, eforZ As E_PIVOT_SRC, _
    e As E_PIVOT_SRC)
    
        Dim wynikParsu As String
    
        ' column A as a matchy :) by hard
        Set r = sh.Range("A2:A" & lastRow)
        wynikParsu = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("Z3").Formula
        
        wynikParsu = Replace(wynikParsu, "X", "R2C" & CStr(eforX) & ":R" & CStr(lastRow) & "C" & CStr(eforX))
        wynikParsu = Replace(wynikParsu, "Y", "R2C" & CStr(eforY) & ":R" & CStr(lastRow) & "C" & CStr(eforY))
        wynikParsu = Replace(wynikParsu, "Z", "R2C" & CStr(eforZ) & ":R" & CStr(lastRow) & "C" & CStr(eforZ))
        
        Set r = r.Offset(0, e - 1)
        r.FormulaArray = wynikParsu
End Sub


Private Sub makeSpecificDiv(sh As Worksheet, r As Range, lastRow As Long, eforX As E_PIVOT_SRC, eforY As E_PIVOT_SRC, e As E_PIVOT_SRC)
    
    Dim wynikParsu As String

    ' column A as a matchy :) by hard
    Set r = sh.Range("A2:A" & lastRow)
    wynikParsu = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("Z2").Formula
    
    wynikParsu = Replace(wynikParsu, "X", "R2C" & CStr(eforX) & ":R" & CStr(lastRow) & "C" & CStr(eforX))
    wynikParsu = Replace(wynikParsu, "Y", "R2C" & CStr(eforY) & ":R" & CStr(lastRow) & "C" & CStr(eforY))
    
    Set r = r.Offset(0, e - 1)
    r.FormulaArray = wynikParsu
End Sub

Private Sub makeSpecificSumIf(sh As Worksheet, r As Range, lastRow As Long, eforZ As E_PIVOT_SRC, e As E_PIVOT_SRC)


    Dim wynikParsu As String
    
    ' column A as a matchy :) by hard
    Set r = sh.Range("A2:A" & lastRow)
    wynikParsu = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("Z1").Formula
    
    wynikParsu = Replace(wynikParsu, "X", "R2C1:R" & CStr(lastRow) & "C1")
    wynikParsu = Replace(wynikParsu, "Y", "R2C1:R" & CStr(lastRow) & "C1")
    wynikParsu = Replace(wynikParsu, "Z", "R2C" & CStr(eforZ) & ":R" & CStr(lastRow) & "C" & CStr(eforZ))
    
    Set r = r.Offset(0, e - 1)
    r.FormulaArray = wynikParsu
End Sub

Private Function calcLastRowInThisPivotSource(sh As Worksheet) As Long
    calcLastRowInThisPivotSource = 2
    
    Dim r As Range
    Set r = sh.Cells(1, 1)
    If sh.Cells(2, 1).Value <> "" Then
        Set r = r.End(xlDown)
        calcLastRowInThisPivotSource = CLng(r.row)
    End If
    
End Function


Public Function makeLeanCopyOfPivotSource(sh As Worksheet) As Worksheet


    Set makeLeanCopyOfPivotSource = Nothing

    Dim r As Range, lastRow As Long, wynikParsu As String, newSh As Worksheet
    lastRow = calcLastRowInThisPivotSource(sh)
    Set newSh = Nothing

    ' Sheets("SrcPivot_20200408_II").Select
    ' sh.Activate
    Set newSh = ThisWorkbook.Sheets.Add
    'sh.UsedRange.Copy newSh.Cells(1, 1)
    ' newSh.PasteSpecial xlValues
    'Application.CutCopyMode = False
    ' Set newSh = ActiveSheet
    'Debug.Print newSh.Name
    ' Range("A1").Select
    
    ' Debug.Print Selection.Address
    'Selection.RemoveDuplicates columns:=1, Header:=xlYes
    
    'newSh.Cells(1, 1).Select
    'quick and dirty hack
    Dim x As Variant, y As Variant
    For x = 1 To 500000
        For y = 1 To 100
        
            If sh.Cells(x, 1).Value = "" Then
                Exit For
            End If
            
            If sh.Cells(x, y).Value <> "" Then
                newSh.Cells(x, y).Value = sh.Cells(x, y).Value
            End If
        Next y
    Next x
    
    newSh.Range(newSh.Cells(1, 1), newSh.Cells(500000, 100)).Select
    Selection.RemoveDuplicates columns:=1, Header:=xlYes
    newSh.Range("A1").Select
    
    With newSh

        ' xtreme static code
        .columns("U:AA").Delete Shift:=xlToLeft
        'columns("U:U").Delete Shift:=xlToLeft
        'columns("U:U").Delete Shift:=xlToLeft
        'columns("U:U").Delete Shift:=xlToLeft
        
        ' route and pilot as seperate columns
        columns("R:S").Delete Shift:=xlToLeft
        
    End With
    
    On Error Resume Next
    Set makeLeanCopyOfPivotSource = newSh
End Function


Public Sub makePivot(ByRef ph As PivotHandler)
    
    Dim r As Range, lastRow As Long, wynikParsu As String, localPivotSh As Worksheet
    Dim sh As Worksheet
    
    
    Set sh = ph.proxy2
    Set localPivotSh = ph.getPivot()
    
    lastRow = calcLastRowInThisPivotSource(sh)


    Set r = sh.UsedRange
    ' Looks like it is OK!
    Debug.Print r.Address
    Debug.Print localPivotSh.name
    
    ph.setRangeDataFromProxy2ForPivot r
    ' ==================================================
    ph.runMainLogicForCreationPivotTable
    ' ==================================================
    
    
End Sub



Public Sub makeTPivot(ByRef ph As PivotHandler)
    
    Dim r As Range, lastRow As Long, wynikParsu As String, localPivotSh As Worksheet
    Dim sh As Worksheet
    
    
    Set sh = ph.proxy2
    Set localPivotSh = ph.getPivot()
    
    lastRow = calcLastRowInThisPivotSource(sh)


    Set r = sh.UsedRange
    ' Looks like it is OK!
    Debug.Print r.Address
    Debug.Print localPivotSh.name
    
    ph.setRangeDataFromProxy2ForPivot r
    ' ==================================================
    ph.runMainLogicForCreationTheoreticPivotTable
    ' ==================================================
    
    
End Sub



Public Sub makePivot2(ByRef ph As PivotHandler2)
    
    Dim r As Range, lastRow As Long, wynikParsu As String, localPivotSh As Worksheet
    Dim sh As Worksheet
    
    
    Set sh = ph.proxy2
    Set localPivotSh = ph.getPivot()
    
    lastRow = calcLastRowInThisPivotSource(sh)


    Set r = sh.UsedRange
'    ' Looks like it is OK!
    Debug.Print r.Address
    Debug.Print localPivotSh.name

    ph.setRangeDataFromProxy2ForPivot r
    ' ==================================================
    ph.runMainLogicForCreationPivotTable
    ' ==================================================
    
    
End Sub



Public Sub makeTPivot2(ByRef ph As PivotHandler2)
    
    Dim r As Range, lastRow As Long, wynikParsu As String, localPivotSh As Worksheet
    Dim sh As Worksheet
    
    
    Set sh = ph.proxy2
    Set localPivotSh = ph.getPivot()
    
    lastRow = calcLastRowInThisPivotSource(sh)
    
    Set r = sh.UsedRange
    ' Looks like it is OK!
    Debug.Print r.Address
    Debug.Print localPivotSh.name
    
    ph.setRangeDataFromProxy2ForPivot r
    ' ==================================================
    ph.runMainLogicForCreationTheoreticPivotTable
    ' ==================================================
    
    
End Sub



Public Sub copyThisPivotToTcamReport(ByRef ph As PivotHandler, Optional e As E_PIVOT_STD)

    
    ph.tcamSheet.Cells(1, 1).Value = "TCAM REPORT"
    ph.tcamSheet.Cells(2, 1).Value = CStr(Now)
    
    ph.tcamSheet.Cells(1, 4).Value = "PROXY2"
    ph.tcamSheet.Cells(1, 5).Value = ph.proxy2.name
    
    ph.tcamSheet.Cells(2, 4).Value = "PIVOT"
    ph.tcamSheet.Cells(2, 5).Value = ph.getPivot.name
    
    
    
    

    With ph.tcamSheet
        .Range("A1:A2").Interior.Color = RGB(221, 235, 247)
        .Range("D1:E2").Interior.Color = RGB(221, 235, 247)
        .Range("D1:E2").Font.Size = 8
        
        .Range("G1").Value = "starting point data: "
        .Range("G2").Value = "columns: "
        .Range("G1:H2").Interior.Color = RGB(221, 235, 247)
        .Range("G1:H2").Font.Size = 8
    End With
    
    
    Dim ileMamyKolumn As Long
    Dim ostatniWiersz As Long
    Dim startingPointData As Range
    Dim r1 As Range, xtraLabelsMountPoint As Range
    Set r1 = ph.getPivot().Cells(1, 1)
    
    Dim x As Integer
    x = 4
    
    
    ' FILTER
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
    Do
        With ph.tcamSheet
            .Cells(x, 1).Value = r1.Value
            .Cells(x, 1).Font.Bold = True
            .Cells(x, 1).Interior.Color = RGB(221, 235, 247)
            .Cells(x, 2).Value = r1.Offset(0, 1).Value
            .Cells(x, 2).Font.Bold = True
            .Cells(x, 2).Interior.Color = RGB(221, 235, 247)
        End With
        x = x + 1
        Set r1 = r1.Offset(1, 0)
    Loop Until CStr(r1) = ""
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
    
    ' data main body
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
    x = x + 2
    Set r1 = ph.getPivot().Cells(5, 1)
    
    With ph.tcamSheet
        r1.Copy .Cells(x, 1)
        x = x + 1
        Set r1 = r1.Offset(1, 0)
        
        Set startingPointData = .Cells(x, 1)
        Debug.Print "startingPointData: " & startingPointData.Address
        .Range("H1").Value = startingPointData.row
        
        Set r1 = ph.getPivot().Range(r1, r1.End(xlToRight).Offset(0, -1))
        ' first - copy labels!
        r1.Copy .Cells(x, 1)
        Set xtraLabelsMountPoint = .Cells(x, 1)
        Set xtraLabelsMountPoint = xtraLabelsMountPoint.End(xlToRight).Offset(0, 1)
        addXtraLabelsInTcamReport xtraLabelsMountPoint
        
        ileMamyKolumn = CLng(xtraLabelsMountPoint.Offset(0, 3).Column)
        'kolumny
        .Range("H2").Value = ileMamyKolumn
        
        
        x = x + 1
        Set r1 = r1.Offset(1, 0)
        
        
        ' DATA
        If CStr(r1.Cells(1, 1).Value) <> "" Then
        
            Do
                r1.Copy .Cells(x, 1)
                x = x + 1
                Set r1 = r1.Offset(1, 0)
            Loop Until CStr(r1.Cells(1, 1).Value) = ""
        End If
        
        ostatniWiersz = x - 1
        
        
        
        ' still in tcam but now we need to beautify data
        ' czyli some dekoracjon
        
        ' zaczynamy od:
        Dim deco_r As Range
        
        Set deco_r = startingPointData
        
        Do
            ' ------------------------------------------------------------------------
            
            If deco_r.row < ostatniWiersz Then
                If IsDate(deco_r) Then
                    usun_sumy_dla_dat_bo_to_bez_sensu deco_r, deco_r.Offset(0, 1), ileMamyKolumn
                ElseIf CStr(deco_r) <> "" Then
                    zaokragliWartosci startingPointData, deco_r, deco_r.Offset(0, 1), ileMamyKolumn
                Else
                    ' nop ?
                End If
            Else
                'ostatniWiersz
                deco_r.Value = "SUM: "
                wsadzFormulySumy deco_r, deco_r.Offset(0, 1), startingPointData, ileMamyKolumn
                
                If e = E_PIVOT_ANGEL_STD Then
                    wsadzFormulyDlaCloeData ph.tcamSheet, deco_r, deco_r.Offset(0, 1), startingPointData, ileMamyKolumn
                ElseIf e = E_PIVOT_CAROLINE_STD Then
                    wsadzFormulyDlaCloeData_CarolineApproach ph.tcamSheet, deco_r, deco_r.Offset(0, 1), startingPointData, ileMamyKolumn
                Else
                    MsgBox "No good std in this section! Call dev! :D", vbCritical
                End If
            End If
            ' ------------------------------------------------------------------------
            Set deco_r = deco_r.Offset(1, 0)
        Loop Until CStr(deco_r) = ""
        
    End With
    
    
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
End Sub


Private Sub usun_sumy_dla_dat_bo_to_bez_sensu(dateRange As Range, tmpDecoRange As Range, imk As Long)


    
    ' imk stands for copy from ileMamyKolumn
    
    
    Do
        
        If IsDate(dateRange) And IsNumeric(tmpDecoRange) Then
            
            tmpDecoRange.Value = ""
        End If
        
        Set tmpDecoRange = tmpDecoRange.Offset(0, 1)
    Loop While tmpDecoRange.Column <= imk
End Sub


Private Sub zaokragliWartosci(sp As Range, dataRange As Range, tmpDecoRange As Range, imk As Long)


    ' imk stands for copy from ileMamyKolumn
    
    Dim tmpValue As Double
    
    
    Do
    
        If tmpDecoRange.row > sp.row Then
        
            If IsNumeric(tmpDecoRange) Then
                
                ' tmpDecoRange.Value = ""
                
                If CStr(tmpDecoRange) <> "" Then
                    With tmpDecoRange
                        tmpValue = CDbl(.Value)
                        .AddComment CStr(tmpValue)
                        .Value = CStr(Application.WorksheetFunction.RoundUp(tmpValue, 0))
                    End With
                End If
            End If
        End If
        
        Set tmpDecoRange = tmpDecoRange.Offset(0, 1)
    Loop While tmpDecoRange.Column <= imk
End Sub


Private Sub wsadzFormulySumy(fstSumRange As Range, tmpDecoRange As Range, sp As Range, imk As Long)


    ' sp stands for starting point
    ' imk stands for copy from ileMamyKolumn
    
    Dim tmpValue As Double, tmpFromFormula As String, adr1 As String, adr2 As String
    Dim columnDiff As Long
    Dim formulaFromRegister As Range
    Set formulaFromRegister = ThisWorkbook.Sheets("register").Range("Z9")
    
    
    
    
    Do
        
        If IsNumeric(tmpDecoRange) Then
            
            ' tmpDecoRange.Value = ""
            
            If CStr(tmpDecoRange) <> "" Then
            
                tmpFromFormula = CStr(formulaFromRegister.Formula)
                
                columnDiff = tmpDecoRange.Column - sp.Column
                
                adr1 = CStr(sp.Offset(1, columnDiff).Address)
                adr2 = CStr(tmpDecoRange.Offset(-1, 0).Address)
                tmpFromFormula = Replace(tmpFromFormula, "X", adr1 & ":" & adr2)
            
            
            
            
                With tmpDecoRange
                    tmpValue = CDbl(.Value)
                    .AddComment CStr(tmpValue)
                    .Formula = CStr(tmpFromFormula)
                End With
            End If
        End If
        
        Set tmpDecoRange = tmpDecoRange.Offset(0, 1)
    Loop While tmpDecoRange.Column <= imk
    
    
    ' troche glupie, ale napewno 4 kolumny sa statyczne na bank!
    Set tmpDecoRange = tmpDecoRange.Offset(0, -4)
    
    
    fstSumRange.Copy tmpDecoRange
    tmpDecoRange.Value = "TRANS"
    
    
    Set tmpDecoRange = tmpDecoRange.Offset(0, 1)
    
    tmpFromFormula = CStr(formulaFromRegister.Formula)
                
    columnDiff = tmpDecoRange.Column - sp.Column
            
    adr1 = CStr(sp.Offset(1, columnDiff).Address)
    adr2 = CStr(tmpDecoRange.Offset(-1, 0).Address)
    tmpFromFormula = Replace(tmpFromFormula, "X", adr1 & ":" & adr2)
    tmpDecoRange.Formula = CStr(tmpFromFormula)
    tmpDecoRange.Interior.Color = RGB(221, 235, 247)
    tmpDecoRange.Font.Bold = True
    
    
    Set tmpDecoRange = tmpDecoRange.Offset(0, 1)
    fstSumRange.Copy tmpDecoRange
    tmpDecoRange.Value = "INC"
    
    
    Set tmpDecoRange = tmpDecoRange.Offset(0, 1)
    
    tmpFromFormula = CStr(formulaFromRegister.Formula)
                
    columnDiff = tmpDecoRange.Column - sp.Column
            
    adr1 = CStr(sp.Offset(1, columnDiff).Address)
    adr2 = CStr(tmpDecoRange.Offset(-1, 0).Address)
    tmpFromFormula = Replace(tmpFromFormula, "X", adr1 & ":" & adr2)
    tmpDecoRange.Formula = CStr(tmpFromFormula)
    tmpDecoRange.Interior.Color = RGB(221, 235, 247)
    tmpDecoRange.Font.Bold = True
    
    
End Sub

Private Sub wsadzFormulyDlaCloeData_CarolineApproach(tcamSh As Worksheet, fstSumRange As Range, tmpDecoRange As Range, sp As Range, imk As Long)
    
    ' sp stands for starting point
    ' imk stands for copy from ileMamyKolumn
    
    Dim tmpValue As Double, tmpFromFormula As String, adr1 As String, adr2 As String
    Dim columnMount As Long, columnDiff As Long
    Dim formulaFromRegister As Range, rangeForCloeFormula As Range
    Set formulaFromRegister = ThisWorkbook.Sheets("register").Range("Z20")
    Set rangeForCloeFormula = Nothing
    
    columnMount = CLng(imk - 4)
    
    tmpFromFormula = CStr(formulaFromRegister.FormulaR1C1Local)
    
    Set tmpDecoRange = tmpDecoRange.Offset(0, columnMount - 1)
    
    adr1 = CStr(sp.Offset(1, columnMount).Address)
    adr2 = CStr(tmpDecoRange.Offset(-1, 0).Address)
    
    Debug.Print "adrs: " & adr1 & ":" & adr2
    
    With tcamSh.Range("" & adr1 & ":" & adr2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=register!$T$1:$T$" & _
            CStr(ThisWorkbook.Sheets("register").Range("T1").End(xlDown).row)
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    
    Set rangeForCloeFormula = tcamSh.Range("" & adr1 & ":" & adr2).Offset(0, 1)
    tmpFromFormula = Replace(tmpFromFormula, "X", "1")
    rangeForCloeFormula.FormulaR1C1Local = tmpFromFormula

End Sub



Private Sub wsadzFormulyDlaCloeData(tcamSh As Worksheet, fstSumRange As Range, tmpDecoRange As Range, sp As Range, imk As Long)


    ' sp stands for starting point
    ' imk stands for copy from ileMamyKolumn
    
    Dim tmpValue As Double, tmpFromFormula As String, adr1 As String, adr2 As String
    Dim columnMount As Long, columnDiff As Long
    Dim formulaFromRegister As Range, rangeForCloeFormula As Range
    Set formulaFromRegister = ThisWorkbook.Sheets("register").Range("Z10")
    Set rangeForCloeFormula = Nothing
    
    columnMount = CLng(imk - 4)
    
    tmpFromFormula = CStr(formulaFromRegister.FormulaR1C1Local)
    
    Set tmpDecoRange = tmpDecoRange.Offset(0, columnMount - 1)
    
    adr1 = CStr(sp.Offset(1, columnMount).Address)
    adr2 = CStr(tmpDecoRange.Offset(-1, 0).Address)
    
    Debug.Print "adrs: " & adr1 & ":" & adr2
    
    With tcamSh.Range("" & adr1 & ":" & adr2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=register!$T$1:$T$" & _
            CStr(ThisWorkbook.Sheets("register").Range("T1").End(xlDown).row)
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    
    Set rangeForCloeFormula = tcamSh.Range("" & adr1 & ":" & adr2).Offset(0, 1)
    tmpFromFormula = Replace(tmpFromFormula, "X", "1")
    rangeForCloeFormula.FormulaR1C1Local = tmpFromFormula

End Sub




Private Sub addXtraLabelsInTcamReport(lr As Range)
    
    lr.Value = "TRANSPORT"
    
    lr.Offset(0, 1).Value = "COST"
    lr.Offset(0, 2).Value = "INCIDENT"
    lr.Offset(0, 3).Value = "INCIDENT COST"
    
    With lr.Parent.Range(lr, lr.Offset(0, 3))
        .Font.Bold = True
        .Interior.Color = CLng(lr.Offset(0, -1).Interior.Color)
        .EntireColumn.ColumnWidth = 20
    End With
    
    
    lr.Parent.Cells(lr.row, 1).EntireColumn.ColumnWidth = 40
    
    Dim leftLabels As Range
    Set leftLabels = lr.Parent.Cells(lr.row, 2)
    Do
        leftLabels.EntireColumn.ColumnWidth = 12
        Set leftLabels = leftLabels.Offset(0, 1)
    Loop Until leftLabels.Address = lr.Address
    
End Sub
