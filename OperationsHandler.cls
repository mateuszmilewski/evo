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



Public Function recalculDatesIfSatOrSun(ByRef li As LineItem)
    
    
    Dim d1 As Date
    Dim d2 As Date
    ' iteration date
    Dim id As Date
    
    d1 = li.puDate
    d2 = li.delDate
    
    id = d1
    Dim howManyDaysOffOffset As Long, wd As Integer
    howManyDaysOffOffset = 0
    Do
        wd = Int(Application.WorksheetFunction.Weekday(id, 2))
        
        If Int(wd) = 6 Or Int(wd) = 7 Then
            howManyDaysOffOffset = howManyDaysOffOffset + 1
        End If
        id = id + 1
    Loop While id < (d2 + 1)
    
    li.puDate = CDate(li.puDate - howManyDaysOffOffset)
    
    wd = Application.WorksheetFunction.Weekday(li.puDate, 2)
    
    If Int(wd) = 7 Then
        li.puDate = CDate(li.puDate - 2)
    End If
    
    If Int(wd) = 6 Then
        li.puDate = CDate(li.puDate - 1)
    End If
End Function


Public Function validateSearching(pleRng As Range, cloeRng As Range, mainRng As Range, condiRng As Range) As Boolean
    validateSearching = False
    
    If Not pleRng Is Nothing Then
        If Not cloeRng Is Nothing Then
            If Not mainRng Is Nothing Then
                If Not condiRng Is Nothing Then
                    validateSearching = True
                End If
            End If
        End If
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
    
    
    If VBA.Left(s, 1) = "S" Then
    
        s = VBA.Mid(s, 2)
    
        arr = Split(s, "/")
        
        Dim cw As String
        Dim yy As String
        
        Dim yyyycw As Long
        ' wazne na 1000 lat :)
        yyyycw = 200000
        
        cw = CStr(arr(LBound(arr)))
        yy = CStr(arr(UBound(arr)))
        
        
        yyyycw = yyyycw + CLng(cw) + (CLng(yy) * 100)
        
        calculateYYYYCW = yyyycw
    
    Else
        MsgBox "Wrong format of the weeks in ECHANCIER ONL (semaine)!", vbCritical
        End
    End If
End Function







Public Sub makeMyFormulas(sh As Worksheet)
    
    Dim r As Range, lastRow As Long, wynikParsu As String
    lastRow = calcLastRowInThisPivotSource(sh)
    
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
    
End Sub

Private Sub changeFormulaIntoValue(sh As Worksheet, r As Range, lastRow As Long, e As E_PIVOT_SRC)
    Set r = sh.Range("A2:A" & lastRow)
    Set r = r.Offset(0, e - 1)
    r.Copy
    r.PasteSpecial xlPasteValues
    sh.Range("A2").Select
    Application.CutCopyMode = False
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
        calcLastRowInThisPivotSource = CLng(r.Row)
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
    sh.UsedRange.Copy
    newSh.Paste
    Application.CutCopyMode = False
    ' Set newSh = ActiveSheet
    'Debug.Print newSh.Name
    ' Range("A1").Select
    Selection.RemoveDuplicates columns:=1, Header:=xlYes
    
    newSh.Cells(1, 1).Select
    
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

