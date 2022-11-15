VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelessLeafHandler"
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


Private greenLightMode As Boolean
Private receptionMode As Boolean

Private frm1 As ModelessLeaf

Private currentColumnToFill As Integer

Private wrk1 As Workbook
Private sh1 As Worksheet


Public Sub setMode(gl As Boolean, rc As Boolean, mfrm1 As ModelessLeaf)
    
    
    ' only one can be True - a little bit confusing but it is like it is...
    greenLightMode = gl
    receptionMode = rc
    
    Set frm1 = mfrm1
End Sub

Private Sub Class_Initialize()
    greenLightMode = False
    receptionMode = False
    
    Set frm1 = Nothing
    
    Set sh1 = Nothing
End Sub

Private Sub Class_Terminate()
    greenLightMode = False
    receptionMode = False
    
    Set frm1 = Nothing
    
    Set sh1 = Nothing
End Sub


Public Sub createLeaf()
    
    If greenLightMode Then
        innerCreateLeafForGreenLight
    End If
    
    If receptionMode Then
        innerCreateLeafForReception
    End If
End Sub

Public Sub addLeafToExisitingOne(nameOfWorkbook As String)

    Set wrk1 = Workbooks(nameOfWorkbook)
    
    If greenLightMode Then
        innerCreateLeafForExistingGreenLight
    End If
    
    If receptionMode Then
        innerCreateLeafForExistingReception
    End If
End Sub

Public Sub createLeafs()
    
    
    
    
    ' make some configuration of all leafs
    ' loop through to make leaf after leaf
    Dim i As Variant, str As String
    For i = 0 To frm1.ListBoxScope.ListCount - 1
        
        
        str = CStr(frm1.ListBoxScope.list(i))
        frm1.ListBoxScope.Selected(i) = True
        ' side effect hell
        createLeaf
    Next i
    
    
    
    ' at the end create some total
    ' ===================================================
    
    ' TOTAL
    ' ===================================================
    
    ' all leafs created - make total
    makeTotal
    
    ' ===================================================
    
    makeSomeFinalColors
    
End Sub

Public Sub makeTotal()


    Application.Calculation = xlCalculationManual

    Dim ref As Range
    Set ref = setNewPlaceForTotal()
    If ref.Value = "" Then
        ' put fresh total
        ' ----------------------------------------------------------------
        
        Dim startingPoint As Range, startingColumn As String
        Set startingPoint = sh1.Range("C2")
        startingColumn = CStr("C")
        
        ref.Value = "TOTAL"
        
        Dim col As Integer
        col = ref.Column
        
        Dim strF As String
        
        
        ' INTERNAL SUM of EUR and COUNT
        '--------------------------------------------------------------------------------------------------
        strF = CStr(ThisWorkbook.Sheets("register").Range("Z23").Formula)
        strF = Replace(strF, "X", startingColumn & CStr(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_INTERNAL))
        strF = Replace(strF, "Y", sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_INTERNAL, col - 1).Address)
        strF = Replace(strF, "$", "")
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_INTERNAL, col).Formula = strF
        
        strF = CStr(ThisWorkbook.Sheets("register").Range("Z23").Formula)
        strF = Replace(strF, "X", startingColumn & CStr(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_INTERNAL))
        strF = Replace(strF, "Y", sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_INTERNAL, col - 1).Address)
        strF = Replace(strF, "$", "")
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_INTERNAL, col).Formula = strF
        '--------------------------------------------------------------------------------------------------
        
        
        ' Mass € Forecast without Tango price
        '--------------------------------------------------------------------------------------------------
        strF = CStr(ThisWorkbook.Sheets("register").Range("Z23").Formula)
        strF = Replace(strF, "X", startingColumn & CStr(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_WO_TANGO))
        strF = Replace(strF, "Y", sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_WO_TANGO, col - 1).Address)
        strF = Replace(strF, "$", "")
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_WO_TANGO, col).Formula = strF
        '--------------------------------------------------------------------------------------------------
        
        
        ' % mass €  without Tango price
        '--------------------------------------------------------------------------------------------------
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_PERC_MASS_EUR_WO_TANGO, col).FormulaR1C1Local = _
            sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_PERC_MASS_EUR_WO_TANGO, col - 2).FormulaR1C1Local
        
        '--------------------------------------------------------------------------------------------------
        
        
        ' number ref without Tango price
        '--------------------------------------------------------------------------------------------------
        strF = CStr(ThisWorkbook.Sheets("register").Range("Z23").Formula)
        strF = Replace(strF, "X", startingColumn & CStr(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_WO_TANGO))
        strF = Replace(strF, "Y", sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_WO_TANGO, col - 1).Address)
        strF = Replace(strF, "$", "")
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_WO_TANGO, col).Formula = strF
        '--------------------------------------------------------------------------------------------------
        
        
        ' rate  : Mass Sigapp pre-serial price / initial price (without Tango price)
        '--------------------------------------------------------------------------------------------------
        
        ' average
        strF = CStr(ThisWorkbook.Sheets("register").Range("Z24").Formula)
        strF = Replace(strF, "X", startingColumn & CStr(EVO.E_MODELESS_LEAF_ORDER_RATE_WO_TANGO))
        strF = Replace(strF, "Y", sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_RATE_WO_TANGO, col - 1).Address)
        strF = Replace(strF, "$", "")
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_RATE_WO_TANGO, col).Formula = strF
        
        '--------------------------------------------------------------------------------------------------
        
        
        
        ' NEW
        ' E_MODELESS_LEAF_ORDER_EMPTY_1 WO TANGO TARGET PRICE BY INITIAL PRICE FROM TP04
        '
        ' Mass € target tp04 Price * 1
        '
        '--------------------------------------------------------------------------------------------------
        
        ' average
        strF = CStr(ThisWorkbook.Sheets("register").Range("Z23").Formula)
        strF = Replace(strF, "X", startingColumn & CStr(EVO.E_MODELESS_LEAF_ORDER_EMPTY_1))
        strF = Replace(strF, "Y", sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_EMPTY_1, col - 1).Address)
        strF = Replace(strF, "$", "")
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_EMPTY_1, col).Formula = strF
        
        '--------------------------------------------------------------------------------------------------
        
        
        ' Mass € Sigapp forecast (With Tango price)
        '--------------------------------------------------------------------------------------------------
        strF = CStr(ThisWorkbook.Sheets("register").Range("Z23").Formula)
        strF = Replace(strF, "X", startingColumn & CStr(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_W_TANGO))
        strF = Replace(strF, "Y", sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_W_TANGO, col - 1).Address)
        strF = Replace(strF, "$", "")
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_W_TANGO, col).Formula = strF
        '--------------------------------------------------------------------------------------------------


        ' number ref with Tango price
        '--------------------------------------------------------------------------------------------------
        strF = CStr(ThisWorkbook.Sheets("register").Range("Z23").Formula)
        strF = Replace(strF, "X", startingColumn & CStr(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO))
        strF = Replace(strF, "Y", sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO, col - 1).Address)
        strF = Replace(strF, "$", "")
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO, col).Formula = strF
        '--------------------------------------------------------------------------------------------------
        
        
        ' Including number no price * 1
        '--------------------------------------------------------------------------------------------------
        strF = CStr(ThisWorkbook.Sheets("register").Range("Z23").Formula)
        strF = Replace(strF, "X", startingColumn & CStr(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO_NOK))
        strF = Replace(strF, "Y", sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO_NOK, col - 1).Address)
        strF = Replace(strF, "$", "")
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO_NOK, col).Formula = strF
        '--------------------------------------------------------------------------------------------------
        
        
        ' Mass € target Price * 1
        '--------------------------------------------------------------------------------------------------
        strF = CStr(ThisWorkbook.Sheets("register").Range("Z23").Formula)
        strF = Replace(strF, "X", startingColumn & CStr(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_TARGET_PRICE))
        strF = Replace(strF, "Y", sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_TARGET_PRICE, col - 1).Address)
        strF = Replace(strF, "$", "")
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_TARGET_PRICE, col).Formula = strF
        '--------------------------------------------------------------------------------------------------
        
        ' rate
        ' rate  : Mass forecast Sigapp price /target price * 1 (with Tango price)
        '--------------------------------------------------------------------------------------------------
        
        ' Debug.Print strF ' OK
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_RATE, col).FormulaR1C1Local = _
            sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_RATE, col - 2).FormulaR1C1Local
        

        '--------------------------------------------------------------------------------------------------
        
        
        ' GAP: Mass € gap Sigapp price / price *1
        '--------------------------------------------------------------------------------------------------
        
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_GAP, col).FormulaR1C1Local = _
            sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_GAP, col - 2).FormulaR1C1Local
        '--------------------------------------------------------------------------------------------------
        
        
        
        ' new 2 rows = extra ratio + extra gap supplied by data from no tango articles
        '--------------------------------------------------------------------------------------------------
        
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_GAP_FINAL, col).FormulaR1C1Local = _
            sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_GAP_FINAL, col - 2).FormulaR1C1Local
        
        sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_RATE_FINAL, col).FormulaR1C1Local = _
            sh1.Cells(EVO.E_MODELESS_LEAF_ORDER_RATE_FINAL, col - 2).FormulaR1C1Local
            
        
        '--------------------------------------------------------------------------------------------------
        
        
        ' ----------------------------------------------------------------
    Else
        
        ' there is something
        ' so... overwrite or more sophi.. scenario?
        ' ----------------------------------------------------------------
        MsgBox "TOTAL IS ON PLACE ALREADY!"
        ' ----------------------------------------------------------------
        
    End If
    
    If Not ref Is Nothing Then
    Else
        MsgBox "Can not define place for total formulas inside synthesis worksheet"
    End If
    
    
    Application.Calculation = xlCalculationAutomatic

End Sub

Private Function setNewPlaceForTotal() As Range
    Set setNewPlaceForTotal = Nothing
    
    ' check if worksheet is in std
    If Me.checkIfGreenLightLabelsAreInStd(sh1) Or Me.checkIfReceptionLabelsAreInStd(sh1) Then
    
        Debug.Print "sh1 is in std!"
        
        Dim tmpr As Range, ir As Range, txt As String
        Set tmpr = sh1.Range("C2:C7")
        
        Dim iter As Integer
        iter = 0
        ' bardzo na okolo, ale bedzie dzialac napewno dobrze
        Do
            txt = ""
            For Each ir In tmpr
                
                txt = txt & CStr(ir.Value)
                
            Next ir
            
            If Len(txt) = 0 Then
                iter = iter + 1
            End If
            
            
            Set tmpr = tmpr.offset(0, 1)
         
        ' one empty column between data and total column!
        Loop While iter < 1
        
        Set setNewPlaceForTotal = tmpr.Cells(1, 1)
        
    Else
        Set setNewPlaceForTotal = Nothing
    End If
End Function

Private Sub makeSomeFinalColors()
    
    ' prepare rng
    iinerColorsColors 3, Int(currentColumnToFill)
    
End Sub


Private Sub iinerColorsColors(col1 As Integer, col2 As Integer)

    
    ' get column letter
    'a -65, so starting point will 64 =>
    ' 64 + 1 = A
    ' 64 + 2 = B
    ' 64 + 3 = C ... you how it goes :)
    
    Dim letterForCol1 As String, letterForCol2 As String, letterForTot As String
    letterForCol1 = Chr(64 + col1)
    
    If col2 + 2 < 27 Then
        letterForCol2 = Chr(64 + col2 + 2)
    Else
        letterForCol2 = "A" & Chr(64 + col2 + 2 - 26)
    End If
    ' not used!
    'letterForTot = Chr(64 + col2 + 2)
    
    
    


    Dim wrArr(1 To 3) As Range
    
    Set wrArr(1) = sh1.Range("B8")
    Set wrArr(2) = sh1.Range("B15")
    Set wrArr(3) = sh1.Range("B18")


    Dim adr_str As String, rng1 As Range
    
    ' first adr
    adr_str = "B3:B19"
    
    sh1.columns("B:B").ColumnWidth = 38
    Set rng1 = sh1.Range(adr_str)
    rng1.Borders(xlDiagonalDown).LineStyle = xlNone
    rng1.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With rng1.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    
    ' wrapping row labels for rates
    ' ----------------------------------------------
    
    
    Dim wrapRngIter As Variant
    Dim wrapRngRngIter As Range
    
    For wrapRngIter = LBound(wrArr) To UBound(wrArr)
    
        Set wrapRngRngIter = Nothing
        
        On Error Resume Next
        Set wrapRngRngIter = wrArr(wrapRngIter)
        
        If Not wrapRngRngIter Is Nothing Then
            With wrapRngRngIter
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlBottom
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
        End If
    Next wrapRngIter
    ' ----------------------------------------------
    
    Dim secondRow As Range
    Set secondRow = sh1.Range("" & CStr(letterForCol1) & "2:" & CStr(letterForCol2) & "2")
    'should sth like this:
    ' Range("C2:R2").Select
    secondRow.Borders(xlDiagonalDown).LineStyle = xlNone
    secondRow.Borders(xlDiagonalUp).LineStyle = xlNone
    With secondRow.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With secondRow.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With secondRow.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With secondRow.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With secondRow.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With secondRow.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With secondRow
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With secondRow
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    columns("" & letterForCol1 & ":" & letterForCol2).ColumnWidth = 10
    ' columns("C:R").Select
    'Selection.ColumnWidth = 10
    
    ' first section of the data
    Set rng1 = sh1.Range("" & letterForCol1 & "3:" & letterForCol2 & "8")
    
    With rng1
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    With rng1
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    rng1.Borders(xlDiagonalDown).LineStyle = xlNone
    rng1.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With rng1.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    
    
    
    ' Range("B7:R7").Select
    Set rng1 = sh1.Range("B7:" & letterForCol2 & "7")
    
    With rng1.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
    End With
    
    ' Range("B4:R4").Select
    Set rng1 = sh1.Range("B4:" & letterForCol2 & "4")
    
    With rng1.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
    End With
    
    ' Range("C10:R12").Select
    ' Set rng1 = sh1.Range("" & letterForCol1 & "3:" & letterForCol2 & "8")
    Set rng1 = sh1.Range("" & letterForCol1 & "3:" & letterForCol2 & "12")
    
    rng1.Borders(xlDiagonalDown).LineStyle = xlNone
    rng1.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With rng1.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ' Range("B11:R12").Select
    ' Set rng1 = sh1.Range("" & letterForCol1 & "3:" & letterForCol2 & "8")
    Set rng1 = sh1.Range("B11:" & letterForCol2 & "12")
    
    With rng1
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With rng1
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With rng1.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
    End With
    
    ' Range("C14:R16").Select
    ' Set rng1 = sh1.Range("" & letterForCol1 & "3:" & letterForCol2 & "8")
    Set rng1 = sh1.Range("" & letterForCol1 & "14:" & letterForCol2 & "19")
    
    
    With rng1
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    With rng1
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    rng1.Borders(xlDiagonalDown).LineStyle = xlNone
    rng1.Borders(xlDiagonalUp).LineStyle = xlNone
    
    
    With rng1.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With rng1.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    'Range("C15:R15").Select
    'Range("R15").Activate
    ' ' Set rng1 = sh1.Range("" & letterForCol1 & "3:" & letterForCol2 & "8")
    Set rng1 = sh1.Range("" & letterForCol1 & "15:" & letterForCol2 & "15")
    
    rng1.Font.Bold = True
    rng1.NumberFormat = "0.00"
    
    Set rng1 = sh1.Range("" & letterForCol1 & "18:" & letterForCol2 & "18")
    
    rng1.Font.Bold = True
    rng1.NumberFormat = "0.00"
    
    
    
    
    'Range("C16:R16").Select
    'Range("R16").Activate
    Set rng1 = sh1.Range("" & letterForCol1 & "16:" & letterForCol2 & "16")
    
    With rng1.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Set rng1 = sh1.Range("" & letterForCol1 & "19:" & letterForCol2 & "19")
    
    With rng1.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    
    
    'Range("C8:R8").Select
    'Range("R8").Activate
    Set rng1 = sh1.Range("" & letterForCol1 & "8:" & letterForCol2 & "8")
    rng1.NumberFormat = "0.00"
    
    ' Range("B15:R15").Select
    Set rng1 = sh1.Range("B15:" & letterForCol2 & "15")
    
    With rng1.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    Set rng1 = sh1.Range("B18:" & letterForCol2 & "18")
    
    With rng1.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    ' Range("B3").Select
    
    
    'Range("C6:R6").Select
    'Selection.Style = "Percent"
    Set rng1 = sh1.Range("" & letterForCol1 & "6:" & letterForCol2 & "6")
    rng1.Style = "Percent"
    
    
    
    ' if reception hide rate for without tango
    If receptionMode Then
        ' not the case anymore
        ' sh1.rows("8:8").Hidden = True
    End If
    
    If greenLightMode Then
        sh1.rows("18:19").Hidden = True
    End If
    
End Sub


Private Sub innerCreateLeafForExistingGreenLight()
    
    Dim r As Range
    Set r = makeGreenLightColumnAtTheEnd()
    
End Sub

Private Sub innerCreateLeafForExistingReception()

    Dim r As Range
    Set r = makeReceptionColumnAtTheEnd()
End Sub

Private Sub innerCreateLeafForGreenLight()
    ' ------------------------------------------
    
    ' MAIN LOGIC FOR GREEN LIGHT
    ' ------------------------------------------
    Dim r As Range
    Set r = makeGreenLightScaffold()
    
    ' ------------------------------------------
End Sub

Private Sub innerCreateLeafForReception()
    ' ------------------------------------------
    
    ' MAIN LOGIC FOR RECEPTION
    ' ------------------------------------------
    Dim r As Range
    Set r = makeReceptionScaffold()
    ' ------------------------------------------
End Sub


Private Function makeGreenLightColumnAtTheEnd() As Range
    Set makeGreenLightColumnAtTheEnd = Nothing
    
    Dim sh As Worksheet
    Set sh = findSynthesisSheetIn(wrk1)
    
    If Not sh Is Nothing Then
    
        If greenLightMode Then
            
            If checkIfGreenLightLabelsAreInStd(sh) Then
                ' --------------------------------------------
                
                Dim r As Range
                Set r = findLastEmptyColumn(sh)
                
                
                
                If Not r Is Nothing Then
                    Set r = insertNewColumn(r)
                    putDataFromSelectedCW r
                End If
                
                ' --------------------------------------------
            End If
            
        End If
        

    End If
End Function

Private Function makeReceptionColumnAtTheEnd() As Range
    Set makeReceptionColumnAtTheEnd = Nothing
    
    Dim sh As Worksheet
    Set sh = findSynthesisSheetIn(wrk1)
    
    
    If Not sh Is Nothing Then
    
        If receptionMode Then
        
            If checkIfReceptionLabelsAreInStd(sh) Then
                ' --------------------------------------------
                
                Dim r As Range
                Set r = findLastEmptyColumn(sh)
                
                
                
                If Not r Is Nothing Then
                    Debug.Print r.Address & " " & r.Parent.name
                    Set r = insertNewColumn(r)
                    putDataFromSelectedCW r
                End If
                ' --------------------------------------------
            End If
        End If
    End If
End Function




' inserting data to exising report
' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------

Private Function insertNewColumn(r As Range) As Range
    Set insertNewColumn = Nothing
    
    
    Dim tmpPrev As Range
    Set tmpPrev = r
    
    r.EntireColumn.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Set insertNewColumn = tmpPrev.offset(0, -1)
End Function


Private Sub putDataFromSelectedCW(r As Range)


    currentColumnToFill = r.Column
    fillWithDataFromForm r.Parent
    
    addGridlines r
    
End Sub

' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------



Private Function findLastEmptyColumn(sh As Worksheet) As Range
    Set findLastEmptyColumn = Nothing
    
    Dim rf As Range
    Dim ir As Range
    Dim tmpTxt As String
    Set rf = sh.Range("XFD2").End(xlToLeft)
    
    If UCase(rf.Value) = "TOTAL" Then
        ' ok - really good
        
        Set rf = rf.offset(0, -1)
        
        If Trim(rf.Value) = "" Then
            

            tmpTxt = ""
            For Each ir In sh.Range(rf, rf.offset(10, 0))
                
                tmpTxt = tmpTxt & Trim(ir.Value)
            Next ir
            
            If tmpTxt = "" Then
                Set findLastEmptyColumn = rf
            Else
                Set findLastEmptyColumn = Nothing
            End If
        End If
    ElseIf Trim(rf.Value) = "" Then
    
    
        tmpTxt = ""
        For Each ir In sh.Range(rf, rf.offset(10, 0))
            
            tmpTxt = tmpTxt & Trim(ir.Value)
        Next ir
        
        If tmpTxt = "" Then
            Set findLastEmptyColumn = rf
        Else
            Set findLastEmptyColumn = Nothing
        End If
        
    Else
        Set findLastEmptyColumn = Nothing
    End If
End Function


Private Function findSynthesisSheetIn(w As Workbook) As Worksheet
    Set findSynthesisSheetIn = Nothing
    
    
    Dim sh As Worksheet, targetSh As Worksheet
    
    Set targetSh = Nothing
    
    For Each sh In w.Sheets
    
        'Private greenLightMode As Boolean
        'Private receptionMode As Boolean
        If receptionMode And checkIfReceptionLabelsAreInStd(sh) Then
            Set targetSh = sh
            Exit For
        End If
        
        If greenLightMode And checkIfReceptionLabelsAreInStd(sh) Then
            Set targetSh = sh
            Exit For
        End If
        
    Next sh
    
    Set findSynthesisSheetIn = targetSh
    
    
End Function


Private Function makeGreenLightScaffold() As Range
    Set makeGreenLightScaffold = Nothing
    
    Dim sh As Worksheet
    Set sh = justMakeWorkbookAndSheetInIt("GREEN_LIGHT")
    Set sh1 = sh
    
    
    ' fill green light labels
    fillGreenLightLabels sh
    fillWithDataFromForm sh
End Function


Private Function makeReceptionScaffold() As Range
    Set makeReceptionScaffold = Nothing
    
    Dim sh As Worksheet
    Set sh = justMakeWorkbookAndSheetInIt("RECEPTION")
    Set sh1 = sh
    
    
    ' fill reception labels
    fillReceptionLabels sh
    fillWithDataFromForm sh
End Function

Private Function justMakeWorkbookAndSheetInIt(shNm As String) As Worksheet

    Set justMakeWorkbookAndSheetInIt = Nothing
    
    Dim wrk As Workbook
    Dim sh As Worksheet
    
    If frm1.outWrk Is Nothing Then
        Set wrk = Application.Workbooks.Add
        Set sh = wrk.Sheets.Add
        sh.name = CStr(shNm)
        Set frm1.outWrk = wrk
        currentColumnToFill = 3
    Else
        Set wrk = frm1.outWrk
        
        Set sh = Nothing
        On Error Resume Next
        Set sh = wrk.Sheets(shNm)
        
        If sh Is Nothing Then
            Set wrk = Application.Workbooks.Add
            Set sh = wrk.Sheets.Add
            sh.name = CStr(shNm)
            
            currentColumnToFill = 3
        Else
            currentColumnToFill = sh.Cells(3, 2).End(xlToRight).Column + 1
        End If
    End If
    

    
    Set justMakeWorkbookAndSheetInIt = sh
    
    
End Function


Private Sub fillGreenLightLabels(sh1 As Worksheet)


    'E_MODELESS_LEAF_ORDER_DATE = 2
    'E_MODELESS_LEAF_ORDER_MASS_EUR_INTERNAL = 3
    'E_MODELESS_LEAF_ORDER_NUM_OF_INTERNAL = 4
    'E_MODELESS_LEAF_ORDER_MASS_EUR_WO_TANGO = 5
    'E_MODELESS_LEAF_ORDER_PERC_MASS_EUR_WO_TANGO = 6
    'E_MODELESS_LEAF_ORDER_NUM_OF_WO_TANGO = 7
    'E_MODELESS_LEAF_ORDER_RATE_WO_TANGO = 8
    'E_MODELESS_LEAF_ORDER_EMPTY_1
    'E_MODELESS_LEAF_ORDER_MASS_EUR_W_TANGO = 10
    'E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO
    'E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO_NOK = 12
    'E_MODELESS_LEAF_ORDER_EMPTY_2
    'E_MODELESS_LEAF_ORDER_MASS_EUR_TARGET_PRICE = 14
    'E_MODELESS_LEAF_ORDER_RATE
    'E_MODELESS_LEAF_ORDER_MASS_EUR_GAP

    With sh1
        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_INTERNAL, 2).Value = "Mass € Internal part (out of perimeter price * 1)"
        .Cells(E_MODELESS_LEAF_ORDER_NUM_OF_INTERNAL, 2).Value = "Number of Internal parts"
        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_WO_TANGO, 2).Value = "Mass € Forecast without Tango price"
        .Cells(E_MODELESS_LEAF_ORDER_PERC_MASS_EUR_WO_TANGO, 2).Value = "% mass €  without Tango price"
        .Cells(E_MODELESS_LEAF_ORDER_NUM_OF_WO_TANGO, 2).Value = "number ref without Tango price"
        
        .Cells(E_MODELESS_LEAF_ORDER_RATE_WO_TANGO, 2).Value = "rate  : Mass Sigapp pre-serial price / initial price (without Tango price)"
        
        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_W_TANGO, 2).Value = "Mass € Sigapp forecast (With Tango price)"
        .Cells(E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO, 2).Value = "number ref with Tango price"
        .Cells(E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO_NOK, 2).Value = "Including number no price * 1"

        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_TARGET_PRICE, 2).Value = "Mass € target Price * 1"
        .Cells(E_MODELESS_LEAF_ORDER_RATE, 2).Value = "rate  : Mass forecast Sigapp price /target price * 1 (with Tango price)"
        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_GAP, 2).Value = "Mass € gap Sigapp price / price *1"
        
        .Cells(EVO.E_MODELESS_LEAF_ORDER_RATE_FINAL, 2).Value = "rate  : Mass reception Sigapp price /TARGET price * 1"
        .Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_GAP_FINAL, 2).Value = "Mass € gap Sigapp price / price *1"
    End With
    
    sh1.Range("B1:B1").columns.AutoFit

End Sub

Public Function checkIfGreenLightLabelsAreInStd(sh1 As Worksheet) As Boolean
    checkIfGreenLightLabelsAreInStd = False
    
    
    With sh1
        If .Cells(3, 2).Value = "Mass € Internal part (out of perimeter price * 1)" Then
            If .Cells(4, 2).Value = "Number of Internal parts" Then
                If .Cells(5, 2).Value = "Mass € Forecast without Tango price" Then
                    If .Cells(6, 2).Value = "% mass €  without Tango price" Then
                    
                        checkIfGreenLightLabelsAreInStd = True
                    End If
                End If
            End If
        End If
    End With
    
End Function


Private Sub fillReceptionLabels(sh1 As Worksheet)


    'E_MODELESS_LEAF_ORDER_DATE = 2
    'E_MODELESS_LEAF_ORDER_MASS_EUR_INTERNAL = 3
    'E_MODELESS_LEAF_ORDER_NUM_OF_INTERNAL = 4
    'E_MODELESS_LEAF_ORDER_MASS_EUR_WO_TANGO = 5
    'E_MODELESS_LEAF_ORDER_PERC_MASS_EUR_WO_TANGO = 6
    'E_MODELESS_LEAF_ORDER_NUM_OF_WO_TANGO = 7
    'E_MODELESS_LEAF_ORDER_RATE_WO_TANGO = 8
    'E_MODELESS_LEAF_ORDER_EMPTY_1
    'E_MODELESS_LEAF_ORDER_MASS_EUR_W_TANGO = 10
    'E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO
    'E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO_NOK = 12
    'E_MODELESS_LEAF_ORDER_EMPTY_2
    'E_MODELESS_LEAF_ORDER_MASS_EUR_TARGET_PRICE = 14
    'E_MODELESS_LEAF_ORDER_RATE
    'E_MODELESS_LEAF_ORDER_MASS_EUR_GAP

    With sh1
        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_INTERNAL, 2).Value = "Mass € Internal part (out of perimeter price * 1)"
        .Cells(E_MODELESS_LEAF_ORDER_NUM_OF_INTERNAL, 2).Value = "Number of Internal parts"
        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_WO_TANGO, 2).Value = "Mass € Reception without Tango price"
        .Cells(E_MODELESS_LEAF_ORDER_PERC_MASS_EUR_WO_TANGO, 2).Value = "% mass €  without Tango price"
        .Cells(E_MODELESS_LEAF_ORDER_NUM_OF_WO_TANGO, 2).Value = "number ref without Tango price"
        
        .Cells(E_MODELESS_LEAF_ORDER_RATE_WO_TANGO, 2).Value = "rate  : Mass Sigapp pre-serial price / initial price (without Tango price)"
        
        .Cells(E_MODELESS_LEAF_ORDER_EMPTY_1, 2).Value = "Mass € target tp04 Price * 1"
        
        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_W_TANGO, 2).Value = "Mass € Sigapp reception (With Tango price)"
        .Cells(E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO, 2).Value = "number ref with Tango price"
        .Cells(E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO_NOK, 2).Value = "Including number no price * 1"

        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_TARGET_PRICE, 2).Value = "Mass € target Price * 1"
        .Cells(E_MODELESS_LEAF_ORDER_RATE, 2).Value = "rate  : Mass reception Sigapp price /target price * 1 (with Tango price)"
        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_GAP, 2).Value = "Mass € gap Sigapp price / price *1"
        
        .Cells(EVO.E_MODELESS_LEAF_ORDER_RATE_FINAL, 2).Value = "rate  : Mass reception Sigapp price /TARGET price * 1"
        .Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_GAP_FINAL, 2).Value = "Mass € gap Sigapp price / price *1"
    End With
    
    sh1.Range("B1:B1").columns.AutoFit

End Sub


Public Function checkIfReceptionLabelsAreInStd(sh1 As Worksheet) As Boolean
    checkIfReceptionLabelsAreInStd = False
    
    
    With sh1
        If .Cells(3, 2).Value = "Mass € Internal part (out of perimeter price * 1)" Then
            If .Cells(4, 2).Value = "Number of Internal parts" Then
                If .Cells(5, 2).Value = "Mass € Reception without Tango price" Then
                    If .Cells(6, 2).Value = "% mass €  without Tango price" Then
                    
                        checkIfReceptionLabelsAreInStd = True
                    End If
                End If
            End If
        End If
    End With
End Function


Private Sub fillWithDataFromForm(sh1 As Worksheet)

    'E_MODELESS_LEAF_ORDER_DATE = 2
    'E_MODELESS_LEAF_ORDER_MASS_EUR_INTERNAL = 3
    'E_MODELESS_LEAF_ORDER_NUM_OF_INTERNAL = 4
    'E_MODELESS_LEAF_ORDER_MASS_EUR_WO_TANGO = 5
    'E_MODELESS_LEAF_ORDER_PERC_MASS_EUR_WO_TANGO = 6
    'E_MODELESS_LEAF_ORDER_NUM_OF_WO_TANGO
    'E_MODELESS_LEAF_ORDER_RATE_WO_TANGO
    'E_MODELESS_LEAF_ORDER_EMPTY_1
    'E_MODELESS_LEAF_ORDER_MASS_EUR_W_TANGO
    'E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO
    'E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO_NOK = 12
    'E_MODELESS_LEAF_ORDER_EMPTY_2
    'E_MODELESS_LEAF_ORDER_MASS_EUR_TARGET_PRICE
    'E_MODELESS_LEAF_ORDER_RATE
    'E_MODELESS_LEAF_ORDER_MASS_EUR_GAP
    
    ' new to final rate and final gap incl. data with articles wo tango prices
    
    sh1.Parent.Activate
    sh1.Activate
    
    Dim r As Range
    With sh1
        .Cells(EVO.E_MODELESS_LEAF_ORDER_DATE, currentColumnToFill).Value = _
            frm1.ListBoxScope.Value
        .Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_INTERNAL, currentColumnToFill).Value = _
            CDbl(parseTextNum(frm1.TextBox_CostInternal.Value))
        .Cells(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_INTERNAL, currentColumnToFill).Value = _
            frm1.TextBox_CountInternal.Value
        .Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_WO_TANGO, currentColumnToFill).Value = _
            CDbl(parseTextNum(frm1.TextBox_CostNoTango.Value))
        .Cells(EVO.E_MODELESS_LEAF_ORDER_PERC_MASS_EUR_WO_TANGO, currentColumnToFill).FormulaR1C1Local = _
            "=R[-1]C/(R[-1]C + R[4]C)"
        .Cells(EVO.E_MODELESS_LEAF_ORDER_NUM_OF_WO_TANGO, currentColumnToFill).Value = _
            frm1.TextBox_CountNoTango.Value
        
        ' "Mass € target tp04 Price * 1"
        If IsNumeric(frm1.TextBox_TargetNoTango.Value) Then
            .Cells(E_MODELESS_LEAF_ORDER_EMPTY_1, currentColumnToFill).Value = CDbl(parseTextNum(frm1.TextBox_TargetNoTango.Value))
        End If
        
        ' row 8 - rate for parts without tango GL
        .Cells(E_MODELESS_LEAF_ORDER_RATE_WO_TANGO, currentColumnToFill).Value = _
            CDbl(parseTextNum(frm1.TextBox_RateNoTango.Value))
        
        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_W_TANGO, currentColumnToFill).Value = _
            CDbl(parseTextNum(frm1.TextBox_CostTango.Value))
            
        On Error Resume Next
        .Cells(E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO, currentColumnToFill).Value = _
            CLng(frm1.TextBox_CountTango.Value) ' + CLng(frm1.TextBox_CountTangoNOK.Value)
        .Cells(E_MODELESS_LEAF_ORDER_NUM_OF_W_TANGO_NOK, currentColumnToFill).Value = _
            frm1.TextBox_CountTangoNOK.Value
        
        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_TARGET_PRICE, currentColumnToFill).Value = _
            CDbl(parseTextNum(frm1.TextBox_CostTarget.Value))
            
        .Cells(E_MODELESS_LEAF_ORDER_RATE, currentColumnToFill).FormulaR1C1Local = _
            "=R[-5]C/R[-1]C"
        .Cells(E_MODELESS_LEAF_ORDER_MASS_EUR_GAP, currentColumnToFill).FormulaR1C1Local = _
            "=R[-6]C - R[-2]C"
            
        ' NEW FINAL RATE and GAP incl. also now data from articles wo tango prices!
        ' row for: rate  : Mass reception Sigapp price /TARGET price * 1
        .Cells(EVO.E_MODELESS_LEAF_ORDER_RATE_FINAL, currentColumnToFill).FormulaR1C1Local = _
            "=(R[-8]C+R[-13]C)/(R[-4]C+R[-9]C)"
        ' row for: REAL Mass € gap Sigapp price / price *1
        .Cells(EVO.E_MODELESS_LEAF_ORDER_MASS_EUR_GAP_FINAL, currentColumnToFill).FormulaR1C1Local = _
            "=R[-3]C+(R[-14]C-R[-10]C)"
    End With
End Sub


Private Function parseTextNum(txt As String) As Double
    
    parseTextNum = 0#
    
    On Error Resume Next
    parseTextNum = CDbl(txt)
    
    If parseTextNum = 0 Then
        
        If IsNumeric(txt) Then
            Debug.Print "is numeric, but it is not working!"
            
            
        Else
            
            Debug.Print "is not numeric!"
            
            If CStr(txt) Like "*.???,??" Then
                
                txt = Replace(txt, ".", "")
                On Error Resume Next
                parseTextNum = CDbl(txt)
               
            ElseIf CStr(txt) Like "*.*" Then
                
                Dim tmpArr As Variant
                tmpArr = Split(txt, ".")
                parseTextNum = CDbl(tmpArr(0) & Application.DecimalSeparator & Left(tmpArr(1), 2))
            End If
        End If
    End If
End Function


Private Sub addGridlines(r As Range)


    Dim whatToGrid As Range
    Set whatToGrid = r.Parent.Range(r, r.offset(17, 0))

    'Range("F2:F16").Select
    whatToGrid.Borders(xlDiagonalDown).LineStyle = xlNone
    whatToGrid.Borders(xlDiagonalUp).LineStyle = xlNone
    With whatToGrid.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With whatToGrid.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With whatToGrid.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With whatToGrid.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With whatToGrid.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With whatToGrid.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub
