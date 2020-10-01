VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TP04"
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


Private stHndlr As StatusHandler

Private pusMasterSheet As Worksheet
Private tp04Sheet As Worksheet

Private pt04Item As TP04ItemHandler
Private pt04ItemDictionary As Dictionary


Private rngOfArticlesInsideTp04 As Range


Public Sub setStatusHandler(sh1 As StatusHandler)
    Set stHndlr = sh1
End Sub



Public Sub init(pusSh As Worksheet, tp04Sh As Worksheet, Optional byUnit As Boolean)


    If (Not pusSh Is Nothing) And (Not tp04Sh Is Nothing) Then

        Set pusMasterSheet = pusSh
        Set tp04Sheet = tp04Sh
        
        
        Set rngOfArticlesInsideTp04 = defineRangeOfArticlesFromPT04()
        
        
        If Not rngOfArticlesInsideTp04 Is Nothing Then
        
            Set pt04ItemDictionary = New Dictionary
            
            Dim irek As Variant
            
            
            Dim r As Range
            Dim wynikSzukania1 As Range, adresPiewrszegoSzukania1 As String
            Dim strArticleWithoutIndice As String
            Dim wynikSzukania2 As Range, adresPiewrszegoSzukania2 As String
            
            Dim bMatched1 As Boolean, bMatched2 As Boolean
            
            
            Set r = pusMasterSheet.Range("A3")
            
            stHndlr.init_statusbar ((r.End(xlDown).row) * 2) / 10#
            stHndlr.show
            DoEvents
            
            Do
            
            
                bMatched1 = False
                bMatched2 = False
            
                If Not r.EntireRow.Hidden Then
                    ' make logic only for visible data
                    ' --------------------------------------------------
                    
                    
                    ' we are working here from PUS perspective data
                    ' so 2 sheets
                    '
                    ' searching 1 pl : wynikSzukania1: with index
                    ' searching 2 pl : wynikSzukania2: base article without index
                    
                    ' Debug.Print r.row
                    
                    ' check if dictionary have this data already
                    If Not pt04ItemDictionary.Exists(CStr(r.Value & "_" & r.Offset(0, EVO.E_MASTER_cofor - 1).Value)) Then
                    
                        ' multi step matching:
                        ' 1.  "this" article - with proper inidice and matched cofor
                        Set wynikSzukania1 = Nothing
                        
                        On Error Resume Next
                        Set wynikSzukania1 = rngOfArticlesInsideTp04.Find(CStr(r.Value))
                        
                        If Not wynikSzukania1 Is Nothing Then
                            adresPiewrszegoSzukania1 = wynikSzukania1.Address
                            
    
                            
                            Do
                            
                                'Debug.Print Cstr(r.Offset(0, EVO.E_MASTER_cofor - 1).Value) & "_" & _
                                '    Cstr(wynikSzukania1.Offset(0, EVO.E_TP04_01_FOUR - EVO.E_TP04_01_ARTICLE))
                            
                                If CStr(r.Offset(0, EVO.E_MASTER_cofor - 1).Value) = _
                                    CStr(wynikSzukania1.Offset(0, EVO.E_TP04_01_FOUR - EVO.E_TP04_01_ARTICLE)) Then
                                    
                                    
                                    ' matched! so OK!
                                    ' -------------------------------------------------------------------------------
                                    bMatched1 = True
                                    Exit Do
                                    ' -------------------------------------------------------------------------------
                                    
                                    
                                Else
                                    
                                    bMatched1 = False
                                    Set wynikSzukania1 = rngOfArticlesInsideTp04.FindNext(wynikSzukania1)
                                    
                                    
                                End If
                            Loop Until adresPiewrszegoSzukania1 = wynikSzukania1.Address
                        End If
                        
                        ' 2. "base" article - with also matched cofor
                        Set wynikSzukania2 = Nothing
                        
                        If CStr(CStr(r.Value)) Like "*-??" Then
                            strArticleWithoutIndice = Left(r.Value, Len(r.Value) - 3)
                        Else
                            strArticleWithoutIndice = CStr(r.Value)
                        End If
                        
                        On Error Resume Next
                        Set wynikSzukania2 = rngOfArticlesInsideTp04.Find(strArticleWithoutIndice) ' , LookIn:=xlValues, LookAt:=xlWhole)
                        
                        If wynikSzukania2 Is Nothing Then
                            DoEvents
                            On Error Resume Next
                            Set wynikSzukania2 = rngOfArticlesInsideTp04.Find(strArticleWithoutIndice) ' , LookIn:=xlValues, LookAt:=xlWhole)
                        End If
                        
                        If wynikSzukania2 Is Nothing Then
                            DoEvents
                            ' dalej cos nie dziala....
                            
                            For Each irek In rngOfArticlesInsideTp04
                            
                                If CStr(strArticleWithoutIndice) = CStr(irek.Value) Then
                                    Set wynikSzukania2 = irek
                                    Exit For
                                End If
                                
                            Next irek
                        End If
                        
                        If Not wynikSzukania2 Is Nothing Then
                            adresPiewrszegoSzukania2 = wynikSzukania2.Address
                            
    
                            
                            Do
                            
                                'Debug.Print Cstr(r.Offset(0, EVO.E_MASTER_cofor - 1).Value) & "_" & _
                                '    Cstr(wynikSzukania2.Offset(0, EVO.E_TP04_01_FOUR - EVO.E_TP04_01_ARTICLE))
                                
                                ' Debug.Print Len(wynikSzukania2)
                            
                                If CStr(r.Offset(0, EVO.E_MASTER_cofor - 1).Value) = _
                                    CStr(wynikSzukania2.Offset(0, EVO.E_TP04_01_FOUR - EVO.E_TP04_01_ARTICLE)) And _
                                    Int(Len(wynikSzukania2)) = Int(10) Then
                                    
                                    
                                    ' matched! so OK!
                                    ' -------------------------------------------------------------------------------
                                    bMatched2 = True
                                    Exit Do
                                    ' -------------------------------------------------------------------------------
                                    
                                    
                                Else
                                    
                                    bMatched2 = False
                                    Set wynikSzukania2 = rngOfArticlesInsideTp04.FindNext(wynikSzukania2)
                                    
                                    
                                    If wynikSzukania2 Is Nothing Then
                                        ' musimy szukac jeszcze raz!
                                        Debug.Print "wynikSzukania2 is nothing :("
                                    End If
                                    
                                    
                                End If
                            Loop Until adresPiewrszegoSzukania2 = wynikSzukania2.Address
                        End If
                        
                        ' article + cofor_ven
                        If Not pt04ItemDictionary.Exists(CStr(r.Value & "_" & r.Offset(0, EVO.E_MASTER_cofor - 1).Value)) Then
                        
                        
                            ' nop
                            'If bMatched Then
                            '
                            'End If
                        
                            
                        
                            ' not exisits
                            Set pt04Item = Nothing
                            Set pt04Item = New TP04ItemHandler
                            With pt04Item
                            
                                .name = CStr(r.Value & "_" & r.Offset(0, EVO.E_MASTER_cofor - 1).Value)
                                
                                If bMatched1 Then
                                    .thisLine.article = r.Value
                                    
                                    ' Debug.Assert Trim(r.Value) <> "9677602580-03"
                                    
                                    
                                    ' new from 060
                                    .thisLine.domain = _
                                        CStr(wynikSzukania1.Offset(0, _
                                            EVO.E_FROM_SQ01_QUASI_TP04_DOMAIN - EVO.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value)
                                            
                                    .thisLine.ru = _
                                        CStr(wynikSzukania1.Offset(0, _
                                            EVO.E_FROM_SQ01_QUASI_TP04_RU - EVO.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value)
                                            
                                    .thisLine.div = _
                                        CStr(wynikSzukania1.Offset(0, _
                                            EVO.E_FROM_SQ01_QUASI_TP04_DIV - EVO.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value)
                                    
                                    
                                    .thisLine.cofor = r.Offset(0, EVO.E_MASTER_cofor - 1).Value
                                    
                                    .thisLine.price = _
                                        parsePriceAgain(wynikSzukania1.Offset(0, _
                                            EVO.E_FROM_SQ01_QUASI_TP04_SUM - EVO.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value)
                                            
                                    .thisLine.curr = _
                                        wynikSzukania1.Offset(0, _
                                            EVO.E_FROM_SQ01_QUASI_TP04_CURRENCY - EVO.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value
                                    
                                    .thisLine.rate = 1
                                    On Error Resume Next
                                    .thisLine.rate = replaceDotWithDecimalPointer(Application.WorksheetFunction.VLookup(UCase(.thisLine.curr), _
                                        ThisWorkbook.Sheets("register").Range("B98:C500"), 2, False))
                                        
                                    .thisLine.priceInEur = (1# / CDbl(.thisLine.rate)) * 1# * .thisLine.price
                                    
                                    
                                    .thisLine.pckgUn = wynikSzukania1.Offset(0, EVO.E_TP04_01_UNITE_DE_PRIX - EVO.E_TP04_01_ARTICLE).Value
                                    .thisLine.pckgDiv = calcUn(.thisLine.pckgUn)
                                    
                                    
                                    .thisLine.finalPrice = .thisLine.priceInEur / .thisLine.pckgDiv
                                    
                                End If
                                
                                
                                If bMatched2 Then
                                    .baseLine.article = wynikSzukania2.Value
                                    .baseLine.cofor = r.Offset(0, EVO.E_MASTER_cofor - 1).Value
                                    .baseLine.price = parsePriceAgain(wynikSzukania2.Offset(0, EVO.E_TP04_01_PRIX - EVO.E_TP04_01_ARTICLE).Value)
                                    .baseLine.curr = wynikSzukania2.Offset(0, EVO.E_TP04_01_DEV - EVO.E_TP04_01_ARTICLE).Value
                                    
                                    .baseLine.rate = 1
                                    On Error Resume Next
                                    .baseLine.rate = replaceDotWithDecimalPointer(Application.WorksheetFunction.VLookup(UCase(.baseLine.curr), _
                                        ThisWorkbook.Sheets("register").Range("B98:C500"), 2, False))
                                        
                                    .baseLine.priceInEur = (1# / CDbl(.baseLine.rate)) * 1# * .baseLine.price
                                    
                                    
                                    .baseLine.pckgUn = wynikSzukania2.Offset(0, EVO.E_TP04_01_UNITE_DE_PRIX - EVO.E_TP04_01_ARTICLE).Value
                                    .baseLine.pckgDiv = calcUn(.baseLine.pckgUn)
                                    
                                    .baseLine.finalPrice = .baseLine.priceInEur / .baseLine.pckgDiv
                                    
                                    
                                    
                                End If
                                
                            End With
                            
                            ' Debug.Assert pt04Item.baseLine.rate = 1
                            
                            pt04ItemDictionary.Add CStr(r.Value & "_" & r.Offset(0, EVO.E_MASTER_cofor - 1).Value), pt04Item
                        End If
                        
                        
                        ' --------------------------------------------------
                    
                    End If
                End If
                Set r = r.Offset(1, 0)
                
                
                If r.row Mod 10 = 0 Then stHndlr.progress_increase
            Loop Until CStr(r) = ""
            
            
            
            
            Set r = pusMasterSheet.Range("A3")
            

            DoEvents
            
            If pt04ItemDictionary.Count > 0 Then
            
                Dim repsh As Worksheet, repRng As Range, outputIter As Variant
            
                Set repsh = ThisWorkbook.Sheets.Add
                ' tryToRename repSh
                repsh.name = EVO.TryToRenameModule.tryToRenameWorksheet(repsh, "TP04_")
                addLabelsToPT04Report repsh, pusMasterSheet
                Set repRng = repsh.Cells(2, 1)
            
                Do
                    If Not r.EntireRow.Hidden Then
                    
                        Set pt04Item = Nothing
                    
                        On Error Resume Next
                        Set pt04Item = pt04ItemDictionary(CStr(r.Value & "_" & r.Offset(0, EVO.E_MASTER_cofor - 1).Value))
                        
                        
                        If Not pt04Item Is Nothing Then
                        
                            If r.Offset(0, EVO.E_MASTER_Reference - 1).Value Like "*-??" Then
                            
                            
                                ' Debug.Assert Trim(r.Value) <> "9677602580-03"
                                
                                ' new 060
                                ' two ENUMS!!!
                                ' E_TP04_OUTPUT_REP_ - just part after blue columns from PUS!
                                ' evo.E_ADJUSTED_SQ01 - all the table!
                                repRng.Offset(0, EVO.E_ADJUSTED_SQ01_DOMAIN - 1).Value = _
                                    pt04Item.thisLine.domain
                                    
                                repRng.Offset(0, EVO.E_ADJUSTED_SQ01_DIV - 1).Value = _
                                    pt04Item.thisLine.div
                                    
                                repRng.Offset(0, EVO.E_ADJUSTED_SQ01_RU - 1).Value = _
                                    pt04Item.thisLine.ru
                                
                            
                            
                                For outputIter = EVO.E_MASTER_Reference To E_MASTER_TMC
                                    repRng.Offset(0, outputIter - 1).Value = r.Offset(0, outputIter - 1).Value
                                Next outputIter
                                
                                'For outputIter = E_TP04_OUTPUT_REP_PART_PRESERIAL To EVO.E_TP04_OUTPUT_REP_PART_OKNOK
                                'Next outputIter
                                
                                If byUnit Then
                                    ' tbd
                                    
                                    repRng.Offset(0, E_TP04_OUTPUT_REP_PART_PRESERIAL - 1).Value = _
                                        pt04Item.thisLine.priceInEur
                                    repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_INIT_SERIAL_PRICE - 1).Value = _
                                        pt04Item.baseLine.priceInEur
                                        
                                    repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_RATE - 1).Value = _
                                        CDbl(pt04Item.thisLine.priceInEur / pt04Item.baseLine.priceInEur)
                                    
                                    
                                    If CDbl(pt04Item.baseLine.priceInEur) = 0# And CDbl(pt04Item.thisLine.priceInEur) = 0# Then
                                        repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_OKNOK - 1).Value = "NOK"
                                    Else
                                        repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_OKNOK - 1).Value = defineOKNOKon(repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_RATE - 1))
                                    End If
        
                                    
                                    If repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_OKNOK - 1).Value = "NOK" Then
                                        repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_OKNOK - 1).Interior.Color = RGB(245, 0, 0)
                                    End If
                                Else
                                
                                    ' byUnit = false is default!
                                    ' so taking final price after div on size of the packaging!
                                
                                    repRng.Offset(0, E_TP04_OUTPUT_REP_PART_PRESERIAL - 1).Value = _
                                        pt04Item.thisLine.finalPrice
                                    repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_INIT_SERIAL_PRICE - 1).Value = _
                                        pt04Item.baseLine.finalPrice
                                        
                                    repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_RATE - 1).Value = _
                                        CDbl(pt04Item.thisLine.finalPrice / pt04Item.baseLine.finalPrice)
                                    
                                    
                                    If CDbl(pt04Item.baseLine.finalPrice) = 0# And CDbl(pt04Item.thisLine.finalPrice) = 0# Then
                                        repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_OKNOK - 1).Value = "NOK"
                                    Else
                                        repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_OKNOK - 1).Value = defineOKNOKon(repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_RATE - 1))
                                    End If
        
                                    
                                    If repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_OKNOK - 1).Value = "NOK" Then
                                        repRng.Offset(0, EVO.E_TP04_OUTPUT_REP_PART_OKNOK - 1).Interior.Color = RGB(245, 0, 0)
                                    End If
                                End If
                                    
                                Set repRng = repRng.Offset(1, 0)
                                
                            Else
                            
                                ' Debug.Print r.Offset(0, EVO.E_MASTER_Reference - 1).Value & " " & r.Offset(0, EVO.E_MASTER_Reference - 1).row
                                
                            End If
                            
                            
                        Else
                            
                            'nop
                            
                        End If
                        
                        
                        
                    End If
                    
                    Set r = r.Offset(1, 0)
                    
                    If r.row Mod 10 = 0 Then
                        stHndlr.progress_increase
                    End If
                Loop Until CStr(r) = ""
            End If
            
            
            
        Else
            MsgBox "the def of rngOfArticlesInsideTp04 is empty!", vbCritical
        End If
        
        stHndlr.hide
        
        
        
        
    Else
        MsgBox "No data to match!", vbCritical
    End If
End Sub


Private Function parsePriceAgain(x As Variant) As Double
    
    If IsNumeric(x) Then
        'Debug.Print x
        
        parsePriceAgain = CDbl(x)
    Else
        'Debug.Print x & " " & Len(x)
        
        If x Like "*.???,??" Then
            
            'Debug.Print "yes it is!"
            x = Replace(x, ".", "")
            parsePriceAgain = CDbl(x)
        End If
        
    End If
End Function



Private Function calcUn(param As String) As Double
    calcUn = 1#
    
    Dim regRef As Range
    Set regRef = ThisWorkbook.Sheets("register").Range("UN_REF")
    
    Do
        If CStr(regRef.Value) = CStr(param) Then
            calcUn = CDbl(regRef.Offset(0, 1).Value)
            Exit Do
        End If
        Set regRef = regRef.Offset(1, 0)
    Loop Until Trim(regRef.Value) = ""
End Function



Private Function defineOKNOKon(r1 As Range) As String
    
    defineOKNOKon = "NOK"
    
    If IsNumeric(r1.Value) Then
        If CDbl(r1.Value) <= CDbl(ThisWorkbook.Sheets("register").Range("TRIGGER").Value) Then
            defineOKNOKon = "OK"
        End If
    End If
End Function


Private Sub tryToRename(ByRef sh1tp04 As Worksheet)
    
    With sh1tp04
        On Error Resume Next
        .name = .name & "_" & CStr(Date)
    End With
End Sub


Private Sub addLabelsToPT04Report(sh1tp04 As Worksheet, pusMasterSheet As Worksheet)

    Dim x As Variant
    With sh1tp04
    
    
        For x = E_MASTER_Reference To E_MASTER_TMC
            .Cells(1, x).Value = CStr(pusMasterSheet.Cells(2, x).Value)
            .Cells(1, x).Interior.Color = RGB(100, 100, 230)
            .Cells(1, x).Font.Bold = True
        Next x
        
        x = E_MASTER_TMC + 2
        .Cells(1, x).Value = "INITIAL SERIAL PRICE"
        .Cells(1, x).Interior.Color = RGB(240, 20, 20)
        .Cells(1, x).Font.Bold = True
        
        x = E_MASTER_TMC + 1
        .Cells(1, x).Value = "PRE-SERIAL PRICE (YPRS contract)"
        .Cells(1, x).Interior.Color = RGB(240, 20, 20)
        .Cells(1, x).Font.Bold = True
        
        x = E_MASTER_TMC + 3
        .Cells(1, x).Value = "RATE = PRE-SERIAL div INIT PRICE"
        .Cells(1, x).Interior.Color = RGB(240, 20, 20)
        .Cells(1, x).Font.Bold = True
        
        x = E_MASTER_TMC + 4
        .Cells(1, x).Value = "OKNOK"
        .Cells(1, x).Interior.Color = RGB(240, 20, 20)
        .Cells(1, x).Font.Bold = True
        
        
        ' x = E_MASTER_TMC + 4 ===     E_ADJUSTED_SQ01_OKNOK
        ' E_ADJUSTED_SQ01_WITH_INDEX
        ' E_ADJUSTED_SQ01_IS_INTERNAL
        ' E_ADJUSTED_SQ01_IS_IN_TANGO
        .Cells(1, E_ADJUSTED_SQ01_WITH_INDEX).Value = "WITH_INDEX"
        .Cells(1, E_ADJUSTED_SQ01_WITH_INDEX).Font.Bold = True
        .Cells(1, E_ADJUSTED_SQ01_IS_INTERNAL).Value = "IS_INTERNAL"
        .Cells(1, x).Font.Bold = True
        .Cells(1, E_ADJUSTED_SQ01_IS_IN_TANGO).Value = "IS_IN_TANGO"
        .Cells(1, E_ADJUSTED_SQ01_IS_IN_TANGO).Font.Bold = True
        .Cells(1, EVO.E_ADJUSTED_SQ01_TANGO_PCS_PRICE).Value = "TANGO_PCS_PRICE"
        .Cells(1, EVO.E_ADJUSTED_SQ01_TANGO_PCS_PRICE).Font.Bold = True
        .Cells(1, EVO.E_ADJUSTED_SQ01_TANGO_PCS_PRICE).EntireColumn.ColumnWidth = 10
        
        .Cells(1, EVO.E_ADJUSTED_SQ01_DIV).Value = "DIV"
        .Cells(1, EVO.E_ADJUSTED_SQ01_DIV).Font.Bold = True
        .Cells(1, EVO.E_ADJUSTED_SQ01_DIV).EntireColumn.ColumnWidth = 8
        
        .Cells(1, EVO.E_ADJUSTED_SQ01_DOMAIN).Value = "DOMAIN"
        .Cells(1, EVO.E_ADJUSTED_SQ01_DOMAIN).Font.Bold = True
        .Cells(1, EVO.E_ADJUSTED_SQ01_DOMAIN).EntireColumn.ColumnWidth = 9
        
        .Cells(1, EVO.E_ADJUSTED_SQ01_RU).Value = "RU"
        .Cells(1, EVO.E_ADJUSTED_SQ01_RU).Font.Bold = True
        .Cells(1, EVO.E_ADJUSTED_SQ01_RU).EntireColumn.ColumnWidth = 6
        
        
        .Cells(1, EVO.E_ADJUSTED_SQ01_MANAGER_DA).Value = "Managers DA"
        .Cells(1, EVO.E_ADJUSTED_SQ01_MANAGER_DA).Font.Bold = True
        .Cells(1, EVO.E_ADJUSTED_SQ01_MANAGER_DA).EntireColumn.ColumnWidth = 12
        
        
        ' auto width
        Dim awtr As Range
        Set awtr = .Range(.Cells(E_MASTER_TMC + 1, 1), .Cells(1, E_MASTER_TMC + 4)) ' :P
        awtr.EntireColumn.AutoFit
        
        Set awtr = .Range(.Cells(1, E_MASTER_Reference), .Cells(1, E_MASTER_TMC)) ' :P
        awtr.EntireColumn.ColumnWidth = 6
    End With
End Sub


Private Function defineRangeOfArticlesFromPT04() As Range
    
    Set defineRangeOfArticlesFromPT04 = Nothing
    
    Dim r As Range
    Set r = tp04Sheet.Cells(1, EVO.E_TP04_01_ARTICLE)
    
    If CStr(r.Offset(1, 0).Value) <> "" Then
    
        Set defineRangeOfArticlesFromPT04 = tp04Sheet.Range(r, r.Parent.Range("A1048576").End(xlUp).Offset(1, 0))
        
        
        Debug.Print defineRangeOfArticlesFromPT04.Address
    End If
End Function
