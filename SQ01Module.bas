Attribute VB_Name = "SQ01Module"
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




Public Sub concatDataFromSq01(ictrl As IRibbonControl)
    concatAndStd Nothing, Nothing, Nothing
End Sub

Public Sub concatAndStd(Optional sq01sh1 As Worksheet, Optional sq01sh2 As Worksheet, Optional resultSh As Worksheet, Optional sap4out As SAP_Handler)



    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False


    Dim whatToConcat As Collection
    Set whatToConcat = Nothing
    
    
    
    If sq01sh1 Is Nothing And sq01sh2 Is Nothing Then
    
        SQ01ConfigForm.ListBox1.Clear
        SQ01ConfigForm.ListBox1.MultiSelect = fmMultiSelectMulti
        
        Dim sh1 As Worksheet
        For Each sh1 In ThisWorkbook.Sheets
            If isInSq01OutputStd(sh1) Then
                SQ01ConfigForm.ListBox1.addItem sh1.name
            End If
        Next sh1
        
        SQ01ConfigForm.show
        
        If SQ01ConfigForm.coll Is Nothing Then
            MsgBox "You do not choose anything!", vbCritical
            End
        Else
            
            If SQ01ConfigForm.coll.count = 1 Then
                MsgBox "You choose only one worksheet - so nothing to do - nothing to concatenate", vbInformation
            Else
                Set whatToConcat = SQ01ConfigForm.coll
            End If
        End If
    ElseIf (Not sq01sh1 Is Nothing) And (Not sq01sh2 Is Nothing) Then
    
    
        Set whatToConcat = New Collection
        whatToConcat.Add sq01sh1.name
        whatToConcat.Add sq01sh2.name
        
    Else
        MsgBox "This logic is not allowed!", vbCritical
    End If
    
    
    
    ' and of form section
    ' whatToConcat is collection check if form give some data
    
    If Not whatToConcat Is Nothing Then
    
        ' test OK
        'Dim el As Variant
        'For Each el In whatToConcat
        '    Debug.Print CStr(el)
        'Next el
        
        Dim concatSh As Worksheet
        Set concatSh = ThisWorkbook.Sheets.Add
        Set resultSh = concatSh
        concatSh.name = EVO.TryToRenameModule.tryToRenameWorksheet(concatSh, "CONCAT_")
        fillLabels concatSh.Range("A1")
        
        ' concat ref
        Dim cr As Range
        Set cr = concatSh.Range("A2")
        
        Dim leftRng As Range, rightRng As Range
        
        Dim el As Variant, tmpSrcSh As Worksheet, lastRow As Long, x1 As Variant
        For Each el In whatToConcat
            
            Set tmpSrcSh = ThisWorkbook.Sheets(CStr(el))
            
            lastRow = CLng(tmpSrcSh.Range("A1048576").End(xlUp).row)
            
            For x1 = 2 To lastRow
                
                ' EVO.E_FROM_SQ01_QUASI_TP04_DOMAIN , EVO.E_FROM_SQ01_QUASI_TP04_CURRENCY
                'concatSh.Range( _
                '    concatSh.Cells(cr.row, EVO.E_FROM_SQ01_QUASI_TP04_DOMAIN), concatSh.Cells(cr.row, EVO.E_FROM_SQ01_QUASI_TP04_CURRENCY)).Value = _
                'tmpSrcSh.Range( _
                '    tmpSrcSh.Cells(x1, EVO.E_FROM_SQ01_QUASI_TP04_DOMAIN), concatSh.Cells(x1, EVO.E_FROM_SQ01_QUASI_TP04_CURRENCY)).Value
                
                Set leftRng = concatSh.Range( _
                    concatSh.Cells(cr.row, EVO.E_FROM_SQ01_QUASI_TP04_DOMAIN), concatSh.Cells(cr.row, EVO.E_FROM_SQ01_QUASI_TP04_CURRENCY))
                Set rightRng = tmpSrcSh.Range( _
                    tmpSrcSh.Cells(x1, EVO.E_FROM_SQ01_QUASI_TP04_DOMAIN), tmpSrcSh.Cells(x1, EVO.E_FROM_SQ01_QUASI_TP04_CURRENCY))
                
                leftRng.Value = rightRng.Value
                
                    
                Set cr = cr.offset(1, 0)
                
                ' Application.Calculation = xlCalculationManual

            Next x1
            
            
            
        Next el
        
        
        ' now initial table in concat is ready - take some time and add some already converted info
        ' will help for reception rep ppx1
        goThroughListAgainAndTryToCalculPcsPriceinEur concatSh, sap4out
        
        ' MsgBox "READY!", vbInformation
    Else
        MsgBox "list of concat is empty!", vbCritical
        End
    End If
    
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = False
    
    
End Sub


Private Sub goThroughListAgainAndTryToCalculPcsPriceinEur(ByRef sh1 As Worksheet, Optional sap4out As SAP_Handler)


    '.thisLine.price = _
    '    parsePriceAgain(wynikSzukania1.Offset(0, _
    '        EVO.E_FROM_SQ01_QUASI_TP04_SUM - EVO.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value)
    '
    '.thisLine.curr = _
    '    wynikSzukania1.Offset(0, _
    '        EVO.E_FROM_SQ01_QUASI_TP04_CURRENCY - EVO.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value
    '
    '.thisLine.rate = 1
    'On Error Resume Next
    '.thisLine.rate = replaceDotWithDecimalPointer(Application.WorksheetFunction.VLookup(UCase(.thisLine.curr), _
    '    ThisWorkbook.Sheets("register").Range("B98:C500"), 2, False))
    '
    '.thisLine.priceInEur = (1# / CDbl(.thisLine.rate)) * 1# * .thisLine.price
    '
    '
    '.thisLine.pckgUn = wynikSzukania1.Offset(0, EVO.E_TP04_01_UNITE_DE_PRIX - EVO.E_TP04_01_ARTICLE).Value
    '.thisLine.pckgDiv = calcUn(.thisLine.pckgUn)
    '
    '
    '.thisLine.finalPrice = .thisLine.priceInEur / .thisLine.pckgDiv
    
    
    Dim tp04_instance As New TP04
    Dim ptim As TP04ItemPrimitive
    
    Dim r1 As Range
    Set r1 = sh1.Range("B2")
    
    Do
        Set ptim = Nothing
        Set ptim = New TP04ItemPrimitive
        ptim.article = CStr(r1)
        ptim.price = tp04_instance.parsePriceAgain(r1.offset(0, EVO.E_FROM_SQ01_QUASI_TP04_SUM - EVO.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value)
        ptim.curr = r1.offset(0, EVO.E_FROM_SQ01_QUASI_TP04_CURRENCY - EVO.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value
        ptim.rate = 1
        On Error Resume Next
        ptim.rate = replaceDotWithDecimalPointer(Application.WorksheetFunction.VLookup(UCase(ptim.curr), _
            ThisWorkbook.Sheets("register").Range("B98:C500"), 2, False))
        ptim.priceInEur = (1# / CDbl(ptim.rate)) * 1# * ptim.price
        ptim.pckgUn = r1.offset(0, EVO.E_TP04_01_UNITE_DE_PRIX - EVO.E_TP04_01_ARTICLE).Value
        ptim.pckgDiv = tp04_instance.calcUn(ptim.pckgUn)
        ptim.finalPrice = ptim.priceInEur / ptim.pckgDiv
        
        ' done
        ' now just put those data in table
        r1.offset(0, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_II_RATE_TO_EUR - _
            EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value = _
                ptim.rate
            
            
        r1.offset(0, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_II_RATE_UNIT_VALUE - _
            EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value = _
                ptim.pckgDiv
            
        r1.offset(0, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_II_RATE_CALCED_SUM - _
            EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value = _
                ptim.finalPrice
        
        Set r1 = r1.offset(1, 0)
        
        
        If (r1.row Mod 500) = 0 Then
            sap4out.justTouch
        End If
    Loop Until Trim(r1) = ""
    

End Sub


Private Function isInSq01OutputStd(refSh As Worksheet) As Boolean


    isInSq01OutputStd = True
    
    Dim lr As Range, x As Variant
    Set lr = refSh.Cells(1, 1)
    
    Dim refVal As Range
    Set refVal = ThisWorkbook.Sheets("forValidation").Range(EVO.G_REF_MOUNT_SQ1_OUT)
    
    For x = EVO.E_FROM_SQ01_QUASI_TP04_DOMAIN To EVO.E_FROM_SQ01_QUASI_TP04_CURRENCY
        If lr.offset(0, x - 1).Value = refVal.offset(0, x - 1).Value Then
        Else
            isInSq01OutputStd = False
            Exit For
        End If
            
    Next x
    
End Function

' ==============================================



Public Sub getSq01DataWithPreDefParams(ictrl As IRibbonControl)
    getDataFromSq01WithPreDefinedParams
End Sub


Public Sub getDataFromSq01WithPreDefinedParams()
    
    ' PRE_DEF_RUN_FOR_SQ01
    Dim refReg As Range, ctrl As TextBox, x As Variant
    Set refReg = ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01")
    
    Dim refStringForDomain As String
    refStringForDomain = ""
    
    With SQ01PreDefForm
        For x = 1 To 5
            On Error Resume Next
            .Controls("TextBox1" & CStr(x)).Value = CStr(refReg.offset(0, x - 1).Value)
            
            On Error Resume Next
            .Controls("TextBox2" & CStr(x)).Value = CStr(refReg.offset(1, x - 1).Value)
        Next x
        
        .show
    End With
    
    ' try to run with those predefs
    innerMainForSq01 True, SQ01PreDefForm.TextBox13.Value, SQ01PreDefForm.TextBox23.Value
End Sub

Public Sub getDataFromSq01(ictrl As IRibbonControl)
    
    ' Debug.Print "Welcome in SQ01 scope!"
    
    innerMainForSq01 False
End Sub





Public Sub innerMainForSq01(Optional preDef As Boolean, _
    Optional tbx13_Str As String, Optional tbx23_Str As String, _
    Optional ByRef osh As Worksheet, Optional ByRef osh2 As Worksheet, _
    Optional autoForCombo As Boolean)


    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    
    Dim numHandler As NumberHandler
    Set numHandler = New NumberHandler


    Dim st_h As StatusHandler, xStHelper As Integer
    Set st_h = New StatusHandler
    st_h.init_statusbar 20
    st_h.show
    
    
    delegacjaDlaProgresu st_h, xStHelper, 20
    
    
    ' inter4sh stands for internal suppliers list worksheet
    Dim ish As Worksheet, ish2 As Worksheet
    ' already as params to have possibility for combo logic
    ' Dim osh As Worksheet, osh2 As Worksheet
    ' dim inter4Sh As Worksheet,
    Dim irng As Range, orng As Range
    Set ish = ThisWorkbook.Sheets.Add
    Set osh = ThisWorkbook.Sheets.Add

    
    If tbx23_Str <> "" Then
        Set ish2 = ThisWorkbook.Sheets.Add
        Set osh2 = ThisWorkbook.Sheets.Add
    Else
        Set ish2 = Nothing
        Set osh2 = Nothing
    End If
    ' Set inter4Sh = ThisWorkbook.Sheets.Add
    ish.name = EVO.TryToRenameModule.tryToRenameWorksheet(ish, "IN1_" & CStr(tbx13_Str) & "_")
    osh.name = EVO.TryToRenameModule.tryToRenameWorksheet(osh, "OUT1_" & CStr(tbx13_Str) & "_")
    
    If Not ish2 Is Nothing And Not osh2 Is Nothing Then
        ish2.name = EVO.TryToRenameModule.tryToRenameWorksheet(ish2, "IN2_" & CStr(tbx23_Str) & "_")
        osh2.name = EVO.TryToRenameModule.tryToRenameWorksheet(osh2, "OUT2_" & CStr(tbx23_Str) & "_")
    End If
    
    
    ' inter4Sh.name = EVO.TryToRenameModule.tryToRenameWorksheet(inter4Sh, "N_" & CStr(tbx3_Str) & "_")
    
    
    ' LABLES ------------------------------------------------
    
    fillLabels osh.Range("A1")
    If Not osh2 Is Nothing Then fillLabels osh2.Range("A1")
    
    ' fillLabels inter4Sh.Range("A1")
    
    ' -------------------------------------------------------
    
    Dim sap__handler As New SAP_Handler
    If preDef Then
        sap__handler.runMainLogicForSQ01__with_preDef E_SQ01_CONFIG_MAKE_ALL, ish, osh, st_h, xStHelper, 1
        
        If Not ish2 Is Nothing And Not osh2 Is Nothing Then
            sap__handler.runMainLogicForSQ01__with_preDef E_SQ01_CONFIG_MAKE_ALL, ish2, osh2, st_h, xStHelper, 2
        End If
    Else
        sap__handler.runMainLogicForSQ01 E_SQ01_CONFIG_MAKE_ALL, ish, osh, st_h, xStHelper, 1
    End If
    
    ' automatisation on sigapp for internal suppliers - no matter what!
    ' sap__handler.runMainLogicFor__Y_PI1_80000391 inter4Sh, st_h, xStHelper
    
    
    
    
    ' COPY AND PASTE AS VALUES ------------------------------
    
    ' ???
    ' copyAndPasteAsValues osh.Range("A1").Offset(0, E_FROM_SQ01_QUASI_TP04_ARTICLE - 1)
    
    ' -------------------------------------------------------
    
    ' data ready - change string price into normal num
    changePricesIntoDouble osh
    changePricesIntoDouble osh2
    
    
    
    st_h.hide
    Set st_h = Nothing
    
    
    
    Set numHandler = Nothing
    
    
    
    
'    Dim answer As Variant
'    answer = MsgBox("Raw output from SQ01 (quasi TP04) ready! Do you want to continue?", vbYesNo + vbQuestion)
'
'    If answer = vbYes Then
'        runAdjusterForDataFromSq01 osh
'    End If
    
    
    If autoForCombo Then
    Else
        MsgBox "GOTOWE!", vbInformation
    End If
    
    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
End Sub





Public Sub changePricesIntoDouble(sh1 As Worksheet)


    Application.Calculation = xlCalculationManual
    
    Dim r As Range, tmpstrv As String, doubleValue As String
    Set r = sh1.Range("A1048576").End(xlUp) ' to jest ostatni
    
    Set r = sh1.Range(sh1.Range("A1"), r)
    
    Dim ir As Range, priceRng As Range
    For Each ir In r
        Set priceRng = ir.offset(0, EVO.E_FROM_SQ01_QUASI_TP04_SUM - 1)
        
        ' Application.Calculation = xlCalculationManual
        
        If priceRng.Value Like "*.*,??" Or priceRng.Value Like "*,??" Then
            tmpstrv = Replace(Replace(priceRng.Value, ".", ""), ",", "")
            
            If IsNumeric(tmpstrv) Then
                doubleValue = CDbl(tmpstrv) / 100#
                
                priceRng.Value = doubleValue
            End If
        End If
    Next ir
    
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub copyAndPasteAsValues(refRng As Range)


    Dim allRange As Range
    
    If refRng.offset(1, 0).Value <> "" Then
    
        Set allRange = refRng.Parent.Range(refRng, refRng.End(xlDown))
    Else
        Set allRange = refRng
    End If
    
    allRange.Copy
    allRange.PasteSpecial xlPasteValues
    
    refRng.Parent.Cells(1, 1).Select
End Sub



Public Sub delegacjaDlaProgresu(s1 As StatusHandler, ByRef h1 As Integer, lm As Long)
    
    s1.progress_increase
    h1 = h1 + 1
    
    If h1 > lm Then
        h1 = 0
        s1.hide
        Set s1 = Nothing
        Set s1 = New StatusHandler
        s1.init_statusbar lm
        s1.show
        s1.progress_increase
    End If
End Sub


Public Sub fillLabels(labelRefRange As Range)


    Dim refLabelInRegister As Range, x As Variant
    


    If labelRefRange.Parent.name Like "OUT1_*" Or _
        labelRefRange.Parent.name Like "OUT2_*" Or _
        labelRefRange.Parent.name Like "CONCAT_*" Then
    
        Set refLabelInRegister = ThisWorkbook.Sheets("forValidation").Range(G_REF_MOUNT_SQ1_OUT)
        With labelRefRange
            
            For x = EVO.E_FROM_SQ01_QUASI_TP04_DOMAIN To EVO.E_FROM_SQ01_QUASI_TP04_CURRENCY
                
                .offset(0, x - 1).Value = refLabelInRegister.offset(0, x - 1).Value
            Next x
        End With
        
        If labelRefRange.Parent.name Like "CONCAT_*" Then
            labelRefRange.offset(0, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_II_RATE_CALCED_SUM - 1).Value = _
                "SUM2"
            labelRefRange.offset(0, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_II_RATE_TO_EUR - 1).Value = _
                "RATE"
            labelRefRange.offset(0, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_II_RATE_UNIT_VALUE - 1).Value = _
                "UN3"
        End If
    
    ElseIf labelRefRange.Parent.name Like "N_*" Then
    
        Set refLabelInRegister = ThisWorkbook.Sheets("forValidation").Range(G_REF_MOUNT_N_SUPPLIERS_OUT)
        With labelRefRange
        
            For x = EVO.E_N_SUPPLIERS_COFOR To EVO.E_N_SUPPLIERS_INT_EXT_VEN
                .offset(0, x - 1).Value = refLabelInRegister.offset(0, x - 1).Value
            Next x
        End With
    
    Else
    End If
    
End Sub
