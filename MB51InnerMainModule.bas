Attribute VB_Name = "MB51InnerMainModule"
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







Public Sub runMainMB51Logic(d As Dictionary, Optional avoidFinalMsgBox As Boolean, Optional osh As Worksheet)



    ' Dim osh As Worksheet
    Set osh = ThisWorkbook.Sheets.Add
    
    osh.name = tryToRenameWorksheet(osh, "MB51_")
    
    Dim rng As Range, orng As Range
    


    
    Dim chbx As SAPFEWSELib.GuiCheckBox
    Dim txt As SAPFEWSELib.GuiTextedit
    
    
    If EVO.GlobalSapModule.sapGuiAuto Is Nothing Then
        Set sapGuiAuto = GetObject("SAPGUI")
        Set EVO.GlobalSapModule.sapGuiAuto = sapGuiAuto
        Set sapGui = sapGuiAuto.GetScriptingEngine
        Set EVO.GlobalSapModule.sapGui = sapGui
    Else
        Set sapGuiAuto = EVO.GlobalSapModule.sapGuiAuto
        Set sapGui = EVO.GlobalSapModule.sapGui
    End If
    
    Dim se As Object
    
    
    Set cnn = sapGui.Connections(0)
    
    Debug.Print cnn.ConnectionString
    Debug.Print cnn.Sessions.Count

    
    Set sess = cnn.Children(0)
    Set item = sess.Children(0)
    Debug.Print item.name
    
    Debug.Print sess.Children.Count
    
    
    ' Set item = sess.Children(0)
    Set item = sess.FindById("wnd[0]/usr")
    Debug.Print item.Children.Count

    
    
    sess.FindById("wnd[0]").Maximize
    
    
    ' im not proud of this "hack"
    ' ----------------------------------------------------
    Dim x17 As Variant
    For x17 = 0 To 10
        On Error Resume Next
        sess.FindById("wnd[0]/tbar[0]/btn[12]").Press
    Next x17
    ' ----------------------------------------------------
    
    sess.FindById("wnd[0]/tbar[0]/okcd").Text = "MB51"
    sess.FindById("wnd[0]").SendVKey 0
    
    
    Dim tb_rows As Long
    Dim tb_columns As Long
    Dim tb_col_order As Object
    
    
    
    Dim ileOstatni As Long
    ileOstatni = 0
    
    
    
    
    Dim st_h As StatusHandler
    
    Dim numHandler As NumberHandler
    Set numHandler = New NumberHandler
    
    
    
    Dim i_mb51 As MB51_InputItem, key As Variant
    Debug.Print "d.Count: " & d.Count
    
    
    ' auto enum variables
    ' ------------------------------------------------------------------------
    ' ------------------------------------------------------------------------
    ' ------------------------------------------------------------------------
    Dim autoDecisionOnTableLayout As E_MB51_AUTO_DECISION_LAYOUT
    Dim offsetForPrice As Integer, offsetForQty As Integer
    ' for loop ref
    Dim y_start As Integer, y_end As Integer
    
    Dim pcsPriceEnum As Integer, qtyEnum As Integer, montantDiEnum As Integer
    Dim extPcsPrice As Integer, unxEnum As Integer, pcsPriceCurrency As Integer, deviseEnum As Integer, currencyRateEnum As Integer
    Dim pcsPriceInEur As Integer
    
    Dim withIndexEnum As Integer, articleEnum As Integer
    Dim isCancelledEnum As Integer, mvtEnum As Integer, cwEnum As Integer
    
    Dim dateEnum As Integer
    Dim refEnum As Integer
    ' ------------------------------------------------------------------------
    ' ------------------------------------------------------------------------
    ' ------------------------------------------------------------------------
    
    
    Dim forEachstart As Boolean
    forEachstart = True
    
    For Each key In d.Keys
    
    
        Set i_mb51 = d(key)
        

            
        sess.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = CStr(i_mb51.mag)
        sess.FindById("wnd[0]/usr/ctxtBUDAT-LOW").Text = CStr(i_mb51.du)
        sess.FindById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = CStr(i_mb51.au)
        sess.FindById("wnd[0]/usr/ctxtBWART-LOW").Text = CStr(i_mb51.mvt1)
        sess.FindById("wnd[0]/usr/ctxtBWART-HIGH").Text = CStr(i_mb51.mvt2)
        sess.FindById("wnd[0]/usr/ctxtALV_DEF").Text = "/STANDARD_2"
            
            
        sess.FindById("wnd[0]/tbar[1]/btn[8]").Press
        sess.FindById("wnd[0]/tbar[1]/btn[48]").Press
            
            
        Set gv = Nothing
        
        
        On Error Resume Next
        Set gv = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            
            
            
        If Not gv Is Nothing Then
            Dim cols As Variant
            Set cols = gv.ColumnOrder
            Debug.Print "gv.RowCount: " & gv.RowCount
            

            
            
            
            
            If forEachstart Then
            
                ' Debug.Print cols(0)
                If Trim(cols(0)) = "MATNR" Then
                    autoDecisionOnTableLayout = E_MB51_AUTO_DECISION_LAYOUT_0
                    
                    offsetForPrice = Int(EVO.E_MB51_0_PCS_PRICE - E_MB51_0_MONTANT_DI)
                    offsetForQty = Int(E_MB51_0_PCS_PRICE - E_MB51_0_QTY)
                    
                    y_start = Int(EVO.E_MB51_0_ARTICLE)
                    y_end = Int(EVO.E_MB51_0_FOUR)
                    
                    pcsPriceEnum = Int(EVO.E_MB51_0_PCS_PRICE)
                    qtyEnum = Int(EVO.E_MB51_0_QTY)
                    montantDiEnum = Int(EVO.E_MB51_0_MONTANT_DI)
                    
                    
                    extPcsPrice = Int(E_MB51_0_EXT_PCS_PRICE)
                    unxEnum = Int(E_MB51_0_UN)
                    
                    pcsPriceCurrency = Int(E_MB51_0_EXT_PCS_PRICE_CURR)
                    deviseEnum = Int(E_MB51_0_DEVISE)
                    currencyRateEnum = Int(E_MB51_0_EXT_USED_RATE)
                    pcsPriceInEur = Int(E_MB51_0_EXT_PCS_PRICE_IN_EUR)
                    
                    withIndexEnum = Int(E_MB51_0_IS_WITH_INDEX)
                    
                    articleEnum = Int(EVO.E_MB51_0_ARTICLE)
                    
                    isCancelledEnum = Int(EVO.E_MB51_0_IS_CANCELLED)
                    mvtEnum = Int(EVO.E_MB51_0_MVT)
                    cwEnum = Int(EVO.E_MB51_0_CW)
                    
                    dateEnum = Int(EVO.E_MB51_0_DATE_SAISIE_LE)
                    
                    refEnum = Int(EVO.E_MB51_0_REF)
                Else
                    autoDecisionOnTableLayout = E_MB51_AUTO_DECISION_LAYOUT_NEW
                    
                    offsetForPrice = Int(E_MB51_NEW_PCS_PRICE - E_MB51_NEW_MONTANT_DI)
                    offsetForQty = Int(E_MB51_NEW_PCS_PRICE - E_MB51_NEW_QTY)
                    
                    y_start = Int(EVO.E_MB51_NEW_MVT)
                    y_end = Int(EVO.E_MB51_NEW_REF)
                    
                    pcsPriceEnum = Int(E_MB51_NEW_PCS_PRICE)
                    qtyEnum = Int(E_MB51_NEW_QTY)
                    montantDiEnum = Int(E_MB51_NEW_MONTANT_DI)
                    
                    
                    extPcsPrice = Int(E_MB51_NEW_EXT_PCS_PRICE)
                    unxEnum = Int(E_MB51_NEW_UNX)
                    
                    pcsPriceCurrency = Int(E_MB51_NEW_EXT_PCS_PRICE_CURR)
                    deviseEnum = Int(E_MB51_NEW_DEVISE)
                    currencyRateEnum = Int(E_MB51_NEW_EXT_USED_RATE)
                    pcsPriceInEur = Int(E_MB51_NEW_EXT_PCS_PRICE_IN_EUR)
                    
                    withIndexEnum = Int(E_MB51_NEW_IS_WITH_INDEX)
                    
                    articleEnum = Int(EVO.E_MB51_NEW_ARTICLE)
                    
                    isCancelledEnum = Int(EVO.E_MB51_NEW_IS_CANCELLED)
                    mvtEnum = Int(EVO.E_MB51_NEW_MVT)
                    cwEnum = Int(EVO.E_MB51_NEW_CW)
                    
                    dateEnum = Int(EVO.E_MB51_NEW_DATE_SAISIE_LE)
                    
                    refEnum = Int(EVO.E_MB51_NEW_REF)
                End If
            
            
            

                ' LABLES ------------------------------------------------
                
                mb51__fillLabels osh.Range("A1"), autoDecisionOnTableLayout
                
                ' -------------------------------------------------------
                
            
                
                
                Set orng = osh.Range("A2")
                orng.Select
                
                forEachstart = False
            End If
            
            
            
            
            Set st_h = Nothing
            Set st_h = New StatusHandler
            st_h.init_statusbar (gv.RowCount / 50)
            st_h.show
            DoEvents
        
        
            Dim x As Variant
            Dim y As Variant
            Dim tmp As Variant
        
        
        
        
            For x = 0 To gv.RowCount
                For y = y_start - 1 To y_end
                    
                    
                    If y = 0 Then
                    
                        If Trim(gv.GetCellValue(x, cols(y))) = "" Then
                            orng.Value = "X"
                        Else
                            orng.Value = gv.GetCellValue(x, cols(y))
                        End If
                    
                    ElseIf y = pcsPriceEnum - 1 Then
                        
                        orng.Offset(0, y).FormulaR1C1 = "=RC[-" & CStr(offsetForPrice) & "]" & _
                            "/RC[-" & CStr(offsetForQty) & "]"
                            
                            
                        If IsError(orng.Offset(0, y).Value) Then
                            orng.Offset(0, y).Value = 0
                        End If
                            
                    ElseIf y = qtyEnum - 1 Then
                        
                        On Error Resume Next
                        orng.Offset(0, y).Value = "'" & gv.GetCellValue(x, cols(y))
                        numHandler.parseStringProperlyToNum orng.Offset(0, y)
                        
                        
                    ElseIf y = montantDiEnum - 1 Then
                        
                        On Error Resume Next
                        orng.Offset(0, y).Value = "'" & gv.GetCellValue(x, cols(y))
                        numHandler.parseStringProperlyToNum orng.Offset(0, y)
                        
                        
                    Else
                        On Error Resume Next
                        orng.Offset(0, y).Value = gv.GetCellValue(x, cols(y))
                    End If
                Next y
            
            
                'E_MB51_EXT_PCS_PRICE
                'E_MB51_EXT_PCS_PRICE_CURR
                'E_MB51_EXT_USED_RATE
                'E_MB51_EXT_PCS_PRICE_IN_EUR
                'E_MB51_PRICE_RATIO
                'E_MB51_OKNOK
                
                ' extended fields
                
                 ' this element is an formula - be careful
                 ' E_MB51_PCS_PRICE
                 
                 ' schema
                 ' orng.Offset(0, y).Value
                 
                ' for extended elements there is no necesity fo for y loop
                ' this loop is only for data from sap sigapp
                
                ' price per UN
                orng.Offset(0, extPcsPrice - 1).Value = _
                    orng.Offset(0, pcsPriceEnum - 1).Value / (1# * findUnQty(orng.Offset(0, unxEnum - 1).Value))
                    
                ' original currency
                orng.Offset(0, pcsPriceCurrency - 1).Value = orng.Offset(0, deviseEnum - 1).Value
                ' rate on original currency to EUR
                orng.Offset(0, currencyRateEnum - 1).Value = _
                    findCurrRate(CStr(orng.Offset(0, deviseEnum - 1).Value))
                ' final price in eur per UN
                orng.Offset(0, pcsPriceInEur - 1).Value = _
                    orng.Offset(0, extPcsPrice - 1).Value / orng.Offset(0, currencyRateEnum - 1).Value
                
                    
                ' BOOL fields
                ' E_MB51_IS_WITH_INDEX
                ' E_MB51_IS_CANCELLED
                orng.Offset(0, withIndexEnum - 1).Value = _
                    withIndex(CStr(orng.Offset(0, articleEnum - 1)))
                    
                
                ' you need to make it after all loop ready with data
                'orng.Offset(0, EVO.E_MB51_IS_CANCELLED - 1).Value = _
                '    isCancelled(orng.Offset(0, EVO.E_MB51_MVT - 1), osh.Range(osh.Range("A2"), osh.Range("A2").End(xlDown)))


                 
        
            
                
                Set orng = orng.Offset(1, 0)
                
                
                
                
                If x Mod 50 = 0 Then
                
                
                    'Debug.Print gv.CurrentCellRow
                    gv.FirstVisibleRow = x
                    gv.CurrentCellRow = x
                    'Debug.Print gv.CurrentCellRow
                    
                    st_h.progress_increase
                    DoEvents
                    StatusForm.Repaint
                End If
                
                
            Next x
        

        

        
            
            st_h.hide
            Set st_h = Nothing
            
            

        
        
            sess.FindById("wnd[0]/tbar[0]/btn[15]").Press
            sess.FindById("wnd[0]/tbar[0]/btn[15]").Press
        
        End If
        
    
        ' Set rng = rng.Offset(1, 0)
            

    
    Next
    
    
    
    
    'orng.Offset(0, EVO.E_MB51_IS_CANCELLED - 1).Value = _
    '    isCancelled(orng.Offset(0, EVO.E_MB51_MVT - 1), osh.Range(osh.Range("A2"), osh.Range("A2").End(xlDown)))
    
    ' need to be outside of major loop
    For Each orng In osh.Range(osh.Range("A2"), osh.Range("A1").End(xlDown))
        orng.Offset(0, isCancelledEnum - 1).Value = _
            isCancelled(orng, _
                osh.Range(osh.Range("A2"), osh.Range("A2").End(xlDown)), _
                mvtEnum, montantDiEnum, refEnum)
            
        
        'On Error Resume Next
        'orng.Offset(0, EVO.E_MB51_CW - 1).Value = _
        '    Year(orng.Offset(0, EVO.E_MB51_DATE_CPT - 1).Value) & " CW" & _
        '    Application.WorksheetFunction.IsoWeekNum(CDbl(orng.Offset(0, EVO.E_MB51_DATE_CPT - 1).Value))
        
        orng.Offset(0, cwEnum - 1).Value = tryToAssignYearAndCW(orng, dateEnum)
    Next orng
    
    
    If avoidFinalMsgBox Then
    Else
        MsgBox "GOTOWE!", vbInformation
    End If
    
End Sub


Private Function tryToAssignYearAndCW(orng As Range, dateEnum As Integer) As String
    tryToAssignYearAndCW = ""
    
    Dim strD As String, y As String, cw As String
    strD = CStr(orng.Offset(0, dateEnum - 1).Value)
    
    If IsDate(strD) Then
        
        y = CStr(Year(strD))
        cw = Application.WorksheetFunction.IsoWeekNum(CDate(strD))
        tryToAssignYearAndCW = CStr(y) & " CW" & CStr(cw)
        
    ElseIf strD Like "??.??.????" Then
    
        y = Right(strD, 4)
        cw = Application.WorksheetFunction.IsoWeekNum(CDbl(CDate(Format(strD, "dd.mm.yyyy"))))
        tryToAssignYearAndCW = CStr(y) & " CW" & CStr(cw)
        
    End If
End Function

Private Function isCancelled(ar As Range, br As Range, _
    mvtEnum As Integer, montantDiEnum As Integer, refEnum As Integer) As Integer


    ' Debug.Print br.Address
    
    Dim rMvt As Range
    Set rMvt = ar.Offset(0, mvtEnum - 1)
    
    isCancelled = 0
    
    
    If CStr(rMvt.Value) = "102" Then
        isCancelled = 2
    ElseIf CStr(rMvt.Value) = "101" Then
        
        Dim ir As Range
        For Each ir In br
            If CStr(ir.Offset(0, mvtEnum - 1).Value) = "102" Then
                If Math.Abs(CDbl(ir.Offset(0, montantDiEnum - 1).Value)) = _
                    Math.Abs(CDbl(rMvt.Offset(0, montantDiEnum - mvtEnum).Value)) Then
                    
                    
                    ' same ref as well
                    If CStr(ir.Offset(0, refEnum - 1).Value) = _
                        CStr(rMvt.Offset(0, refEnum - mvtEnum).Value) Then
                        isCancelled = 1
                    End If
                    
                End If
            End If
        Next ir
    End If
End Function


Private Function withIndex(strArticle As String) As Integer
    withIndex = -1
    
    If strArticle Like "*-??" Then
        withIndex = 1
    ElseIf IsNumeric(strArticle) Then
        withIndex = 0
    End If
End Function


Private Function findUnQty(strUn As String) As Long
    findUnQty = 1
    
    'Dim strFormula As String
    Dim ref As Range
    Set ref = ThisWorkbook.Sheets("register").Range("J100")
    
    'strFormula = Replace(strFormula, "X", """" & CStr(strUn) & """")
    'findUnQty = Evaluate(strFormula)
    'findUnQty = CLng(findUnQty)
    
    Do
        If UCase(Trim(ref.Value)) = UCase(Trim(strUn)) Then
            findUnQty = CLng(ref.Offset(0, 1).Value)
            Exit Function
        End If
        
        Set ref = ref.Offset(1, 0)
    Loop Until Trim(ref.Value) = ""
    
End Function


Private Function findCurrRate(strCurr As String) As Double
    findCurrRate = 1#
    
    Dim strFormula As String
    strFormula = ThisWorkbook.Sheets("register").Range("D99").Formula
    strFormula = Replace(strFormula, "X", """" & CStr(strCurr) & """")
    
    findCurrRate = Evaluate(strFormula)
    
    findCurrRate = CDbl(findCurrRate)
End Function


Private Sub mb51__fillLabels(labelRefRange As Range, autoDecisionOnTableLayout As E_MB51_AUTO_DECISION_LAYOUT)


    Dim fv As Worksheet
    Set fv = ThisWorkbook.Sheets("forValidation")
    Dim rfv As Range
    
    If autoDecisionOnTableLayout = E_MB51_AUTO_DECISION_LAYOUT_NEW Then
        Set rfv = fv.Range("D35")
    Else
        Set rfv = fv.Range("D38")
    End If
    
    Dim ref As Range
    Set ref = labelRefRange.Parent.Cells(1, 1)
    
    Do
    
        ref.Value = rfv.Value
        ref.Interior.Color = rfv.Interior.Color
    
        Set rfv = rfv.Offset(0, 1)
        Set ref = ref.Offset(0, 1)
        
    Loop While Trim(rfv.Value) <> ""

End Sub
