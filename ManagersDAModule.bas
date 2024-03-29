Attribute VB_Name = "ManagersDAModule"
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


Global flagForCycle As Boolean



Public Sub multi__ManagersDa3()
    
    innerGetManagersDa3 ThisWorkbook.Sheets("CONCAT_20210719_CJ")
    innerGetManagersDa3 ThisWorkbook.Sheets("CONCAT_20210719_eK9")
    innerGetManagersDa3 ThisWorkbook.Sheets("CONCAT_20210719_X250")
End Sub


Public Sub getManagersDa(ictrl As IRibbonControl)

    innerGetManagersDa ActiveSheet
    
End Sub


Public Sub loopForAllConcatsForGetManagersDa()
    
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        
        If sh.name Like "CONCAT_*" Then
            sh.Activate
            innerGetManagersDa ActiveSheet
        End If
    Next sh
End Sub

Public Sub innerGetManagersDa(sh1 As Worksheet, Optional osh As Worksheet)

    Dim mainCollection As Collection, semicolonedDomains As String
    ' semicolonedDomains treated with side effect inside below function
    semicolonedDomains = ""
    Set mainCollection = prepareInputFOrManagersDa(sh1, semicolonedDomains)
    inner0GetManagersDa semicolonedDomains, mainCollection, sh1, osh
    
End Sub

Public Sub innerGetManagersDa3(sh1 As Worksheet, Optional osh As Worksheet)

    Dim mainCollection As Collection, semicolonedDomains As String
    ' semicolonedDomains treated with side effect inside below function
    semicolonedDomains = ""
    Set mainCollection = prepareInputFOrManagersDa(sh1, semicolonedDomains)
    inner0GetManagersDa3 semicolonedDomains, mainCollection, sh1, osh
    
End Sub



Public Sub inner0GetManagersDa3(semicolonedDomains As String, c As Collection, ish As Worksheet, Optional osh As Worksheet)


    EVO.SuppressingMessageModule.KillMessageFilter
    
    If validateIfTheSemicolonedDomainsAreSemicolonedDomains(semicolonedDomains) Then
    
    
    
       
    
    
        Dim arr As Variant
        semicolonedDomains = Replace(semicolonedDomains, " ", "")
        arr = Split(semicolonedDomains, ";")
        
        

        
        Set osh = ThisWorkbook.Sheets.Add
        osh.name = tryToRenameWorksheet(osh, "MANAGERS_DA_")
        
        Dim rng As Range, orng As Range
        
        
        ' LABLES ------------------------------------------------
        
        managers_DA__fillLabels osh.Range("A1")
        
        ' -------------------------------------------------------
        
        ManagersDaLoading.show vbModeless
    
        initManagersDaLoading
        ' connected with ManagersDaLoading!
        ManagersDaLoading.Repaint
        flagForCycle = True
        cycleForIteration
        
        
        Set orng = osh.Range("A2")
        orng.Select
        
        
        Dim chbx As Variant '  SAPFEWSELib.GuiCheckBox
        Dim txt As Variant ' SAPFEWSELib.GuiTextedit
        
        
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
        
        'Debug.Print cnn.ConnectionString
        'Debug.Print cnn.Sessions.Count
    
        
        Set sess = cnn.Children(0)
        Set item = sess.Children(0)
        ' Debug.Print item.name
        
        ' Debug.Print sess.Children.Count
        
        
        ' Set item = sess.Children(0)
        Set item = sess.FindById("wnd[0]/usr")
        ' Debug.Print item.Children.Count
    
        
        
        sess.FindById("wnd[0]").Maximize
        
        
        ' im not proud of this "hack"
        ' ----------------------------------------------------
        Dim x17 As Variant
        For x17 = 0 To 10
            On Error Resume Next
            sess.FindById("wnd[0]/tbar[0]/btn[12]").press
        Next x17
        ' ----------------------------------------------------
        
        ' loop for each domin
        ' beg of loop
        Dim cols As Variant
        Dim x As Variant

        
        
        sess.FindById("wnd[0]").Maximize
        sess.FindById("wnd[0]/tbar[0]/okcd").Text = "Y_DI3_80000594"
        sess.FindById("wnd[0]").sendVKey 0
        
        
        
        

        Dim stdStr As String

        
        
    
        ' NOA CODE here
        ' session.findById("wnd[0]/usr/txtSP$00026-LOW").text = "*"
        sess.FindById("wnd[0]/usr/txtSP$00026-LOW").Text = "*"
        
        ' type of output - darwin manager
        sess.FindById("wnd[0]/usr/ctxt%ALVL").Text = "/MANAGER"
          
        'sess.FindById("wnd[0]/usr/txtSP$00004-LOW").Text = "375" ' this domain should be iterate
        sess.FindById("wnd[0]/usr/txtSP$00004-LOW").Text = Left(CStr(arr(x)), 3) ' this domain should be iterate
    
    
    

        sess.FindById("wnd[0]/usr/btn%_SP$00004_%_APP_%-VALU_PUSH").press
        
        sess.FindById("wnd[1]/tbar[0]/btn[16]").press
    
        For x = LBound(arr) To UBound(arr)
        
            ' session.
            '  findById(
            '   "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]")
            '    .text = "375"
            'sess.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1," & _
            '    CStr(x) & "]").SetFocus
            sess.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1," & _
                CStr(x) & "]").Text = Left(CStr(arr(x)), 3)
        
        Next x
    
        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        sess.FindById("wnd[1]/tbar[0]/btn[8]").press
    
        ' final submit!
        sess.FindById("wnd[0]/tbar[1]/btn[8]").press
    
        
        
        ' Debug.Print "inside interation!"
        orng.Select
        
        cycleForIteration
    
        
        Set gv = Nothing
        Set gv = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        
        
    
        Set cols = Nothing
        Set cols = gv.ColumnOrder
        Debug.Print "gv.RowCount: " & gv.RowCount
    
    
        If Not gv Is Nothing Then
            Set cols = gv.ColumnOrder
            Debug.Print "gv.RowCount: " & gv.RowCount
            
            'Dim st_h As StatusHandler
            'Set st_h = Nothing
            'Set st_h = New StatusHandler
            'st_h.init_statusbar (gv.RowCount / 50)
            'st_h.show
            DoEvents
        
        
            Dim x1 As Variant
            Dim y1 As Variant
            Dim tmp As Variant
        
        
            Dim lastNotEmptyPartNumber As String
            lastNotEmptyPartNumber = ""
            
            
            
        
            
            For x1 = 0 To (gv.RowCount - 1)
                ' For y = EVO.E_MANAGERS_DA_ARTICLE - 1 To EVO.E_MANAGERS_DA_RU - 1
                ' next y
                
                ' Debug.Print orng.Address
                
                With orng
                
                    
                
                    ' Debug.Print "PN: " & gv.GetCellValue(x1, cols(EVO.E_MANAGERS_DA_ARTICLE - 1))
                    If Trim(gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA_ARTICLE - 1))) <> "" Then
                        lastNotEmptyPartNumber = Trim(gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA_ARTICLE - 1)))
                    End If
                    .offset(0, EVO.E_MANAGERS_DA_ARTICLE - 1).Value = lastNotEmptyPartNumber
                    
                    
                    .offset(0, EVO.E_MANAGERS_DA_ACHETEUR - 1).Value = gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA_ACHETEUR - 1))
                    .offset(0, EVO.E_MANAGERS_DA_RU - 1).Value = gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA_RU - 1))
                    
                    ' new from 0.92
                    .offset(0, EVO.E_MANAGERS_DA.E_MANAGERS_DA_FAMILY - 1).Value = _
                        gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA.E_MANAGERS_DA_FAMILY - 1))
                    .offset(0, EVO.E_MANAGERS_DA.E_MANAGERS_DA_GROUP - 1).Value = _
                        gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA.E_MANAGERS_DA_GROUP - 1))
                End With
                
                Set orng = orng.offset(1, 0)
                
                
                
                
                If x1 Mod 50 = 0 Then
                
                
                    'Debug.Print gv.CurrentCellRow
                    gv.FirstVisibleRow = x1
                    gv.CurrentCellRow = x1
                    'Debug.Print gv.CurrentCellRow
                    
                    orng.Select
                    cycleForIteration
                    ' ManagersDaLoading.show vbModeless
                End If
            Next x1
                
            
            Debug.Print gv.id & Chr(10) & gv.Children.count & " " & gv.ColumnOrder(0)
            
            If Not gv Is Nothing Then
                sess.FindById("wnd[0]/tbar[0]/btn[15]").press
            End If

        
        End If
    End If


    
    flagForCycle = False
    ManagersDaLoading.hide
    
    
    ' evo.SuppressingMessageModule.RestoreMessageFilter
End Sub


Public Sub inner0GetManagersDa(semicolonedDomains As String, c As Collection, ish As Worksheet, Optional osh As Worksheet)


    
    If validateIfTheSemicolonedDomainsAreSemicolonedDomains(semicolonedDomains) Then
    
    
    
        ' prepare batches for the list inside Y_DI3_80000594
        ' it is taking too looong!
        
        ' dic in dic - to avoid duplicates!
        Dim majorDic As New Dictionary
        Dim innerDic As Dictionary
        Dim justReferenceToAvoidDuplicates As New Dictionary
        Dim key As Variant, ikey As Variant
        
        
        Dim el As Variant
        
        Dim iter As Long, counterForInnerDic As Long
        iter = 0
        counterForInnerDic = 0
        ' this loop creating batches of input data
        ' for the sigapp screen
        Dim vl As String
        For Each el In c
        
        
            If Not justReferenceToAvoidDuplicates.Exists(el) Then
            
                justReferenceToAvoidDuplicates.Add el, 1
        
            
                If iter Mod 200 = 0 Or iter = (c.count - 1) Then
                
                    Debug.Print "iter and counter: " & iter & " " & c.count
                    
                    counterForInnerDic = counterForInnerDic + 1
                    Set innerDic = New Dictionary
                    majorDic.Add counterForInnerDic, innerDic
                    
                End If
            
            
                vl = CStr(Split(el, "-")(0))
                
                If Trim(vl) <> "" Then
                    
                    Set innerDic = majorDic(counterForInnerDic)
                    
                    If innerDic.Exists(vl) Then
                    Else
                        innerDic.Add vl, 1
                        
                    End If
                    
                    
    
                End If
                
                iter = iter + 1
            
            Else
                ' here we have duplicate so nothing to do...
            End If
        Next el
        
        
        ' Batches created properly
        ' but for now not in usage!
        
        For Each key In majorDic.Keys
        
            Set innerDic = majorDic(key)
            
            
            
            'For Each ikey In innerDic.Keys
            '   Debug.Print key & " " & ikey
            'Next
            
            Debug.Print "TEST: innerDic.Count: " & innerDic.count
        Next
    
    
        ' at the end not used - diff idea!
        'Dim tmpWorksheetForArticlesFromMb51 As Worksheet
        'Dim tmpRangrInWorksheetForArticlesFromMb51 As Range
    
    
        Dim arr As Variant
        semicolonedDomains = Replace(semicolonedDomains, " ", "")
        arr = Split(semicolonedDomains, ";")
        
        
        
        
        ' extra tmp worksheet for articles from mb51 proxy
        'Set tmpWorksheetForArticlesFromMb51 = ThisWorkbook.Sheets.Add
        'tmpWorksheetForArticlesFromMb51.name = tryToRenameWorksheet(tmpWorksheetForArticlesFromMb51, "TMP_ARTICLE_LIST_")
        'Set tmpRangrInWorksheetForArticlesFromMb51 = tmpWorksheetForArticlesFromMb51.Cells(1, 1)
        '
        'tmpRangrInWorksheetForArticlesFromMb51.Value = "ARTICLES"
        'Set tmpRangrInWorksheetForArticlesFromMb51 = tmpRangrInWorksheetForArticlesFromMb51.Offset(1, 0)
        'Dim a As Variant
        'For Each a In c
        '
        '    tmpRangrInWorksheetForArticlesFromMb51.Value = CStr(a)
        '    Set tmpRangrInWorksheetForArticlesFromMb51 = tmpRangrInWorksheetForArticlesFromMb51.Offset(1, 0)
        '
        'Next c
        '
        '
        'tmpWorksheetForArticlesFromMb51.Range(tmpWorksheetForArticlesFromMb51.Cells(1, 1), _
        '    tmpRangrInWorksheetForArticlesFromMb51).RemoveDuplicates Array(1), xlYes
        
        
        ' not using dim anymore is in optional - side effect trick
        ' Dim osh As Worksheet
        
        Set osh = ThisWorkbook.Sheets.Add
        osh.name = tryToRenameWorksheet(osh, "MANAGERS_DA_")
        
        Dim rng As Range, orng As Range
        
        
        ' LABLES ------------------------------------------------
        
        managers_DA__fillLabels osh.Range("A1")
        
        ' -------------------------------------------------------
        
        ManagersDaLoading.show vbModeless
    
        initManagersDaLoading
        ' connected with ManagersDaLoading!
        ManagersDaLoading.Repaint
        flagForCycle = True
        cycleForIteration
        
        
        Set orng = osh.Range("A2")
        orng.Select
        
        
        Dim chbx As Variant '  SAPFEWSELib.GuiCheckBox
        Dim txt As Variant ' SAPFEWSELib.GuiTextedit
        
        
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
        
        'Debug.Print cnn.ConnectionString
        'Debug.Print cnn.Sessions.Count
    
        
        Set sess = cnn.Children(0)
        Set item = sess.Children(0)
        ' Debug.Print item.name
        
        ' Debug.Print sess.Children.Count
        
        
        ' Set item = sess.Children(0)
        Set item = sess.FindById("wnd[0]/usr")
        ' Debug.Print item.Children.Count
    
        
        
        sess.FindById("wnd[0]").Maximize
        
        
        ' im not proud of this "hack"
        ' ----------------------------------------------------
        Dim x17 As Variant
        For x17 = 0 To 10
            On Error Resume Next
            sess.FindById("wnd[0]/tbar[0]/btn[12]").press
        Next x17
        ' ----------------------------------------------------
        
        ' loop for each domin
        ' beg of loop
        Dim cols As Variant
        Dim x As Variant

        
        
        sess.FindById("wnd[0]").Maximize
        sess.FindById("wnd[0]/tbar[0]/okcd").Text = "Y_DI3_80000594"
        sess.FindById("wnd[0]").sendVKey 0
        
        
        
        

        Dim stdStr As String
        'For Each key In majorDic.Keys
        '
        '    Set innerDic = majorDic(key)
        '    For Each ikey In innerDic.Keys
        '
        '        Debug.Print key & " " & ikey
        '
        '
        '    Next
        'Next
        
        
        For Each key In majorDic.Keys
        
            Set innerDic = majorDic(key)
            
        
            ' NOA CODE here
            ' session.findById("wnd[0]/usr/txtSP$00026-LOW").text = "*"
            sess.FindById("wnd[0]/usr/txtSP$00026-LOW").Text = "*"
            
            ' type of output - darwin manager
            sess.FindById("wnd[0]/usr/ctxt%ALVL").Text = "/MANAGER"
              
            'sess.FindById("wnd[0]/usr/txtSP$00004-LOW").Text = "375" ' this domain should be iterate
            sess.FindById("wnd[0]/usr/txtSP$00004-LOW").Text = Left(CStr(arr(x)), 3) ' this domain should be iterate
        
        
        

            sess.FindById("wnd[0]/usr/btn%_SP$00004_%_APP_%-VALU_PUSH").press
            
            sess.FindById("wnd[1]/tbar[0]/btn[16]").press
        
            For x = LBound(arr) To UBound(arr)
            
                ' session.
                '  findById(
                '   "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]")
                '    .text = "375"
                'sess.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1," & _
                '    CStr(x) & "]").SetFocus
                sess.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1," & _
                    CStr(x) & "]").Text = Left(CStr(arr(x)), 3)
            
            Next x
        
            sess.FindById("wnd[1]/tbar[0]/btn[0]").press
            sess.FindById("wnd[1]/tbar[0]/btn[8]").press
        
            '
            stdStr = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I"
            '
            '' list of pns
            With sess
                .FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").press
                
                
                Dim a As Variant, a_i As Long
                Dim VerticalScrollbar_Position As Integer
                Dim verticalstr As String
                verticalstr = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
                VerticalScrollbar_Position = 0
                
                
                'clear prev list
                sess.FindById("wnd[1]/tbar[0]/btn[16]").press
                
                a_i = 0
                For Each ikey In innerDic.Keys
                    ' session
                    '   .findById(
                    '   "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/
                    '       txtRSCSEL-SLOW_I[1,0]")
                    '           .text = "9831503380"
                    
                    
                    
                    
                    If a_i < 7 Then
                    
                    
                        '.FindById(stdStr & "[1," & CStr(a_i) & "]").SetFocus
                        .FindById(stdStr & "[1," & CStr(a_i) & "]").Text = CStr(ikey)
                        
                        
                    Else
                        ' a_i = 0
                        ' session.findById(
                        '   "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE
                        '       ").verticalScrollbar.position = 2
                        
                        If a_i Mod 7 = 0 Then
                            sess.FindById(verticalstr).VerticalScrollbar.Position = a_i
                            cycleForIteration
                        End If
                        
                        '.FindById(stdStr & "[1," & CStr((a_i Mod 5) + 1) & "]").SetFocus
                        .FindById(stdStr & "[1," & CStr((a_i Mod 7) + 1) & "]").Text = CStr(ikey)
                        
                        
                    End If
                    
                    a_i = a_i + 1
                Next
            End With
            
            sess.FindById("wnd[1]/tbar[0]/btn[0]").press
            sess.FindById("wnd[1]/tbar[0]/btn[8]").press
        
        
            ' final submit!
            sess.FindById("wnd[0]/tbar[1]/btn[8]").press
        
            
            
            ' Debug.Print "inside interation!"
            orng.Select
            
            cycleForIteration
        
            
            Set gv = Nothing
            Set gv = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            
            
        
            Set cols = Nothing
            Set cols = gv.ColumnOrder
            Debug.Print "gv.RowCount: " & gv.RowCount
        
        
            If Not gv Is Nothing Then
                Set cols = gv.ColumnOrder
                Debug.Print "gv.RowCount: " & gv.RowCount
                
                'Dim st_h As StatusHandler
                'Set st_h = Nothing
                'Set st_h = New StatusHandler
                'st_h.init_statusbar (gv.RowCount / 50)
                'st_h.show
                DoEvents
            
            
                Dim x1 As Variant
                Dim y1 As Variant
                Dim tmp As Variant
            
            
                Dim lastNotEmptyPartNumber As String
                lastNotEmptyPartNumber = ""
                
                
                
            
                
                For x1 = 0 To (gv.RowCount - 1)
                    ' For y = EVO.E_MANAGERS_DA_ARTICLE - 1 To EVO.E_MANAGERS_DA_RU - 1
                    ' next y
                    
                    ' Debug.Print orng.Address
                    
                    With orng
                    
                        
                    
                        ' Debug.Print "PN: " & gv.GetCellValue(x1, cols(EVO.E_MANAGERS_DA_ARTICLE - 1))
                        If Trim(gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA_ARTICLE - 1))) <> "" Then
                            lastNotEmptyPartNumber = Trim(gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA_ARTICLE - 1)))
                        End If
                        .offset(0, EVO.E_MANAGERS_DA_ARTICLE - 1).Value = lastNotEmptyPartNumber
                        
                        
                        .offset(0, EVO.E_MANAGERS_DA_ACHETEUR - 1).Value = gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA_ACHETEUR - 1))
                        .offset(0, EVO.E_MANAGERS_DA_RU - 1).Value = gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA_RU - 1))
                        
                        ' new from 0.92
                        .offset(0, EVO.E_MANAGERS_DA.E_MANAGERS_DA_FAMILY - 1).Value = _
                            gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA.E_MANAGERS_DA_FAMILY - 1))
                        .offset(0, EVO.E_MANAGERS_DA.E_MANAGERS_DA_GROUP - 1).Value = _
                            gv.getCellValue(x1, cols(EVO.E_MANAGERS_DA.E_MANAGERS_DA_GROUP - 1))
                    End With
                    
                    Set orng = orng.offset(1, 0)
                    
                    
                    
                    
                    If x1 Mod 50 = 0 Then
                    
                    
                        'Debug.Print gv.CurrentCellRow
                        gv.FirstVisibleRow = x1
                        gv.CurrentCellRow = x1
                        'Debug.Print gv.CurrentCellRow
                        
                        orng.Select
                        cycleForIteration
                        ' ManagersDaLoading.show vbModeless
                    End If
                Next x1
                    
                
                Debug.Print gv.id & Chr(10) & gv.Children.count & " " & gv.ColumnOrder(0)
                
                If Not gv Is Nothing Then
                    sess.FindById("wnd[0]/tbar[0]/btn[15]").press
                End If

            
            End If

        Next
        
    End If
    
    flagForCycle = False
    ManagersDaLoading.hide
End Sub



Private Function prepareInputFOrManagersDa(ish As Worksheet, ByRef semicolonedDomains As String) As Collection

    Set prepareInputFOrManagersDa = Nothing
    Dim tmpColl As New Collection
    
    
    Dim articleEnum As Integer, magEnum As Integer
    
    Dim vd As Range
    
    
    If UCase(ish.Cells(1, 1).Value) = "ARTICLE" Then
        articleEnum = EVO.E_MB51_0_ARTICLE
        magEnum = EVO.E_MB51_0_MAG
        Set vd = ThisWorkbook.Sheets("forValidation").Range("D38")
    Else
        articleEnum = EVO.E_MB51_NEW_ARTICLE
        magEnum = EVO.E_MB51_NEW_MAG
        Set vd = ThisWorkbook.Sheets("forValidation").Range("D35")
    End If
    
    
    Dim ir As Range
    
    If validMb51data(ish, vd) Then
        ' go with logic for reception report!
        Set ir = ish.Range("A2")
        
        Do
        
            ' just add list of part numbers into collection
            ' -------------------------------------------------------------------
            ' ir is column A so it might be in raw data still some "X"
            ' we need to avoid this...
            If Trim(ir.Value) <> "X" Then
                tmpColl.Add CStr(ir.offset(0, articleEnum - 1).Value)
            ' -------------------------------------------------------------------
            
                
                ' only uniq!
                ' --------------------------------------------------------------------------------------------------
                If semicolonedDomains Like "*" & CStr(ir.offset(0, magEnum - 1).Value) & "*" Then
                Else
                    semicolonedDomains = semicolonedDomains & CStr(ir.offset(0, magEnum - 1).Value) & ";"
                End If
                ' --------------------------------------------------------------------------------------------------
            End If
            
            Set ir = ir.offset(1, 0)
        Loop Until Trim(ir.Value) = ""
        
        
        Set prepareInputFOrManagersDa = tmpColl
        
    ElseIf valid_TP04_data(ish) Then
    
    
        articleEnum = EVO.E_ADJUSTED_SQ01_Reference
        magEnum = EVO.E_ADJUSTED_SQ01_DOMAIN
        
        Set ir = ish.Range("A2")
        
        Do
        
        

            tmpColl.Add CStr(ir.offset(0, articleEnum - 1).Value)
            

            If semicolonedDomains Like "*" & CStr(ir.offset(0, magEnum - 1).Value) & "*" Then
            Else
            
                If Trim(CStr(ir.offset(0, magEnum - 1).Value)) <> "" Then
                    If Len(CStr(ir.offset(0, magEnum - 1).Value)) = 3 Then
                        semicolonedDomains = semicolonedDomains & CStr(ir.offset(0, magEnum - 1).Value) & ";"
                    End If
                End If
            End If

            
            Set ir = ir.offset(1, 0)
        Loop Until Trim(ir.Value) = ""
        
        
        
        Set prepareInputFOrManagersDa = tmpColl
        
        
        Debug.Print "prepareInputFOrManagersDa: " & prepareInputFOrManagersDa.count
        
        
    ElseIf specialValidConcat(ish) Then
    
    
        articleEnum = 2
        magEnum = 1
        
        Set ir = ish.Range("A2")
        
        Do
        
        

            tmpColl.Add CStr(ir.offset(0, articleEnum - 1).Value)
            

            If semicolonedDomains Like "*" & CStr(ir.offset(0, magEnum - 1).Value) & "*" Then
            Else
            
                If Trim(CStr(ir.offset(0, magEnum - 1).Value)) <> "" Then
                    If Len(CStr(ir.offset(0, magEnum - 1).Value)) = 3 Then
                        semicolonedDomains = semicolonedDomains & CStr(ir.offset(0, magEnum - 1).Value) & ";"
                    End If
                End If
            End If

            
            Set ir = ir.offset(1, 0)
        Loop Until Trim(ir.Value) = ""
        
        
        
        Set prepareInputFOrManagersDa = tmpColl
        
        
        Debug.Print "prepareInputFOrManagersDa: " & prepareInputFOrManagersDa.count
        
        
    Else
        MsgBox "Wrong standard of activesheet!", vbCritical
        End
    End If
End Function





Private Function specialValidConcat(sh1 As Worksheet) As Boolean
    specialValidConcat = False
    
    If sh1.Cells(1, 1).Value = "DOMAIN" Then
        If sh1.Cells(1, 2).Value = "ARTICLE" Then
            If sh1.Cells(1, 7).Value = "DIV" Then
                specialValidConcat = True
            End If
        End If
    End If
End Function


Private Function validateIfTheSemicolonedDomainsAreSemicolonedDomains(str As String)
    validateIfTheSemicolonedDomainsAreSemicolonedDomains = False
    
    Dim arr As Variant
    arr = Split(str, ";")
    
    If UBound(arr) > 0 Then
        validateIfTheSemicolonedDomainsAreSemicolonedDomains = True
    End If
End Function

Private Sub managers_DA__fillLabels(orng As Range)

    orng.offset(0, EVO.E_MANAGERS_DA_ARTICLE - 1).Value = "PN"
    orng.offset(0, EVO.E_MANAGERS_DA_ACHETEUR - 1).Value = "ACHETEUR"
    orng.offset(0, EVO.E_MANAGERS_DA_RU - 1).Value = "RU"
    
    ' new from 0.92
    orng.offset(0, EVO.E_MANAGERS_DA.E_MANAGERS_DA_FAMILY - 1).Value = "FAMILY ACHAT"
    orng.offset(0, EVO.E_MANAGERS_DA.E_MANAGERS_DA_GROUP - 1).Value = "GROUP ACHETEUR"
    
End Sub





Private Sub initManagersDaLoading()
    
    ManagersDaLoading.TextBox2.Width = 0
End Sub

Private Sub iterateManagersDaLoading()
    
    ManagersDaLoading.TextBox2.Width = ManagersDaLoading.TextBox2.Width + 20
    
    If ManagersDaLoading.TextBox2.Width > ManagersDaLoading.TextBox1.Width Then
        ManagersDaLoading.TextBox2.Width = 0
    End If
    
    ManagersDaLoading.Repaint

End Sub

Private Sub cycleForIteration()
    Dim alertTime As Variant
    If flagForCycle Then
        iterateManagersDaLoading
    End If
End Sub






' ----------
Public Sub fillReceptionManagersDaColumn(mb51Out As Worksheet, managersDaSh As Worksheet)



    Application.Calculation = xlCalculationManual
    

    Dim rfa As Range, rm As Range
    Set rfa = mb51Out.Cells(2, EVO.E_MB51_0_ARTICLE)
    
    
    Do
        ' starting from beg every time!
        Set rm = managersDaSh.Cells(2, EVO.E_MANAGERS_DA_ARTICLE)
        
        Do
            
            If rm.Value = Split(rfa.Value, "-")(0) Then
                rfa.offset(0, EVO.E_MB51_0_RU - EVO.E_MB51_0_ARTICLE).Value = _
                    rm.offset(0, EVO.E_MANAGERS_DA_RU - EVO.E_MANAGERS_DA_ARTICLE).Value
                    
                rfa.offset(0, EVO.E_MB51_0_MANAGER_DA - EVO.E_MB51_0_ARTICLE).Value = _
                    rm.offset(0, EVO.E_MANAGERS_DA_ACHETEUR - EVO.E_MANAGERS_DA_ARTICLE).Value
                    
                rfa.offset(0, EVO.E_MB51_0_FAMILY - EVO.E_MB51_0_ARTICLE).Value = _
                    rm.offset(0, EVO.E_MANAGERS_DA_FAMILY - EVO.E_MANAGERS_DA_ARTICLE).Value
                    
                rfa.offset(0, EVO.E_MB51_0_GROUP - EVO.E_MB51_0_ARTICLE).Value = _
                    rm.offset(0, EVO.E_MANAGERS_DA_GROUP - EVO.E_MANAGERS_DA_ARTICLE).Value
                    
                ' Application.Calculation = xlCalculationManual
            End If
            Set rm = rm.offset(1, 0)
        Loop Until Trim(rm.Value) = ""
        
        Set rfa = rfa.offset(1, 0)
    Loop Until Trim(rfa.Value) = ""
    
    Application.Calculation = xlCalculationAutomatic
End Sub



' ----------
Public Sub fillGreenLightManagersDaColumn(ash As Worksheet, managersDaSh As Worksheet)


    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim rfa As Range, rm As Range
    Set rfa = ash.Cells(2, EVO.E_ADJUSTED_SQ01_Reference)
    
    Dim strPn As String
    
    Dim bufferDic17 As Dictionary
    Set bufferDic17 = New Dictionary
    
    
    Dim areaOfRefInManagersDa As Range
    Set areaOfRefInManagersDa = managersDaSh.Range(managersDaSh.Cells(2, EVO.E_MANAGERS_DA_ARTICLE), managersDaSh.Cells(2, EVO.E_MANAGERS_DA_ARTICLE).End(xlDown))
    Debug.Print "areaOfRefInManagersDa: " & areaOfRefInManagersDa.Address
    Dim match1 As Range
    Set match1 = Nothing
    
    
    
    ' standard looking for data
    Do

        strPn = Split(rfa.Value, "-")(0)
        
        ' 0. opcja jesli juz w buferze mamy dane
        Set match1 = Nothing
        
        If IsEmpty(bufferDic17(strPn)) Then
            ' Debug.Print "IsEmpty(bufferDic17(strPn))"
        Else
            On Error Resume Next
            Set match1 = managersDaSh.Range(CStr(bufferDic17(strPn)))
        End If
        
        ' Debug.Print bufferDic17.count
        
        
        ' 0 take from dic buffer
        If Not match1 Is Nothing Then
            ' =========================================================================================
            ' =========================================================================================
            rfa.offset(0, EVO.E_ADJUSTED_SQ01_RU - 1).Value = _
                rm.offset(0, EVO.E_MANAGERS_DA_RU - EVO.E_MANAGERS_DA_ARTICLE).Value
                
            rfa.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_DA - 1).Value = _
                rm.offset(0, EVO.E_MANAGERS_DA_ACHETEUR - EVO.E_MANAGERS_DA_ARTICLE).Value
                
                
            'new 092
            rfa.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_FAMILY - 1).Value = _
                rm.offset(0, EVO.E_MANAGERS_DA_FAMILY - EVO.E_MANAGERS_DA_ARTICLE).Value
                
            rfa.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_GROUP - 1).Value = _
                rm.offset(0, EVO.E_MANAGERS_DA_GROUP - EVO.E_MANAGERS_DA_ARTICLE).Value
            ' =========================================================================================
            ' =========================================================================================
        Else
        
        
            ' 1. najpierw robimy std find
            Set match1 = Nothing
            On Error Resume Next
            Set match1 = areaOfRefInManagersDa.Find(strPn)
            
            
            
            ' 2. jesli sukces dla std find wtedy po prostu przepisujemy dane
            If Not match1 Is Nothing Then
            
                
                Set rm = match1
                
                
                If IsEmpty(bufferDic17(strPn)) Then
                    ' Debug.Print "we already have this key"
                    bufferDic17(strPn) = match1.AddressLocal
                Else
                    bufferDic17.Add strPn, match1.AddressLocal
                End If
                
                
                ' Debug.Print bufferDic17.count
                
                ' =========================================================================================
                ' =========================================================================================
                rfa.offset(0, EVO.E_ADJUSTED_SQ01_RU - 1).Value = _
                    rm.offset(0, EVO.E_MANAGERS_DA_RU - EVO.E_MANAGERS_DA_ARTICLE).Value
                    
                rfa.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_DA - 1).Value = _
                    rm.offset(0, EVO.E_MANAGERS_DA_ACHETEUR - EVO.E_MANAGERS_DA_ARTICLE).Value
                    
                    
                'new 092
                rfa.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_FAMILY - 1).Value = _
                    rm.offset(0, EVO.E_MANAGERS_DA_FAMILY - EVO.E_MANAGERS_DA_ARTICLE).Value
                    
                rfa.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_GROUP - 1).Value = _
                    rm.offset(0, EVO.E_MANAGERS_DA_GROUP - EVO.E_MANAGERS_DA_ARTICLE).Value
                ' =========================================================================================
                ' =========================================================================================
                
            Else
            
                ' jesli dane sie nie znalazly robimy loop dlugi
                ' ------------------------------------------------
                ' jesli wszystko zgodne z planem - ta logika w ogole juz nie powinna sie uruchamiac
            
            
                ' std looking by loop
                ' starting from beg every time!
                Set rm = managersDaSh.Cells(2, EVO.E_MANAGERS_DA_ARTICLE)
                Do
            
                
                    ' Debug.Assert rfa.Value <> "9832114080-02"
                    
                    ' Debug.Assert rm.row < 71
                    
                    
                    
                    If CStr(rm.Value) = strPn Then
                    
                    
                        ' =========================================================================================
                        ' =========================================================================================
                        rfa.offset(0, EVO.E_ADJUSTED_SQ01_RU - 1).Value = _
                            rm.offset(0, EVO.E_MANAGERS_DA_RU - EVO.E_MANAGERS_DA_ARTICLE).Value
                            
                        rfa.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_DA - 1).Value = _
                            rm.offset(0, EVO.E_MANAGERS_DA_ACHETEUR - EVO.E_MANAGERS_DA_ARTICLE).Value
                            
                            
                        'new 092
                        rfa.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_FAMILY - 1).Value = _
                            rm.offset(0, EVO.E_MANAGERS_DA_FAMILY - EVO.E_MANAGERS_DA_ARTICLE).Value
                            
                        rfa.offset(0, EVO.E_ADJUSTED_SQ01_MANAGER_GROUP - 1).Value = _
                            rm.offset(0, EVO.E_MANAGERS_DA_GROUP - EVO.E_MANAGERS_DA_ARTICLE).Value
                        ' =========================================================================================
                        ' =========================================================================================
                        
                        
                    End If
                    Set rm = rm.offset(1, 0)
                    
                    ' Application.Calculation = xlCalculationManual
                    
                Loop Until Trim(rm.Value) = ""
            End If
        
        
        End If
        
        
        
        ' offset dla adjusted sh
        Set rfa = rfa.offset(1, 0)
    Loop Until Trim(rfa.Value) = ""
    
    
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub


Public Sub testForFillingManagersDaForGreenLightAdjustedWorkhsheet()
    
    fillGreenLightManagersDaColumn ThisWorkbook.Sheets("TP04_20201007_I"), ThisWorkbook.Sheets("MANAGERS_DA_20201007_I")
End Sub







Public Sub matchSq01WithDAManagers(ictrl As IRibbonControl)
    Debug.Print "matchSq01With da managers"
    
    fillFormDaManagers "SQ01"
End Sub


Public Sub matchMb51WithDaManagers(ictrl As IRibbonControl)
    Debug.Print "matchMb51With da managers"
    
    fillFormDaManagers "MB51"
End Sub

Private Sub fillFormDaManagers(typ As String)
    
    FindTangoOrIntSData.ComboBox1.Clear
    ' FindTangoOrIntSData.Caption = "Match with Internal Suppliers"
    
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.name Like "MANAGERS_DA_*" Then
            FindTangoOrIntSData.ComboBox1.addItem sh.name
        End If
    Next sh
    
    FindTangoOrIntSData.Caption = "MANAGERS_DA_"
    FindTangoOrIntSData.typ = CStr(typ)
    FindTangoOrIntSData.show
End Sub
    
    

