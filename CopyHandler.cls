VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CopyHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cofors As Dictionary
Private linesIter As Dictionary
Private pns As Dictionary

' for optimisation logic - init when dwa = 2
Private pns2 As Dictionary


Private master As Worksheet
Private feed As Worksheet

Private Sub Class_Initialize()
    Set cofors = New Dictionary
    Set pns2 = Nothing
    
    Set master = Nothing
    Set feed = Nothing
End Sub

Private Sub Class_Terminate()
    Set cofors = Nothing
    Set pns2 = Nothing
    
    Set master = Nothing
    Set feed = Nothing
End Sub


Public Sub optimise(sh As StatusHandler)


    Dim li As LineItem
    Dim byTMC As Dictionary
    Dim innerDicForByTMC As Dictionary

    
    If Not pns2 Is Nothing Then
        If pns2.Count > 0 Then
        
            sh.init_statusbar (cofors.Count + 1)
            sh.show
            sh.progress_increase
            
            For Each k1 In pns2.Keys
            
            
                Dim sourceLi As LineItem
                Dim minDate As Date
                Dim minWiersz As Long
                Dim poczatek As Boolean
                poczatek = True
            
                Set byTMC = New Dictionary
                Set linesIter = pns2(k1)
                
                For Each k2 In linesIter.Keys
                
                    Set li = linesIter(k2)
                    tmcKey = li.tmc
                
                    If Not byTMC.Exists(tmcKey) Then
                        Set innerDicForByTMC = New Dictionary
                        innerDicForByTMC.Add li.wiersz, li
                        byTMC.Add tmcKey, innerDicForByTMC
                    Else
                        
                        Set innerDicForByTMC = byTMC(tmcKey)
                        innerDicForByTMC.Add li.wiersz, li
                    End If
                Next
                
                Debug.Print "count TMCs: " & byTMC.Count & " for PN: " & k1 & " wiersz: " & li.wiersz
                
                
                ' Debug.Assert li.wiersz <> 11
                
                
                For Each k3 In byTMC.Keys
                
                
                    Set innerDicForByTMC = byTMC(k3)
                    
                    poczatek = True
                    minWiersz = 0
                    
                    
                    If innerDicForByTMC.Count > 1 Then
                    
                        
                        
                            For Each k4 In innerDicForByTMC.Keys
                            
                            
                                Set li = innerDicForByTMC(k4)
                                
                                If li.tmc = EVO.G_NON_TMC Then
                                
                                    ' ------------------------------------------
                                    ' if tmc is non then no action is required!
                                    ' ------------------------------------------
                                    
                                    
                                Else
                                
                                    If IsDate(master.Cells(k4, EVO.G_DHEF_COL).Value) Then
                            
                                        If poczatek Then
                                            ' DHAS always assign
                                            ' this obj is empty working only with prev logic
                                            ' minDate = li.dateAfterOffset
                                            minDate = master.Cells(k4, EVO.G_DHEF_COL).Value
                                            minWiersz = k4
                                            poczatek = False
                                        Else
                                            
                                            
                                            If minDate > master.Cells(k4, EVO.G_DHEF_COL).Value Then
                                                
                                                minDate = master.Cells(k4, EVO.G_DHEF_COL).Value
                                                minWiersz = k4
                                            End If
                                        End If
                                    
                                    Else
                                        master.Cells(k4, EVO.G_DHEF_COL).Value = "ERR: not a date!"
                                        ' no operation
                                    End If
                                End If
                                
                            Next
                            
                            Set sourceLi = Nothing
                            On Error Resume Next
                            Set sourceLi = innerDicForByTMC(minWiersz)
                            
                            
                            Debug.Print "sourceLi - minWiersz &" _
                                ; " : " & master.Cells(minWiersz, EVO.G_DHEF_COL).Value
                            
                            For Each k4 In innerDicForByTMC.Keys
                            
                            
                                Set li = innerDicForByTMC(k4)
                                
                                
                                ' DHAS changes date time according to TMC
                                ' ---------------------------------------------------------------------
                                If master.Cells(k4, EVO.G_DHEF_COL).Comment Is Nothing Then
                                    master.Cells(k4, EVO.G_DHEF_COL).AddComment li.tmc & " PREV: " & _
                                        CStr(master.Cells(k4, EVO.G_DHEF_COL).Value)
                                Else
                                    master.Cells(k4, EVO.G_DHEF_COL).Comment.Text li.tmc & " PREV: " & _
                                        CStr(master.Cells(k4, EVO.G_DHEF_COL).Value)
                                End If
                                
                                
                                master.Cells(k4, EVO.G_DHEF_COL).Value = _
                                    master.Cells(minWiersz, EVO.G_DHEF_COL).Value
                                
                                ' ---------------------------------------------------------------------
                                
                                
                                ' DHAS changes date time according to TMC
                                ' ---------------------------------------------------------------------
                                If master.Cells(k4, EVO.G_DHAS_COL).Comment Is Nothing Then
                                    master.Cells(k4, EVO.G_DHAS_COL).AddComment li.tmc & " PREV: " & _
                                        CStr(master.Cells(k4, EVO.G_DHAS_COL).Value)
                                Else
                                    master.Cells(k4, EVO.G_DHAS_COL).Comment.Text li.tmc & " PREV: " & _
                                        CStr(master.Cells(k4, EVO.G_DHAS_COL).Value)
                                End If
                                
                                
                                master.Cells(k4, EVO.G_DHAS_COL).Value = _
                                    master.Cells(minWiersz, EVO.G_DHAS_COL).Value
                                    
                                ' ---------------------------------------------------------------------
                                
                                If CDbl(master.Cells(k4, EVO.G_DHAS_COL).Value) < _
                                    CDbl(master.Cells(k4, EVO.G_DHEF_COL).Value) Then
                                    
                                        master.Cells(k4, EVO.G_DHAS_COL).Value = _
                                            master.Cells(k4, EVO.G_DHEF_COL).Value
                                            
                                        master.Cells(k4, EVO.G_DHAS_COL).Comment.Text "timline fix " & Chr(10), 1, False
                                End If
                                
                            
                            Next
                    End If
                Next
                
            
                sh.progress_increase
            Next
            
            
            sh.hide
        End If
    End If

End Sub

Public Sub workWithData(sh As StatusHandler)


    
    
    If Not cofors Is Nothing Then
        If cofors.Count > 0 Then
        
            sh.init_statusbar (cofors.Count + 1)
            sh.show
            sh.progress_increase
            
            Set pns = New Dictionary
        
            
            If Not feed Is Nothing Then
            
            
                Dim pn As String, k1 As Variant, k2 As Variant
                Dim li As LineItem
                
                Dim srcRng As Range, f As Range
                Set srcRng = feed.Cells(1, 1)
                Set srcRng = srcRng.EntireColumn
            
                For Each k1 In cofors.Keys
                
                    pn = ""
                
                    Set linesIter = cofors(k1)
                    
                    For Each k2 In linesIter
                        
                        Set li = Nothing
                        
                        On Error Resume Next
                        Set li = linesIter(k2)
                        
                        pn = removeIndice(CStr(li.pn))
                        
                        
                        Set f = Nothing
                        
                        If pns.Exists(pn) Then
                            Set f = pns(pn)
                        Else
                        
                        
                            On Error Resume Next
                            Set f = srcRng.Find(pn)
                            
                            pns.Add pn, f
                            
                        End If
                        
                        If f Is Nothing Then
                            li.found = False
                            
                            ' no data in input!
                            ' put red mark on part number!
                            ' ======================================================
                            master.Cells(li.wiersz, 1).Font.Bold = True
                            master.Cells(li.wiersz, 1).Font.Color = RGB(255, 0, 0)
                            ' ======================================================
                            
                            
                        Else
                            li.found = True
                            Set li.sourceRng = f
                            li.sourceRow = f.Row
                            
                            
                            ' check if past
                            ' ================================================================
                            If li.monday < Date Then
                                master.Cells(li.wiersz, EVO.E_MASTER_ECHANCIER_ONL_S).Font.Bold = True
                                master.Cells(li.wiersz, EVO.E_MASTER_ECHANCIER_ONL_S).Font.Color = RGB(255, 255, 255)
                                master.Cells(li.wiersz, EVO.E_MASTER_ECHANCIER_ONL_S).Interior.Color = RGB(160, 160, 160)
                            End If
                            ' ================================================================
                            
                            
                            ' main coping!!
                            ' ================================================================
                            ''
                            '
                            mainCopyBasedOnRegisterData li, linesIter
                            '
                            ''
                            ' ================================================================
                            
                            
                        End If
                        
                    Next
                    
                    sh.progress_increase
                    
                Next
            
            End If
            
            
            sh.hide
        End If
    End If
End Sub


Private Sub mainCopyBasedOnRegisterData(li As LineItem, ByRef linesIter As Dictionary)
    
    ' main iteration logic
    ' ---------------------------------------------------
    ''
    '
    
    ' Debug.Print "working on line : " & li.wiersz
    Dim kolumna1 As Integer, kolumna2 As Variant
    Dim od As Range, msh As Worksheet, lsh As Worksheet
    
    Dim optWiersz As Long, optSrcWiersz As Long
    optWiersz = li.wiersz
    optSrcWiersz = li.sourceRow
    
    Set od = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("D2")
    Set msh = li.masterRng.Parent
    Set lsh = li.sourceRng.Parent
    
    Do
        
        kolumna1 = od.Value
        kolumna2 = od.Offset(0, 3).Value
        
        'msh.Cells(optWiersz, 1).Interior.Color = RGB(255, 255, 255)
        'msh.Cells(optWiersz, 1).Font.Color = RGB(0, 0, 0)
        
    
        If IsNumeric(od.Offset(0, 3).Value) Then
            
            'simple copy without X
            If Trim(od.Offset(0, 5).Value) = "" Then
            
                If IsError(lsh.Cells(optSrcWiersz, kolumna2).Value) Then
                    ' msh.Cells(optWiersz, kolumna1).Value = "error n/a in source file"
                    msh.Cells(optWiersz, kolumna1).Interior.Color = RGB(255, 200, 200)
                   ' msh.Cells(optWiersz, 1).Interior.Color = RGB(255, 200, 200)
                Else
                
                
                    ' same data - so blue
                    If msh.Cells(optWiersz, kolumna1).Value = lsh.Cells(optSrcWiersz, kolumna2).Value Then
                        ' msh.Cells(optWiersz, kolumna1).Value = lsh.Cells(optSrcWiersz, kolumna2).Value
                        msh.Cells(optWiersz, kolumna1).Interior.Color = RGB(200, 200, 255)
                        'msh.Cells(optWiersz, 1).Interior.Color = RGB(200, 200, 255)
                        
                        If Not msh.Cells(optWiersz, kolumna1).Comment Is Nothing Then
                            msh.Cells(optWiersz, kolumna1).Comment.Text CStr(Date) & " : " & CStr(tmpTxt) & Chr(10), _
                                1, False
                                
                            msh.Cells(optWiersz, kolumna1).Comment.Shape.TextFrame.AutoSize = True
                            
                        End If
                        
                        
                    ' pus cell was empty so filling by source and some green
                    ElseIf Trim(msh.Cells(optWiersz, kolumna1).Value) = "" Then
                        msh.Cells(optWiersz, kolumna1).Value = lsh.Cells(optSrcWiersz, kolumna2).Value
                        msh.Cells(optWiersz, kolumna1).Interior.Color = RGB(200, 255, 200)
                        'msh.Cells(optWiersz, 1).Interior.Color = RGB(200, 255, 200)
                        
                    ' diff data - yellow color
                    ElseIf msh.Cells(optWiersz, kolumna1).Value <> lsh.Cells(optSrcWiersz, kolumna2).Value Then
                    
                        ' tmpTxt = CStr(msh.Cells(optWiersz, kolumna1).Value)
                        tmpTxt = CStr(lsh.Cells(optSrcWiersz, kolumna2).Value)
                        
                        'msh.Cells(optWiersz, kolumna1).Value = lsh.Cells(optSrcWiersz, kolumna2).Value
                        msh.Cells(optWiersz, kolumna1).Interior.Color = RGB(255, 255, 150)
                        'msh.Cells(optWiersz, 1).Interior.Color = RGB(255, 255, 200)
                        
                        If msh.Cells(optWiersz, kolumna1).Comment Is Nothing Then
                            msh.Cells(optWiersz, kolumna1).AddComment CStr(Date) & " : " & CStr(tmpTxt) & Chr(10)
                            msh.Cells(optWiersz, kolumna1).Comment.Shape.TextFrame.AutoSize = True
                        Else
                            msh.Cells(optWiersz, kolumna1).Comment.Text CStr(Date) & " : " & CStr(tmpTxt) & Chr(10), _
                                1, False
                            
                            msh.Cells(optWiersz, kolumna1).Comment.Shape.TextFrame.AutoSize = True
                            
                        End If
                        
                        
                    End If
                End If
            ElseIf Trim(od.Offset(0, 5).Value) = "X" Then

                calculationForDH li, linesIter, Trim(od.Offset(0, 4).Value), lsh, msh, CLng(optWiersz)
                
            End If
        End If

    
        Set od = od.Offset(1, 0)
    Loop Until Trim(od) = ""
    
    '
    ''
    ' ---------------------------------------------------
End Sub


Private Sub calculationForDH(li As LineItem, ByRef linesIter As Dictionary, dhType As String, lsh As Worksheet, msh As Worksheet, optWiersz As Long)
    
    
    ' logic only on the beginning so DHEF calc all!
    If dhType = "DHEF" Then
        ' -----------------------------------------------------------------------
        ' -----------------------------------------------------------------------
        'Dim kolumnaCodu As Long
        'Dim kolumnaPuTime As Long

        ' UWAGA STATIC CODE!
        ' ------------------------------------------------------------------
        'kolumnaCodu = CLng(ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("B37"))
        'kolumnaPuTime = CLng(ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("B38"))
        ' kolumnaCodu
        

        ' ------------------------------------------------------------------
        
        ' li.CodEntrega = li.sourceRng.Parent.Cells(li.sourceRow, kolumnaCodu).Value
        li.CodEntrega = lsh.Cells(li.sourceRow, EVO.G_COD_TRANSPORT_COLUMN).Value
        
        
        If li.CodEntrega <> EVO.G_DAP Then
            li.dateAfterOffset = offsetDate(li.CodEntrega, li.monday)
            li.delTime = CDate(lsh.Cells(li.sourceRow, EVO.G_DEL_TIME_COLUMN).Value)
            li.delDate = li.dateAfterOffset + CDate(li.delTime)
            
            li.puDate = li.delDate - li.delTime - CDate(lsh.Cells(li.sourceRow, EVO.G_T_TIME_COLUMN).Value)
            li.puTime = CDate(lsh.Cells(li.sourceRow, EVO.G_PU_TIME_COLUMN).Value)
            li.puDate = li.puDate + li.puTime
            
            
            msh.Cells(optWiersz, G_DHEF_COL).Value = li.puDate
            msh.Cells(optWiersz, G_DHAS_COL).Value = li.delDate
            msh.Cells(optWiersz, G_DHAS_COL + 1).Formula = "=C" & optWiersz & "/Y" & optWiersz
        End If
        
        ' -----------------------------------------------------------------------
        ' -----------------------------------------------------------------------
    ElseIf dhType = "DHAS" Then
        '        ' -----------------------------------------------------------------------
        '        ' -----------------------------------------------------------------------
        '        'Dim kolumnaTT1 As Long, kolumnaTT2 As Long
        '        'Dim kolumnaDelTime As Long
        '
        '        'kolumnaTT1 = CLng(ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("B39"))
        '        'kolumnaTT2 = CLng(ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("B41"))
        '        'kolumnaDelTime = CLng(ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("B42"))
        '
        '        'li.delDate = li.dateAfterOffset _
        '        '    - li.puTime _
        '        '    + Int(li.sourceRng.Parent.Cells(li.sourceRow, kolumnaTT1).Value) _
        '        '    + Int(li.sourceRng.Parent.Cells(li.sourceRow, kolumnaTT2).Value)
        '
        '        li.delDate = li.dateAfterOffset _
        '            - li.puTime _
        '            + Int(lsh.Cells(li.sourceRow, 112).Value) _
        '            + Int(lsh.Cells(li.sourceRow, 114).Value)
        '
        '        On Error Resume Next
        '        'li.delDate = li.delDate + CDate(li.sourceRng.Parent.Cells(li.sourceRow, kolumnaDelTime).Value)
        '        li.delDate = li.delDate + CDate(lsh.Cells(li.sourceRow, 115).Value)
        '
        '        ' -----------------------------------------------------------------------
        '        ' -----------------------------------------------------------------------
    Else
        ' nop
    End If
End Sub

 
Private Function getMondayFromYYYYCW(yyyycw As Long) As Date
    
    getMondayFromYYYYCW = CDate(Format("1900-01-01", "yyyy-mm-dd"))
    
    
    Dim tmpDate As Date
    tmpDate = Date
    
    Dim currWeekday As Double
    currWeekday = Application.WorksheetFunction.Weekday(tmpDate, 2)
    tmpDate = tmpDate - currWeekday + 1
    
    currYear = Year(tmpDate)
    currCW = Application.WorksheetFunction.IsoWeekNum(tmpDate)
    
    Dim curr_yyyycw As Long
    curr_yyyycw = currYear * 100 + CLng(currCW)
    
    Dim diff As Long
    diff = yyyycw - curr_yyyycw
    
    ' to jest current monday: tmpDate
    tmpDate = tmpDate + 7 * diff
    
    getMondayFromYYYYCW = tmpDate
    
    
    
End Function



Private Function removeIndice(fullPn As String) As String
    
    arr = Split(fullPn, "-")
    
    removeIndice = arr(LBound(arr))
    
End Function


Public Sub init(m_master As Worksheet, m_feed As Worksheet, Optional dwa As Integer)


    Set master = m_master
    Set feed = m_feed
    
    
    If (Not master Is Nothing) And (Not feed Is Nothing) Then
        ' we are OK to go further
        
        ' no need to pass args cos master and feed are class fields!
        ' -------------------------------------------------------------
        ' init by wrong cofro - but that giving any hurt
        If dwa = 1 Then
            initAllDictionaries 1
        Else
            initAllDictionaries 2
        End If
        ' -------------------------------------------------------------
        
    Else
        ' NOK situation
        MsgBox "Sth went wrong!", vbCritical
        End
    End If

End Sub

Private Sub initAllDictionaries(dwa As Integer)

    Dim thisLine As LineItem
    Dim base As Worksheet
    Set base = master
    
    Dim r As Range
    Set r = base.Cells(3, EVO.E_MASTER_Reference)


    If dwa = 2 Then
        Dim strTempPn As String
        Set pns2 = New Dictionary
        Do
            strTempPn = Trim(r.Value)
            
            If Not pns2.Exists(strTempPn) Then
                Set thisLine = New LineItem
            Else
                With thisLine
                    .pn = strTempPn
                    .cofor = Trim(CStr(r.Offset(0, EVO.G_COFOR_VENDEUR_COL - 1).Value))
                    ' !!!
                    Set .masterRng = r
                    '
                    .wiersz = CLng(r.Row)
                    .tmc = Trim(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                    
                    .yyyycw = CLng(calculateYYYYCW(Trim(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                    
                    .monday = CDate(getMondayFromYYYYCW(CLng(.yyyycw)))
                    
                End With
                
                pns2.Add strTemp, thisLine
                
            End If
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""

    ElseIf dwa = 1 Then

        
        Dim strTempCofor As String
        Set cofors = New Dictionary

        Do
        
        
            strTempCofor = Trim(CStr(r.Offset(0, EVO.E_MASTER_cofor - 1).Value))
        
            If Not cofors.Exists(strTempCofor) Then
            
                Set linesIter = Nothing
                Set linesIter = New Dictionary
                
            
                Set thisLine = New LineItem
                
                With thisLine
                    .cofor = strTempCofor
                    .pn = Trim(CStr(r.Offset(0, EVO.E_MASTER_Reference - 1).Value))
                    ' !!!
                    Set .masterRng = r
                    '
                    .wiersz = CLng(r.Row)
                    .tmc = Trim(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                    
                    .yyyycw = CLng(calculateYYYYCW(Trim(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                    
                    .monday = CDate(getMondayFromYYYYCW(CLng(.yyyycw)))
                    
                End With
                
                linesIter.Add r.Row, thisLine
                
            
                cofors.Add strTempCofor, linesIter
            Else
            
                ' this cofor is listed - also it means that dictionary inside that cofor as well
                
                Set linesIter = cofors(strTempCofor)
                
                
                Set thisLine = New LineItem
                
                With thisLine
                    .cofor = strTempCofor
                    .pn = Trim(CStr(r.Offset(0, EVO.E_MASTER_Reference - 1).Value))
                    ' !!!
                    Set .masterRng = r
                    '
                    .wiersz = CLng(r.Row)
                    .tmc = Trim(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                    .yyyycw = CLng(calculateYYYYCW(Trim(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                    
                    .monday = CDate(getMondayFromYYYYCW(CLng(.yyyycw)))
                End With
                
                
                linesIter.Add r.Row, thisLine
                
                
            End If
        
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        
        
        
        ' afer all check math
        Debug.Print "count cofors: " & cofors.Count
        
        ' check all lines
        xl = 0
        For Each k In cofors.Keys
            
            Set linesIter = cofors(k)
            
            Debug.Print "for cofor: " & k & " we have: " & linesIter.Count & " lines"
            xl = xl + linesIter.Count
            
        Next
        
        ' -1 for loop -2 for labels
        Debug.Print "checksum on all lines: " & xl & " rows in file: " & (r.Row - 1 - 2)
    
        
        If CLng(xl) = CLng(r.Row - 1 - 2) Then
            '  OK
        Else
            MsgBox "Checksum on data is not right - there is sth not right with the BASE worksheet!", vbCritical
        End If
    End If
End Sub



Private Function calculateYYYYCW(s As String)
    
    If Left(s, 1) = "S" Then
    
        s = Mid(s, 2)
    
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
