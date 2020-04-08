VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CopyHandler"
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


Private cofors As Dictionary
Private expCofors As Dictionary
Private linesIter As Dictionary
Private pns As Dictionary

' for optimisation logic - init when dwa = 2
Private pns2 As Dictionary


Private master As Worksheet
Private feed As Worksheet

' dwa = 3
Private routeSheet As Worksheet
Private pleSheet As Worksheet
Private uaSheet As Worksheet



' delegacja priv methods
Private operacja As OperationsHandler
Private ci As CopyIterationForPivotSource


Private Sub Class_Initialize()
    Set cofors = New Dictionary
    Set expCofors = New Dictionary
    Set pns2 = Nothing
    
    Set master = Nothing
    Set feed = Nothing
    
    Set routeSheet = Nothing
    Set pleSheet = Nothing
    Set uaSheet = Nothing
    
    
    Set operacja = New OperationsHandler
    Set ci = New CopyIterationForPivotSource
End Sub

Private Sub Class_Terminate()
    Set cofors = Nothing
    Set expCofors = Nothing
    Set pns2 = Nothing
    
    Set master = Nothing
    Set feed = Nothing
    
    Set routeSheet = Nothing
    Set pleSheet = Nothing
    Set uaSheet = Nothing
    
    
    Set operacja = Nothing
    Set ci = Nothing
End Sub


Public Sub optimise(sh As StatusHandler)


    Dim li As LineItem
    Dim byTMC As Dictionary
    Dim innerDicForByTMC As Dictionary
    Dim k1 As Variant, k2 As Variant, k3 As Variant, k4 As Variant
    Dim tmcKey As String

    
    If Not pns2 Is Nothing Then
        If pns2.Count > 0 Then
        
            sh.init_statusbar (pns2.Count + 1)
            sh.show
            sh.progress_increase
            
            
            For Each k1 In pns2.Keys
            
            
                Dim sourceLi As LineItem
                Dim minDate As Date
                Dim minWiersz As Long
                Dim backupWiersz As Long
                Dim poczatek As Boolean
                poczatek = True
                
                ' brak ven... coforu zatem nie ma co liczyc!
                ' -------------------------------------------------
                If k1 Like "*X" Then
                    ' nop req
                Else
            
                    Set byTMC = New Dictionary
                    Set linesIter = pns2(k1)
                    
                    For Each k2 In linesIter.Keys
                    
                        Set li = linesIter(k2)
                        
                        
                        ' Debug.Assert li.wiersz <> 250
                        
                        ' Debug.Print "we are in key: " & k1 & " " & li.pn & " " & li.cofor & " " & li.masterRng.Row
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
                    
                    Debug.Print "count TMCs: " & tmcKey & " " & byTMC.Count & " for PN: " & k1 & " wiersz: " & li.wiersz
                    
                    
                    ' Debug.Assert li.wiersz <> 152
                    
                    
                    minDate = Date
                    
                    
                    For Each k3 In byTMC.Keys
                    
                        
                        Set innerDicForByTMC = byTMC(k3)
                        
                        
                        If k3 = EVO.G_NON_TMC Then
                        
                        
                            ' non TMC ' ----------------------------------------------------------
                            ' ---------------------------------------------------------------------------------------
                            ' ---------------------------------------------------------------------------------------
                            ' ---------------------------------------------------------------------------------------
                            iterateAndAdjustDatesOnDHEFandDHAS innerDicForByTMC, 0
                            ' ---------------------------------------------------------------------------------------
                            ' ---------------------------------------------------------------------------------------
                            ' ---------------------------------------------------------------------------------------
                            
                        ElseIf k3 <> EVO.G_NON_TMC Then
                        
                            poczatek = True
                            minWiersz = 0
                            backupWiersz = 0
                            
                            If innerDicForByTMC.Count = 1 Then
                            
                                For Each k4 In innerDicForByTMC.Keys
                                    minWiersz = k4
                                Next
                                
                            ElseIf innerDicForByTMC.Count > 1 Then
                            
                                
                                
                                    For Each k4 In innerDicForByTMC.Keys
                                    
                                        backupWiersz = k4
                                        
                                        ' Debug.Assert backupWiersz <> 168
                                    
                                    
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
                                                    ' Debug.Print li.puDate ' now it is OK!
                                                    
                                                    If li.isPast Then
                                                        ' nop
                                                    Else
                                                        minDate = master.Cells(k4, EVO.G_DHEF_COL).Value
                                                        minWiersz = k4
                                                        poczatek = False
                                                    End If
                                                Else
                                                    
                                                    If li.isPast Then
                                                        ' nop
                                                    Else
                                                        If minDate > master.Cells(k4, EVO.G_DHEF_COL).Value Then
                                                            
                                                            minDate = master.Cells(k4, EVO.G_DHEF_COL).Value
                                                            minWiersz = k4
                                                        End If
                                                    End If
                                                End If
                                            
                                            Else
                                                master.Cells(k4, EVO.G_DHEF_COL).Value = "ERR: not a date!"
                                                ' no operation
                                            End If
                                        End If
                                        
                                    Next
                            End If
                                    
                            Set sourceLi = Nothing
                            
                            If minWiersz <> 0 Then
                            
                            
                                On Error Resume Next
                                Set sourceLi = innerDicForByTMC(minWiersz)
                            
                                Debug.Print "sourceLi - minWiersz &" _
                                    ; " : " & master.Cells(minWiersz, EVO.G_DHEF_COL).Value
                            Else
                                
                                Debug.Print "all data in the past! or there was only data once!"
                                Debug.Print li.yyyycw & " " & CStr(Application.WorksheetFunction.IsoWeekNum(Date))
                            End If
                            
                            
                            ' ---------------------------------------------------------------------------------------
                            ' ---------------------------------------------------------------------------------------
                            ' ---------------------------------------------------------------------------------------
                            ' ---------------------------------------------------------------------------------------
                            iterateAndAdjustDatesOnDHEFandDHAS innerDicForByTMC, minWiersz
                            ' ---------------------------------------------------------------------------------------
                            ' ---------------------------------------------------------------------------------------
                            ' ---------------------------------------------------------------------------------------
                            ' ---------------------------------------------------------------------------------------
                            
                        ' TMC <> 'non' END
                        End If
                    Next
                    
                End If
                    
                
                sh.progress_increase
            Next
                
            
            
            sh.hide
        End If
    End If

End Sub



Private Sub iterateAndAdjustDatesOnDHEFandDHAS(innerDicForByTMC As Dictionary, minWiersz As Long)


    Dim li As LineItem, k4 As Variant
    Dim cw1 As Long, cw2 As Long, diff As Long
    
    For Each k4 In innerDicForByTMC.Keys
    
    
        Set li = innerDicForByTMC(k4)
        
        ' Debug.Assert k4 <> 152
        
        
        
        ' DHEF changes date time according to TMC
        ' ---------------------------------------------------------------------
        If master.Cells(k4, EVO.G_DHEF_COL).Comment Is Nothing Then
            master.Cells(k4, EVO.G_DHEF_COL).AddComment li.tmc & " PREV: " & _
                CStr(master.Cells(k4, EVO.G_DHEF_COL).Value)
                
            master.Cells(k4, EVO.G_DHEF_COL).Comment.Shape.TextFrame.AutoSize = True
        Else
            master.Cells(k4, EVO.G_DHEF_COL).Comment.Text li.tmc & " PREV: " & _
                CStr(master.Cells(k4, EVO.G_DHEF_COL).Value), 1, False
                
            master.Cells(k4, EVO.G_DHEF_COL).Comment.Shape.TextFrame.AutoSize = True
        End If
        
        
        If minWiersz <> 0 Then
        
            If minWiersz = k4 And li.isPast Then
            
                ' move to current week!
                cw1 = CLng(CLng(Year(li.puDate)) * 100 + CLng(Application.WorksheetFunction.IsoWeekNum(li.puDate)))
                cw2 = CLng(CLng(Year(Date)) * 100 + CLng(Application.WorksheetFunction.IsoWeekNum(Date)))
                
                diff = cw1 - cw2
                    
                master.Cells(minWiersz, EVO.G_DHEF_COL).Comment.Text "date moved " & CStr(diff) & " weeks! " & _
                    "old date: " & CStr(master.Cells(minWiersz, EVO.G_DHEF_COL).Value)
                    
                master.Cells(minWiersz, EVO.G_DHEF_COL).Comment.Shape.TextFrame.AutoSize = True
                    
                master.Cells(minWiersz, EVO.G_DHEF_COL).Font.Color = RGB(240, 100, 0)
                    
                ' uwaga negative logic!
                master.Cells(minWiersz, EVO.G_DHEF_COL).Value = CDate(master.Cells(minWiersz, EVO.G_DHEF_COL).Value) - (diff * 7)
            
            
            End If
        
        
            master.Cells(k4, EVO.G_DHEF_COL).Value = _
                master.Cells(minWiersz, EVO.G_DHEF_COL).Value
                
            If li.isPast Then
                master.Cells(k4, EVO.G_DHEF_COL).Font.Color = RGB(240, 0, 0)
            End If
        Else
        
            If li.isPast Then
        
                ' move to current week!
                cw1 = CLng(CLng(Year(li.puDate)) * 100 + CLng(Application.WorksheetFunction.IsoWeekNum(li.puDate)))
                cw2 = CLng(CLng(Year(Date)) * 100 + CLng(Application.WorksheetFunction.IsoWeekNum(Date)))
                
                diff = cw1 - cw2
                    
                master.Cells(k4, EVO.G_DHEF_COL).Comment.Text "date moved " & CStr(diff) & " weeks! " & _
                    "old date: " & CStr(master.Cells(k4, EVO.G_DHEF_COL).Value)
                    
                master.Cells(k4, EVO.G_DHEF_COL).Comment.Shape.TextFrame.AutoSize = True
                    
                master.Cells(k4, EVO.G_DHEF_COL).Font.Color = RGB(240, 100, 100)
                    
                ' uwaga negative logic!
                master.Cells(k4, EVO.G_DHEF_COL).Value = CDate(master.Cells(k4, EVO.G_DHEF_COL).Value) - (diff * 7)
                
            End If
                
            
        End If
        
        ' ---------------------------------------------------------------------
        
        
        ' DHAS changes date time according to TMC
        ' ---------------------------------------------------------------------
        If master.Cells(k4, EVO.G_DHAS_COL).Comment Is Nothing Then
            master.Cells(k4, EVO.G_DHAS_COL).AddComment li.tmc & " PREV: " & _
                CStr(master.Cells(k4, EVO.G_DHAS_COL).Value)
                
            master.Cells(k4, EVO.G_DHAS_COL).Comment.Shape.TextFrame.AutoSize = True
        Else
            master.Cells(k4, EVO.G_DHAS_COL).Comment.Text li.tmc & " PREV: " & _
                CStr(master.Cells(k4, EVO.G_DHAS_COL).Value)
                
            master.Cells(k4, EVO.G_DHAS_COL).Comment.Shape.TextFrame.AutoSize = True
        End If
        
        If minWiersz <> 0 Then
        
        
            If minWiersz = k4 And li.isPast Then
            
                ' move to current week!
                cw1 = CLng(CLng(Year(li.puDate)) * 100 + CLng(Application.WorksheetFunction.IsoWeekNum(li.puDate)))
                cw2 = CLng(CLng(Year(Date)) * 100 + CLng(Application.WorksheetFunction.IsoWeekNum(Date)))
                
                diff = cw1 - cw2
                    
                master.Cells(minWiersz, EVO.G_DHAS_COL).Comment.Text "date moved " & CStr(diff) & " weeks! " & _
                    "old date: " & CStr(master.Cells(minWiersz, EVO.G_DHAS_COL).Value)
                    
                master.Cells(minWiersz, EVO.G_DHAS_COL).Comment.Shape.TextFrame.AutoSize = True
                    
                master.Cells(minWiersz, EVO.G_DHAS_COL).Font.Color = RGB(240, 100, 100)
                    
                ' uwaga negative logic!
                master.Cells(minWiersz, EVO.G_DHAS_COL).Value = CDate(master.Cells(minWiersz, EVO.G_DHAS_COL).Value) - _
                    (7 * diff)
            
            
            End If
        
        
            master.Cells(k4, EVO.G_DHAS_COL).Value = _
                master.Cells(minWiersz, EVO.G_DHAS_COL).Value
                
            If li.isPast Then
                master.Cells(k4, EVO.G_DHAS_COL).Font.Color = RGB(240, 0, 0)
            End If
        Else
        
        
            If li.isPast Then
        
                ' move to current week!
                cw1 = CLng(CLng(Year(li.puDate)) * 100 + CLng(Application.WorksheetFunction.IsoWeekNum(li.puDate)))
                cw2 = CLng(CLng(Year(Date)) * 100 + CLng(Application.WorksheetFunction.IsoWeekNum(Date)))
                
                diff = cw1 - cw2
                
                    
                master.Cells(k4, EVO.G_DHAS_COL).Comment.Text "date moved " & CStr(diff) & " weeks! " & _
                    "old date: " & CStr(master.Cells(k4, EVO.G_DHAS_COL).Value)
                
                master.Cells(k4, EVO.G_DHAS_COL).Comment.Shape.TextFrame.AutoSize = True
                    
                master.Cells(k4, EVO.G_DHAS_COL).Font.Color = RGB(240, 0, 0)
                    
                
                ' uwaga negative logic!
                master.Cells(k4, EVO.G_DHAS_COL).Value = CDate(master.Cells(k4, EVO.G_DHAS_COL).Value) - _
                    (7 * diff)
                
                
            End If
        End If
        ' ---------------------------------------------------------------------
        
        
        
        ' this action required if is not chronological!
        ' ------------------------------------------------------------------------------------------
        ' ------------------------------------------------------------------------------------------
        If CDbl(master.Cells(k4, EVO.G_DHAS_COL).Value) < _
            CDbl(master.Cells(k4, EVO.G_DHEF_COL).Value) Then
            
                master.Cells(k4, EVO.G_DHAS_COL).Value = _
                    master.Cells(k4, EVO.G_DHEF_COL).Value
                    
                master.Cells(k4, EVO.G_DHAS_COL).Comment.Text "timline fix " & Chr(10), 1, False
        End If
        ' ------------------------------------------------------------------------------------------
        ' ------------------------------------------------------------------------------------------
        
    
    Next

End Sub

Public Sub copyForSourcePivot(ByRef ph As PivotHandler, sh As StatusHandler)


    Dim condi As String
    Dim expCofor As String
    
    If Not expCofors Is Nothing Then
        If expCofors.Count > 0 Then
        
            sh.init_statusbar (expCofors.Count + 1)
            sh.show
            sh.progress_increase
            
            
            Dim quickFeedback As String
            quickFeedback = ""
            If Not feed Is Nothing Then
                If Not pleSheet Is Nothing Then
                    If Not uaSheet Is Nothing Then
                        If Not routeSheet Is Nothing Then
                        
                        
                            ' beofre main logic
                            ' init some sheet for source pivot
                            ' ---------------------------------------------------
                            ' !
                            ph.initPivotSource
                            ph.fillPivotSourceWithLabels
                            ph.mountStartingPoint
                            ' !
                        
                            ' main logic
                            ' ---------------------------------------------------
                            
                            
                            
                            Dim k1 As Variant, k2 As Variant
                            Dim li As LineItem
                            
                            
                            
                            Dim srcRng As Range, pleRng As Range, cloeRng As Range, condiRng As Range
                            Dim findByExpCoforInPle As Range
                            Dim findByExpCoforInCloe As Range
                            Dim findByExpCoforInMain As Range
                            Dim findByCondiInUa As Range
                            
                            Set srcRng = feed.Cells(1, EVO.G_FEED_MAIN_SH_SHIPPER_COFOR)
                            Set srcRng = srcRng.EntireColumn
                            
                            Set pleRng = pleSheet.Cells(1, EVO.G_PLE_SHIPPER_COFOR)
                            Set pleRng = pleRng.EntireColumn
                            
                            Set cloeRng = routeSheet.Cells(1, EVO.G_CLOE_SHIPPER_COFOR)
                            Set cloeRng = cloeRng.EntireColumn
                            
                            Set condiRng = uaSheet.Cells(1, 1)
                            Set condiRng = condiRng.EntireColumn
                            
                            
                        
                            For Each k1 In expCofors.Keys
                            
                                expCofor = ""
                                condi = ""
                                
                                Set linesIter = expCofors(k1)
                                
                                For Each k2 In linesIter
                                    
                                    Set li = Nothing
                                    
                                    On Error Resume Next
                                    Set li = linesIter(k2)
                                    
                                    expCofor = li.cofor
                                    condi = li.condi
                                    
                                    
                                    ' Debug.Assert li.wiersz <> 2866
                                    
                                    
                                    
                                    ' findByPartNumber
                                    Set findByExpCoforInPle = Nothing
                                    Set findByExpCoforInCloe = Nothing
                                    Set findByExpCoforInMain = Nothing
                                    Set findByCondiInUa = Nothing
                                    
                                    On Error Resume Next
                                    Set findByExpCoforInPle = pleRng.Find(expCofor)
                                    On Error Resume Next
                                    Set findByExpCoforInCloe = cloeRng.Find(expCofor)
                                    On Error Resume Next
                                    Set findByExpCoforInMain = srcRng.Find(expCofor)
                                    On Error Resume Next
                                    Set findByCondiInUa = condiRng.Find(condi)
                                    
                                    
                                    If operacja.validateSearching(findByExpCoforInPle, _
                                        findByExpCoforInCloe, _
                                        findByExpCoforInMain, _
                                        findByCondiInUa) Then
                                    
                                            li.validationForPivotSource = True
                                            Set li.supplementForPivotSource = New LineItemPivotSrouceSupplement
                                            
                                            With li.supplementForPivotSource
                                                Set .refRangeForCloe = findByExpCoforInCloe
                                                Set .refRangeForPLE = findByExpCoforInPle
                                                Set .refRaneForCondi = findByCondiInUa
                                            End With
                                            
                                            Set li.sourceRng = findByExpCoforInMain
                                            On Error Resume Next
                                            li.sourceRow = findByExpCoforInMain.Row
                                                
                                            ' main coping for pivot source!!
                                            ' ================================================================
                                            ''
                                            '
                                            insideCopyForSourcePivot_setupOrderDate li, linesIter, pleSheet
                                            mainCopyForPivotSource li, linesIter, ph
                                            '
                                            ''
                                            ' ================================================================
                                    Else
                                        li.validationForPivotSource = False
                                    End If
                                    
                                    
                                Next
                                
                                
                                sh.progress_increase
                                
                            Next
                            
                            
                            ' ---------------------------------------------------
                        Else
                            quickFeedback = quickFeedback & " " & " no data from CLOE provided!"
                        End If
                    Else
                        quickFeedback = quickFeedback & " " & " no data from UA at all!"
                    End If
                Else
                    quickFeedback = quickFeedback & " no data from PLE!"
                End If
            Else
                quickFeedback = quickFeedback & " where is you main table in control tower database?"
            End If
            
            If quickFeedback <> "" Then
                MsgBox quickFeedback, vbCritical
            End If
            
            
            
            ' array formulas section!
            ' ===============================================================================
            
            operacja.makeMyFormulas ph.getPivotSource
            Set ph.proxy2 = operacja.makeLeanCopyOfPivotSource(ph.getPivotSource)
            ph.proxy2.Name = ph.tryToRenameProxy2Worksheet(ph.proxy2)
            
            ' ===============================================================================
            
            
            ' HIDE!
            sh.hide
            
            
        End If
    End If
End Sub




Public Sub copyData(sh As StatusHandler)


    
    
    If Not cofors Is Nothing Then
        If cofors.Count > 0 Then
        
            sh.init_statusbar (cofors.Count + 1)
            sh.show
            sh.progress_increase
            
            Set pns = New Dictionary
        
            
            If Not feed Is Nothing Then
            
            
            
                Dim ycwMOnday As Long
                Dim ycwToday As Long
            
            
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
                        
                        pn = operacja.removeIndice(CStr(li.pn))
                        
                        
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
                            
                            ycwMOnday = CLng(Year(li.monday)) * 100 + Application.WorksheetFunction.IsoWeekNum(li.monday)
                            ycwToday = CLng(Year(Date)) * 100 + Application.WorksheetFunction.IsoWeekNum(Date)
                            
                            If ycwMOnday < ycwToday Then
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


Private Sub insideCopyForSourcePivot_setupOrderDate(li As LineItem, ByRef linesIter As Dictionary, pleSh As Worksheet)
    Dim tmpS As LineItemPivotSrouceSupplement
    Set tmpS = li.supplementForPivotSource
    
    On Error Resume Next
    tmpS.orderDate = calculateOrderDate(li, pleSh)
End Sub

Private Function calculateOrderDate(li As LineItem, pleSh As Worksheet) As Date


    Dim cellWithProperValue As Range
    Set cellWithProperValue = _
        pleSh.Cells(li.supplementForPivotSource.refRangeForPLE.Row, EVO.G_PLE_SUB_FOR_ORDER_COLUMN)

    Dim howManyDaysOfOffset As Integer
    
    If IsNumeric(cellWithProperValue) Then
        howManyDaysOfOffset = Int(cellWithProperValue.Value)
    Else
        howManyDaysOfOffset = -1
    End If
    
    
    
    If howManyDaysOfOffset <> -1 Then
    
        Dim tmpdate As Date
        On Error Resume Next
        tmpdate = CDate(li.delDate) - howManyDaysOfOffset
        
        li.supplementForPivotSource.heoDecalageJX = Int(howManyDaysOfOffset)
        calculateOrderDate = CDate(tmpdate)
    End If
End Function


Private Sub mainCopyForPivotSource(li As LineItem, ByRef linesIter As Dictionary, ph As PivotHandler)


    If checkLineItemIfItMakeAnySense(li) Then
    
    
        
        
        ' there is data!
        'Debug.Print "checkLineItemIfItMakeAnySense = true " & li.wiersz
        
        Dim sh As Worksheet, linia As Long, linia2 As Long
        Set sh = ph.getPivotSource()
        linia2 = li.wiersz
        linia = ph.getRefRangeInSrcSheet().Row
        ' ------------------------------------------------------
        ci.D_copyIterationForPivotSource master, sh, li, ph, linia
        ' ------------------------------------------------------
        ph.moveRefRangeInSrcSheet
        
    Else
        ' no proper data from source sheets
        Debug.Print "checkLineItemIfItMakeAnySense = false"
    End If

End Sub


Private Function checkLineItemIfItMakeAnySense(li As LineItem) As Boolean
    checkLineItemIfItMakeAnySense = False
    
    If Not li.supplementForPivotSource.refRangeForCloe Is Nothing Then
        If Not li.supplementForPivotSource.refRangeForPLE Is Nothing Then
            If Not li.sourceRng Is Nothing Then
                checkLineItemIfItMakeAnySense = True
            End If
        End If
    End If
End Function


Private Sub mainCopyBasedOnRegisterData(li As LineItem, ByRef linesIter As Dictionary)
    
    ' main iteration logic
    ' ---------------------------------------------------
    ''
    '
    
    Dim tmpTxt As String
    
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
        
        
        ' Debug.Assert li.wiersz <> 3628
        
        
        If li.CodEntrega <> EVO.G_DAP Then
            li.dateAfterOffset = offsetDate(li.CodEntrega, li.monday)
            li.delTime = CDate(lsh.Cells(li.sourceRow, EVO.G_DEL_TIME_COLUMN).Value)
            li.delDate = li.dateAfterOffset + CDate(li.delTime)
            
            li.puDate = li.delDate - li.delTime - CDate(lsh.Cells(li.sourceRow, EVO.G_T_TIME_COLUMN).Value)
            li.puTime = CDate(lsh.Cells(li.sourceRow, EVO.G_PU_TIME_COLUMN).Value)
            li.puDate = li.puDate + li.puTime
            
            
            
            operacja.recalculDatesIfSatOrSun li
            
            If Trim(msh.Cells(optWiersz, G_DHEF_COL).Value) = "" Then
                msh.Cells(optWiersz, G_DHEF_COL).Value = li.puDate
                msh.Cells(optWiersz, G_DHEF_COL).Interior.Color = RGB(200, 255, 200)
            Else
                ' msh.Cells(optWiersz, G_DHEF_COL).Value = li.puDate
                msh.Cells(optWiersz, G_DHEF_COL).Interior.Color = RGB(255, 255, 150)
            End If
            
            If Trim(msh.Cells(optWiersz, G_DHAS_COL).Value) = "" Then
                msh.Cells(optWiersz, G_DHAS_COL).Value = li.delDate
                msh.Cells(optWiersz, G_DHAS_COL).Interior.Color = RGB(200, 255, 200)
            Else
                ' msh.Cells(optWiersz, G_DHAS_COL).Value = li.delDate
                msh.Cells(optWiersz, G_DHAS_COL).Interior.Color = RGB(255, 255, 150)
            End If
                
            If Trim(msh.Cells(optWiersz, G_DHAS_COL + 1).Value) = "" Then
            
                If IsNumeric(msh.Range("C" & CStr(optWiersz))) Then
                    If IsNumeric(msh.Range("Y" & CStr(optWiersz))) Then
                        If CLng(msh.Range("Y" & CStr(optWiersz))) > 0 Then
                            msh.Cells(optWiersz, G_DHAS_COL + 1).Formula = "=C" & optWiersz & "/Y" & optWiersz
                        Else
                            msh.Cells(optWiersz, G_DHAS_COL + 1).Value = "calc err: div by zero!"
                        End If
                    Else
                        msh.Cells(optWiersz, G_DHAS_COL + 1).Value = "calc err: source is not numeric!"
                    End If
                Else
                    msh.Cells(optWiersz, G_DHAS_COL + 1).Value = "calc err: source is not numeric!"
                End If
                
                
                If IsNumeric(msh.Cells(optWiersz, G_DHAS_COL + 1).Value) Then
                    
                    On Error Resume Next
                    If CDbl(msh.Cells(optWiersz, G_DHAS_COL + 1).Value) = Int(msh.Cells(optWiersz, G_DHAS_COL + 1).Value) Then
                        On Error Resume Next
                        msh.Cells(optWiersz, G_DHAS_COL + 1).Value = Int(msh.Cells(optWiersz, G_DHAS_COL + 1).Value)
                    Else
                        msh.Cells(optWiersz, G_DHAS_COL + 1).Value = Int(msh.Cells(optWiersz, G_DHAS_COL + 1).Value) + 1
                    End If
                
                End If
            Else
                ' msh.Cells(optWiersz, G_DHAS_COL + 1).Formula = "=C" & optWiersz & "/Y" & optWiersz
                msh.Cells(optWiersz, G_DHEF_COL + 1).Interior.Color = RGB(255, 255, 150)
            End If
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




Public Sub init(m_master As Worksheet, m_feed As Worksheet, Optional dwa As E_COPY_HANDLER_INIT)


    Set master = m_master
    Set feed = m_feed
    
    
    If (Not master Is Nothing) And (Not feed Is Nothing) Then
        ' we are OK to go further
        
        ' no need to pass args cos master and feed are class fields!
        ' -------------------------------------------------------------
        ' init by wrong cofro - but that giving any hurt
        ' dwa is funny name for enum for type of filling dictionary!
        
        
        
        ' additional sheets to verify in this case!
        If dwa = E_COPY_HANDLER_FOR_PIVOT_CREATION Then
        
            ' additional sheets to verify!
            ' ------------------------------------------------
            
            Set routeSheet = Nothing
            On Error Resume Next
            Set routeSheet = feed.Parent.Sheets(EVO.G_FEED_SH_CLOE)
            
            Set pleSheet = Nothing
            On Error Resume Next
            Set pleSheet = feed.Parent.Sheets(EVO.G_FEED_SH_PLE)
            
            
            Set uaSheet = Nothing
            On Error Resume Next
            Set uaSheet = feed.Parent.Sheets(EVO.G_FEED_SH_UA)
            
            
            
            
            ' ------------------------------------------------
        End If
        
        initAllDictionaries dwa
        
        ' -------------------------------------------------------------
        
    Else
        ' NOK situation
        MsgBox "Sth went wrong with crucial data!", vbCritical
        End
    End If

End Sub

Private Sub initAllDictionaries(dwa As E_COPY_HANDLER_INIT)

    Dim thisLine As LineItem
    Dim base As Worksheet
    Set base = master
    
    Dim r As Range
    Set r = base.Cells(3, EVO.E_MASTER_Reference)

    If dwa = E_COPY_HANDLER_FOR_PIVOT_CREATION Then
    
        Dim strTempExpCofor As String
        Set expCofors = New Dictionary
    
        Do
            If r.EntireRow.Hidden Then
                ' no operation req for filter out data!
            Else
            
                ' main logic for dwa = 3 internal joke ;)
                ' ------------------------------------------------------------
                
                strTempExpCofor = Trim(CStr(r.Offset(0, EVO.G_COFOR_EXPEDITEUR_COL - 1).Value))
                
                If Not expCofors.Exists(strTempExpCofor) Then
                
                
                    Set linesIter = Nothing
                    Set linesIter = New Dictionary
                    Set thisLine = New LineItem
                    
                    With thisLine
                        .pn = Trim(r.Value)
                        .cofor = strTempExpCofor
                        ' !!!
                        Set .masterRng = r
                        '
                        .wiersz = CLng(r.Row)
                        .tmc = Trim(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                        .yyyycw = CLng(operacja.calculateYYYYCW(Trim(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                        .monday = CDate(operacja.getMondayFromYYYYCW(CLng(.yyyycw)))
                        On Error Resume Next
                        .puDate = CDate(r.Offset(0, EVO.G_DHEF_COL - 1).Value)
                        On Error Resume Next
                        .delDate = CDate(r.Offset(0, EVO.G_DHAS_COL - 1).Value)
                        .isPast = CBool(operacja.isThisLineIsPast2(.puDate))
                        
                        
                        ' new from 0.16
                        ' -------------------------------------------------------------------
                        If IsNumeric(r.Offset(0, EVO.E_MASTER_QTY - 1).Value) Then
                            .qty2 = CLng(r.Offset(0, EVO.E_MASTER_QTY - 1).Value)
                        End If
                        
                        If IsNumeric(r.Offset(0, EVO.E_MASTER_CQTY - 1).Value) Then
                            .confirmedQty = r.Offset(0, EVO.E_MASTER_CQTY - 1).Value
                        End If
                        
                        .condi = master.Cells(r.Row, EVO.G_CONDI_COL).Value
                        ' -------------------------------------------------------------------
                    End With
                    
                    linesIter.Add r.Row, thisLine
                    expCofors.Add strTempExpCofor, linesIter
                
                Else
                
                    ' already exists
                    Set linesIter = expCofors(strTempExpCofor)
                    Set thisLine = New LineItem
                
                    With thisLine
                        .pn = Trim(r.Value)
                        .cofor = strTempExpCofor
                        ' !!!
                        Set .masterRng = r
                        '
                        .wiersz = CLng(r.Row)
                        .tmc = Trim(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                        
                        .yyyycw = CLng(operacja.calculateYYYYCW(Trim(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                        .monday = CDate(operacja.getMondayFromYYYYCW(CLng(.yyyycw)))
                        On Error Resume Next
                        .puDate = CDate(r.Offset(0, EVO.G_DHEF_COL - 1).Value)
                        On Error Resume Next
                        .delDate = CDate(r.Offset(0, EVO.G_DHAS_COL - 1).Value)
                        .isPast = CBool(operacja.isThisLineIsPast2(.puDate))
                        
                        
                        
                        ' new from 0.16
                        ' -------------------------------------------------------------------
                        If IsNumeric(r.Offset(0, EVO.E_MASTER_QTY - 1).Value) Then
                            .qty2 = CLng(r.Offset(0, EVO.E_MASTER_QTY - 1).Value)
                        End If
                        
                        If IsNumeric(r.Offset(0, EVO.E_MASTER_CQTY - 1).Value) Then
                            .confirmedQty = r.Offset(0, EVO.E_MASTER_CQTY - 1).Value
                        End If
                        
                        .condi = master.Cells(r.Row, EVO.G_CONDI_COL).Value
                        ' -------------------------------------------------------------------
                        
                    End With
                    
                    
                    ' just to add inner dictionary
                    ' because linesIter is already a ref inside dictionary!
                    linesIter.Add r.Row, thisLine
                
                
                End If
                
                ' ------------------------------------------------------------
            End If
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
    
        

    ElseIf dwa = E_COPY_HANDLER_BY_TMC_OPT Then
        Dim strTempPn As String
        Dim strTempCoforInner As String
        Dim strTempPnAndCofor As String
        Set pns2 = New Dictionary
        
        Do
        
            If r.EntireRow.Hidden Then
                ' no operation req for filter out data!
            Else
            
                strTempPn = Trim(r.Value)
                strTempCoforInner = Trim(CStr(r.Offset(0, EVO.G_COFOR_VENDEUR_COL - 1).Value))
                strTempCoforInner = operacja.adjustStringCofor(CStr(strTempCoforInner))
                
                strTempPnAndCofor = strTempPn & "_" & strTempCoforInner
                
                If Not pns2.Exists(strTempPnAndCofor) Then
                
                
                    Set linesIter = Nothing
                    Set linesIter = New Dictionary
                    Set thisLine = New LineItem
                    
                    With thisLine
                        .pn = strTempPn
                        .cofor = strTempCoforInner
                        ' !!!
                        Set .masterRng = r
                        '
                        .wiersz = CLng(r.Row)
                        .tmc = Trim(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                        .yyyycw = CLng(operacja.calculateYYYYCW(Trim(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                        .monday = CDate(operacja.getMondayFromYYYYCW(CLng(.yyyycw)))
                        On Error Resume Next
                        .puDate = CDate(r.Offset(0, EVO.G_DHEF_COL - 1).Value)
                        On Error Resume Next
                        .delDate = CDate(r.Offset(0, EVO.G_DHAS_COL - 1).Value)
                        .isPast = CBool(operacja.isThisLineIsPast2(.puDate))
                        
                    End With
                    
                    linesIter.Add r.Row, thisLine
                    pns2.Add strTempPnAndCofor, linesIter
                Else
                
                    Set linesIter = pns2(strTempPnAndCofor)
                    Set thisLine = New LineItem
                
                    With thisLine
                        .pn = strTempPn
                        .cofor = strTempCoforInner
                        ' !!!
                        Set .masterRng = r
                        '
                        .wiersz = CLng(r.Row)
                        .tmc = Trim(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                        
                        .yyyycw = CLng(operacja.calculateYYYYCW(Trim(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                        .monday = CDate(operacja.getMondayFromYYYYCW(CLng(.yyyycw)))
                        On Error Resume Next
                        .puDate = CDate(r.Offset(0, EVO.G_DHEF_COL - 1).Value)
                        On Error Resume Next
                        .delDate = CDate(r.Offset(0, EVO.G_DHAS_COL - 1).Value)
                        .isPast = CBool(operacja.isThisLineIsPast2(.puDate))
                        
                    End With
                    
                    linesIter.Add r.Row, thisLine
                    
                End If
            End If
            
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""

    ElseIf dwa = E_COPY_HANDLER_COPY_ONE Then

        
        Dim strTempCofor As String
        Set cofors = New Dictionary

        Do
        
        
            If r.EntireRow.Hidden Then
                ' no operation req for filter out data!
            Else
        
        
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
                        .yyyycw = CLng(operacja.calculateYYYYCW(Trim(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                        .monday = CDate(operacja.getMondayFromYYYYCW(CLng(.yyyycw)))
                        .isPast = CBool(operacja.isThisLineIsPast(CLng(.yyyycw)))
                        
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
                        .yyyycw = CLng(operacja.calculateYYYYCW(Trim(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                        .monday = CDate(operacja.getMondayFromYYYYCW(CLng(.yyyycw)))
                        .isPast = CBool(operacja.isThisLineIsPast(CLng(.yyyycw)))
                        
                    End With
                    
                    
                    linesIter.Add r.Row, thisLine
                    
                    
                End If
            End If
        
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        
        
        
        ' afer all check math
        Debug.Print "count cofors: " & cofors.Count
        
        ' check all lines
        Dim xl As Integer, k As Variant
        xl = 0
        For Each k In cofors.Keys
            
            Set linesIter = cofors(k)
            
            Debug.Print "for cofor: " & k & " we have: " & linesIter.Count & " lines"
            xl = xl + linesIter.Count
            
        Next
        
        ' -1 for loop -2 for labels
        Debug.Print "checksum on all lines: " & xl & " rows in file: " & (r.Row - 1 - 2)
    
        
        ' for now without this checking!
        'If CLng(xl) = CLng(r.Row - 1 - 2) Then
        '    '  OK
        'Else
        '    MsgBox "Checksum on data is not right - there is sth not right with the BASE worksheet!", vbCritical
        'End If
    End If
End Sub




