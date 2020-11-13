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


Private v As Validator



Private e As E_ECHANCIER_ONL_semaine_SCENARIO


Private xtraRep As XtraCopyDataReport



Private Sub Class_Initialize()

    e = E_ECHANCIER_ONL_semaine_SCENARIO_PU

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
    
    
    Set v = New Validator
    
    Set xtraRep = New XtraCopyDataReport
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
    
    Set v = Nothing
    
    Set xtraRep = Nothing
End Sub


Public Sub optimise(sh As StatusHandler)


    Dim li As LineItem
    Dim byTMC As Dictionary
    Dim innerDicForByTMC As Dictionary
    Dim k1 As Variant, k2 As Variant, k3 As Variant, k4 As Variant
    Dim tmcKey As String

    
    If Not pns2 Is Nothing Then
        If pns2.count > 0 Then
        
            sh.init_statusbar (pns2.count + 1)
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
                    
                    ' Debug.Print "count TMCs: " & tmcKey & " " & byTMC.Count & " for PN: " & k1 & " wiersz: " & li.wiersz
                    
                    
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
                            
                            If innerDicForByTMC.count = 1 Then
                            
                                For Each k4 In innerDicForByTMC.Keys
                                    minWiersz = k4
                                Next
                                
                            ElseIf innerDicForByTMC.count > 1 Then
                            
                                
                                
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
    Dim cofors1 As String
    Dim cofors2 As String
    Dim tmp_x As String, tmp_x2 As String
    Dim xadr1 As String, xadr2 As String
    
    If Not expCofors Is Nothing Then
        If expCofors.count > 0 Then
        
            sh.init_statusbar (expCofors.count + 1)
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
                                cofors1 = ""
                                cofors2 = ""
                                
                                Set linesIter = expCofors(k1)
                                
                                For Each k2 In linesIter
                                    
                                    Set li = Nothing
                                    
                                    On Error Resume Next
                                    Set li = linesIter(k2)
                                    
                                    expCofor = li.cofor
                                    cofors1 = li.COFOR_COFOR
                                    cofors2 = Replace(cofors1, "_", "")
                                    condi = li.condi
                                    
                                    
                                    ' Debug.Assert li.wiersz <> 2866
                                    
                                    
                                    
                                    ' findByPartNumber
                                    Set findByExpCoforInPle = Nothing
                                    Set findByExpCoforInCloe = Nothing
                                    Set findByExpCoforInMain = Nothing
                                    Set findByCondiInUa = Nothing
                                    
                                    'Global Const G_PLE_VEN_COFOR = 6
                                    'Global Const G_PLE_SHIPPER_COFOR = 7
                                    'Global Const G_CLOE_SHIPPER_COFOR = 1
                                    'Global Const G_CLOE_COFORS = 2
                                    'Global Const G_UA_KEY = 1
                                    'Global Const G_UA_MAX_CAPACITY_COLUMN = 6
                                    'Global Const G_FEED_MAIN_SH_VEN_COFOR = 3
                                    'Global Const G_FEED_MAIN_SH_SHIPPER_COFOR = 4
                                    'Global Const G_FEED_MAIN_SH_CONDI = 10
                                    
                                    On Error Resume Next
                                    Set findByExpCoforInPle = pleRng.Find(expCofor)
                                    
                                    
                                    ' Debug.Assert expCofor <> "A001KK  01"
                                    
                                    If Not findByExpCoforInPle Is Nothing Then
                                    
                                        xadr1 = findByExpCoforInPle.Address
                                    
                                        tmp_x = _
                                            findByExpCoforInPle.Offset(0, G_PLE_VEN_COFOR - G_PLE_SHIPPER_COFOR).Value & "_" & _
                                            CStr(findByExpCoforInPle.Value)
                                            
                                        If tmp_x <> cofors1 Then
                                        
                                            Do
                                                ' Debug.Print pleRng.Address
                                            
                                                Set findByExpCoforInPle = pleRng.FindNext(findByExpCoforInPle)
                                                xadr2 = findByExpCoforInPle.Address
                                            
                                                tmp_x2 = _
                                                    findByExpCoforInPle.Offset(0, G_PLE_VEN_COFOR - G_PLE_SHIPPER_COFOR).Value & "_" & _
                                                    CStr(findByExpCoforInPle.Value)
                                                    
                                                If xadr1 = xadr2 Or tmp_x2 = cofors1 Then
                                                    Exit Do
                                                End If
                                                    
                                                
                                            Loop Until xadr1 = xadr2
                                            
                                            If tmp_x2 <> cofors1 Then
                                                Set findByExpCoforInPle = Nothing
                                            End If
                                            
                                        End If
                                    End If
                                    
                                    
                                    On Error Resume Next
                                    Set findByExpCoforInCloe = cloeRng.Find(expCofor)
                                    
                                    

                                    
                                    
                                    ' Debug.Assert expCofor <> "A001KK  01"
                                    
                                    If Not findByExpCoforInCloe Is Nothing Then
                                    
                                        xadr1 = findByExpCoforInCloe.Address
                                    
                                        tmp_x = _
                                            findByExpCoforInCloe.Offset(0, G_CLOE_COFORS - G_CLOE_SHIPPER_COFOR).Value
                                            
                                        If tmp_x <> cofors2 Then
                                        
                                            Do
                                                ' Debug.Print pleRng.Address
                                            
                                                Set findByExpCoforInCloe = cloeRng.FindNext(findByExpCoforInCloe)
                                                xadr2 = findByExpCoforInCloe.Address
                                            
                                                tmp_x2 = _
                                                    findByExpCoforInCloe.Offset(0, G_CLOE_COFORS - G_CLOE_SHIPPER_COFOR).Value
                                                    
                                                If xadr1 = xadr2 Or tmp_x2 = cofors2 Then
                                                    Exit Do
                                                End If

                                                
                                            Loop Until xadr1 = xadr2
                                            
                                                    
                                            If tmp_x2 <> cofors2 Then
                                                Set findByExpCoforInCloe = Nothing
                                            End If
                                            
                                        End If
                                    End If
                                    
                                    
                                    On Error Resume Next
                                    Set findByExpCoforInMain = srcRng.Find(expCofor)
                                    
                                    ' Debug.Assert expCofor <> "A001KK  01"
                                    
                                    If Not findByExpCoforInMain Is Nothing Then
                                    
                                        xadr1 = findByExpCoforInMain.Address
                                    
                                        tmp_x = _
                                            findByExpCoforInMain.Offset(0, G_FEED_MAIN_SH_VEN_COFOR - G_FEED_MAIN_SH_SHIPPER_COFOR).Value & "_" & _
                                            CStr(findByExpCoforInMain.Value)
                                            
                                        If tmp_x <> cofors1 Then
                                        
                                            Do
                                                ' Debug.Print pleRng.Address
                                            
                                                Set findByExpCoforInMain = srcRng.FindNext(findByExpCoforInMain)
                                                xadr2 = findByExpCoforInMain.Address
                                            
                                                tmp_x2 = _
                                                    findByExpCoforInMain.Offset(0, G_FEED_MAIN_SH_VEN_COFOR - G_FEED_MAIN_SH_SHIPPER_COFOR).Value & _
                                                    "_" & _
                                                    CStr(findByExpCoforInMain.Value)
                                                    
                                                If xadr1 = xadr2 Or tmp_x2 = cofors1 Then
                                                    Exit Do
                                                End If
                                                    
                                                
                                            Loop Until xadr1 = xadr2
                                            
                                            If tmp_x2 <> cofors1 Then
                                                Set findByExpCoforInMain = Nothing
                                            End If
                                            
                                        End If
                                    End If
                                    
                                    
                                    
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
                                            li.sourceRow = findByExpCoforInMain.row
                                                
                                            ' main coping for pivot source!!
                                            ' ================================================================
                                            ''
                                            '
                                            insideCopyForSourcePivot_setupOrderDate li, linesIter, pleSheet
                                            ' readjustDatesAgainIfThereIsDifferentInPusMaster li, linesIter, master
                                            mainCopyForPivotSource li, linesIter, ph
                                            '
                                            ''
                                            ' ================================================================
                                    Else
                                        ' =====================================================================
                                        ' =====================================================================
                                        li.validationForPivotSource = False
                                        Set li.supplementForPivotSource = New LineItemPivotSrouceSupplement
                                        
                                        With li.supplementForPivotSource
                                            Set .refRangeForCloe = findByExpCoforInCloe
                                            Set .refRangeForPLE = findByExpCoforInPle
                                            Set .refRaneForCondi = findByCondiInUa
                                        End With
                                        
                                        ' those data are not full!
                                        
                                        ' =====================================================================
                                        ' =====================================================================
                                    End If
                                    
                                    
                                Next
                                
                                
                                sh.progress_increase
                                
                            Next
                            
                            
                            
                            ' double loop for creating NOK report for TCAM logic!
                            doubleLoopForNOKrepForTCAMLogic k1, k2, expCofors, linesIter, li

                            
                            
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
            ' ph.proxy2.name = ph.tryToRenameProxy2Worksheet(ph.proxy2)
            
            ph.proxy2.name = EVO.TryToRenameModule.tryToRenameWorksheet(ph.proxy2, "Proxy2_")
            
            ' ===============================================================================
            
            
            ' HIDE!
            sh.hide
            
            
        End If
    End If
End Sub




Public Sub copyData(sh As StatusHandler)



    ' Private e As E_ECHANCIER_ONL_semaine_SCENARIO
    
    
    If Not cofors Is Nothing Then
        If cofors.count > 0 Then
        
            sh.init_statusbar (cofors.count + 1)
            sh.show
            sh.progress_increase
            
            Set pns = New Dictionary
        
            
            If Not feed Is Nothing Then
            
            
            
                Dim ycwMOnday As Long
                Dim ycwToday As Long
            
            
                Dim pn As String, k1 As Variant, k2 As Variant, pnWithIndice As String
                Dim li As LineItem
                
                Dim srcRng As Range, f As Range
                Set srcRng = feed.Cells(2, 1)
                ' Set srcRng = srcRng.EntireColumn ' ?
                
                If srcRng.Offset(1, 0).Value <> "" Then
                    Set srcRng = feed.Range(srcRng, srcRng.End(xlDown))
                End If
                
                Debug.Print "control tower range: " & srcRng.Address & " " & srcRng.count
            
                For Each k1 In cofors.Keys
                
                    pn = ""
                
                    Set linesIter = cofors(k1)
                    
                    For Each k2 In linesIter
                        
                        Set li = Nothing
                        
                        On Error Resume Next
                        Set li = linesIter(k2)
                        
                        pnWithIndice = li.pn
                        pn = operacja.removeIndice(CStr(li.pn))
                        
                        
                        Set f = Nothing
                        
                        If pns.Exists(pn) Then
                            Set f = pns(pn)
                        Else
                        
                        
                            On Error Resume Next
                            Set f = srcRng.Find(pn)
                            
                            If f Is Nothing Then
                                Dim irSrc As Range
                                For Each irSrc In srcRng
                                    ' Debug.Print irSrc.Value & " , " & pn
                                    
                                    
                                    If irSrc.Value = pnWithIndice Then
                                        Set f = irSrc
                                        Exit For
                                        
                                    ElseIf irSrc.Value = pn Then
                                        Set f = irSrc
                                        Exit For
                                    
                                    ElseIf irSrc.Value Like pn & "-??" Then
                                        Set f = irSrc
                                        Exit For
                                    End If
                                    

                                Next irSrc
                                
                            End If
                            
                            pns.Add pn, f
                            
                        End If
                        
                        If f Is Nothing Then
                            li.found = False
                            li.addRawInfoToLog "this PN is missing in CPL BASE: Control Tower DB!"
                            
                            ' no data in input!
                            ' put red mark on part number!
                            ' ======================================================
                            master.Cells(li.wiersz, 1).Font.Bold = True
                            master.Cells(li.wiersz, 1).Font.Color = RGB(255, 0, 0)
                            ' ======================================================
                            
                            
                        Else
                            li.found = True
                            Set li.sourceRng = f
                            li.sourceRow = f.row
                            
                            
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
                        
                        
                        
                        ' xtra NOK report!
                        ' =============================================================
                        If Len(li.getFullLog()) > 0 Then xtraRep.addLiToCollection li
                        ' =============================================================
                        
                    Next
                                        
                    sh.progress_increase
                    
                Next
                
                
                xtraRep.createXtraReport
            
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


Private Sub readjustDatesAgainIfThereIsDifferentInPusMaster(li As LineItem, ByRef linesIter As Dictionary, pusSh As Worksheet)
    
    ' Debug.Print "inside readjustDatesAgainIfThereIsDifferentInPusMaster and pus nameL " & pusSh.name
    ' Debug.Print li.wiersz & " " & li.sourceRow ' OK!
    
    Dim dhefInPus As Range
    Dim dhasInPus As Range
    
    Set dhefInPus = pusSh.Cells(li.wiersz, EVO.E_MASTER_2_DHEF)
    Set dhasInPus = pusSh.Cells(li.wiersz, EVO.E_MASTER_2_DHAS)
    
    Debug.Print "dhefInPus: " & dhefInPus.Value
    Debug.Print "dhasInPus: " & dhasInPus.Value
    
    Debug.Print "suple: " & li.supplementForPivotSource.orderDate '& " " & li.supplementForPivotSource.orderTime
    Debug.Print "pus del: " & li.delDate '& " " & li.delTime
    Debug.Print "pus pu: " & li.puDate '& " " & li.puTime
    
    If IsDate(dhefInPus.Value) And IsDate(li.supplementForPivotSource.orderDate) Then
        
        Debug.Print "dhefInPus is date"
    End If
End Sub


Private Function calculateOrderDate(li As LineItem, pleSh As Worksheet) As Date


    Dim cellWithProperValue As Range
    Set cellWithProperValue = _
        pleSh.Cells(li.supplementForPivotSource.refRangeForPLE.row, EVO.G_PLE_SUB_FOR_ORDER_COLUMN)

    Dim howManyDaysOfOffset As Integer
    
    If IsNumeric(cellWithProperValue) Then
        howManyDaysOfOffset = Int(cellWithProperValue.Value)
    Else
        howManyDaysOfOffset = -1
    End If
    
    
    
    If howManyDaysOfOffset <> -1 Then
    
        Dim tmpdate As Date, xo As Variant
        On Error Resume Next
        tmpdate = CDate(li.delDate) '  - howManyDaysOfOffset
        
        For xo = 1 To howManyDaysOfOffset
        
            tmpdate = tmpdate - 1
            
            If isSunday(tmpdate) Then
                tmpdate = tmpdate - 2
            ElseIf isSaturday(tmpdate) Then
                tmpdate = tmpdate - 1
            End If
            
        Next xo
        
        li.supplementForPivotSource.heoDecalageJX = Int(howManyDaysOfOffset)
        
        
        If isSaturday(tmpdate) Then
            calculateOrderDate = CDate(tmpdate) - 1
        ElseIf isSunday(tmpdate) Then
            calculateOrderDate = CDate(tmpdate) - 2
        Else
            calculateOrderDate = CDate(tmpdate)
        End If
    End If
End Function


Private Function isSaturday(psd As Date) As Boolean
    
    If Weekday(psd, vbMonday) = 6 Then
        isSaturday = True
    Else
        isSaturday = False
    End If
    
End Function

Private Function isSunday(psd As Date)

    If Weekday(psd, vbMonday) = 7 Then
        isSunday = True
    Else
        isSunday = False
    End If

End Function


Private Sub mainCopyForPivotSource(li As LineItem, ByRef linesIter As Dictionary, ph As PivotHandler)


    If checkLineItemIfItMakeAnySense(li) Then
    
    
        
        
        ' there is data!
        'Debug.Print "checkLineItemIfItMakeAnySense = true " & li.wiersz
        
        Dim sh As Worksheet, linia As Long, linia2 As Long
        Set sh = ph.getPivotSource()
        linia2 = li.wiersz
        linia = ph.getRefRangeInSrcSheet().row
        ' ------------------------------------------------------
        ci.D_copyIterationForPivotSource master, sh, li, ph, linia
        ' ------------------------------------------------------
        ph.moveRefRangeInSrcSheet
        
    Else
        ' no proper data from source sheets
        Debug.Print "checkLineItemIfItMakeAnySense = false"
    End If

End Sub

Private Sub doubleLoopForNOKrepForTCAMLogic(mmk1 As Variant, mmk2 As Variant, _
    ByRef mmexpCofors As Dictionary, _
    ByRef mmlinesIter As Dictionary, _
    mmli As LineItem)
    
    Dim shnok As Worksheet
    Set shnok = ThisWorkbook.Sheets.Add
    
    Dim o As New OperationsHandler
    shnok.name = EVO.TryToRenameModule.tryToRenameWorksheet(shnok, "TCAM_NOKs_")
    
    Dim r1 As Range
    Set r1 = shnok.Cells(1, 1)
    r1.Value = "REF"
    r1.Offset(0, 1).Value = "ROW"
    r1.Offset(0, 2).Value = "CC"
    r1.Offset(0, 3).Value = "CONDI NOK"
    r1.Offset(0, 4).Value = "CLOE NOK"
    r1.Offset(0, 5).Value = "PLE NOK"
    
    Set r1 = r1.Offset(1, 0)
    
    

    For Each mmk1 In mmexpCofors.Keys
        
        Set mmlinesIter = mmexpCofors(mmk1)
        
        For Each mmk2 In mmlinesIter
            
            Set mmli = Nothing
            
            On Error Resume Next
            Set mmli = mmlinesIter(mmk2)
            
            
            If Not mmli.validationForPivotSource Then
                r1.Value = mmli.masterRng.Value
                r1.Offset(0, 1).Value = mmli.masterRng.row
                r1.Offset(0, 2).Value = mmli.COFOR_COFOR
                
                If Not mmli.supplementForPivotSource Is Nothing Then
                
                    If mmli.supplementForPivotSource.refRaneForCondi Is Nothing Then
                        r1.Offset(0, 3).Value = "CONDI NOK"
                    Else
                        r1.Offset(0, 3).Value = mmli.supplementForPivotSource.refRaneForCondi.Value
                    End If
                    
                    If mmli.supplementForPivotSource.refRangeForCloe Is Nothing Then
                        r1.Offset(0, 4).Value = "CLOE NOK"
                    Else
                        r1.Offset(0, 4).Value = mmli.supplementForPivotSource.refRangeForCloe.Value
                    End If
                    
                    If mmli.supplementForPivotSource.refRangeForPLE Is Nothing Then
                        r1.Offset(0, 5).Value = "PLE NOK"
                    Else
                        r1.Offset(0, 5).Value = mmli.supplementForPivotSource.refRangeForPLE.Value
                    End If
                
                End If
                
                
                Set r1 = r1.Offset(1, 0)
            End If
            
            
        Next mmk2
    Next mmk1
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
        
        If IsNumeric(od.Value) Then
            kolumna1 = od.Value
        Else
            MsgBox "Fatal type error on CopyHandler line 1112", vbCritical
            End
        End If
        
        If IsNumeric(od.Offset(0, 3).Value) Then
            kolumna2 = od.Offset(0, 3).Value
        Else
            MsgBox "Fatal type error on CopyHandler line 1112", vbCritical
            End
        End If
        
        'msh.Cells(optWiersz, 1).Interior.Color = RGB(255, 255, 255)
        'msh.Cells(optWiersz, 1).Font.Color = RGB(0, 0, 0)
        
    
        If IsNumeric(od.Offset(0, 3).Value) Then
            
            'simple copy without X
            If CStr(od.Offset(0, 5).Value) = "" Then
            
                If IsError(lsh.Cells(optSrcWiersz, kolumna2).Value) Then
                    ' msh.Cells(optWiersz, kolumna1).Value = "error n/a in source file"
                    msh.Cells(optWiersz, kolumna1).Interior.Color = RGB(255, 200, 200)
                   ' msh.Cells(optWiersz, 1).Interior.Color = RGB(255, 200, 200)
                   li.addRawInfoToLog "error in source file! source: " & CStr(lsh.Cells(optSrcWiersz, kolumna2).Address)
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
                    ElseIf CStr(msh.Cells(optWiersz, kolumna1).Value) = "" Then
                        msh.Cells(optWiersz, kolumna1).Value = lsh.Cells(optSrcWiersz, kolumna2).Value
                        msh.Cells(optWiersz, kolumna1).Interior.Color = RGB(200, 255, 200)
                        'msh.Cells(optWiersz, 1).Interior.Color = RGB(200, 255, 200)
                        
                    ' diff data - yellow color
                    ElseIf msh.Cells(optWiersz, kolumna1).Value <> lsh.Cells(optSrcWiersz, kolumna2).Value Then
                    
                    
                        ' li.addRawInfoToLog "yellow data! PLEASE EXAMINE: " & CStr(lsh.Cells(optSrcWiersz, kolumna2).Address)
                    
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
            ElseIf CStr(od.Offset(0, 5).Value) = "X" Then

                calculationForDH li, linesIter, CStr(od.Offset(0, 4).Value), lsh, msh, CLng(optWiersz)
                
            End If
        End If

    
        Set od = od.Offset(1, 0)
    Loop Until CStr(od) = ""
    
    '
    ''
    ' ---------------------------------------------------
End Sub


Private Sub calculationForDH(li As LineItem, ByRef linesIter As Dictionary, dhType As String, lsh As Worksheet, msh As Worksheet, optWiersz As Long)
    
    
    
    ' new stuff!
    ' Private e As E_ECHANCIER_ONL_semaine_SCENARIO
    
    ' find in project catch keywords/ hashtags:
    ' calculation For DH, calculation For DHxx, #defining_DH, #DHxx, #DHXX, #DH
    
    
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
        li.CodEntrega = Trim(lsh.Cells(li.sourceRow, EVO.G_COD_TRANSPORT_COLUMN).Value)
        
        
        ' could be one line of code but i want to show boldly that is important!
        If v.checkIsCodeOK(li.CodEntrega) Then
            li.isCodEntregaAvailable = True
            
        Else
            li.isCodEntregaAvailable = False
            li.addRawInfoToLog "cod entrega not available"
        End If
        
        
        ' Debug.Assert li.wiersz <> 3628
        
        
        ' full ignore on DH if there is a DAP or DDP in CodEntrega
        If li.CodEntrega <> EVO.G_DAP And li.CodEntrega <> EVO.G_DDP Then
        
            ' this part is added on 202009xx
            ' when you already know what you want to put
            ' as dates
            ' of coure it is much better
            ' to use prev logic
            ' but what can I say...
            ' ---------------------------------------------------------------------------
            ' this part is for copy
            ' static dates in DHxx
            ' if there is any inside control tower db
            ' ---------------------------------------------------------------------------
            
            assignStaticDatesIfThereIsAnyData lsh, li
            

            ' ---------------------------------------------------------------------------
            ' ---------------------------------------------------------------------------
            
            ' e = E_ECHANCIER_ONL_semaine_SCENARIO_DEL
            ' e = E_ECHANCIER_ONL_semaine_SCENARIO_PU
            ' ---------------------------------------------------------------------------
            ' ---------------------------------------------------------------------------
            
            tryToAssignProperlyDHxxDatesWithDefaultCalculation e, li, linesIter, dhType, lsh, msh, optWiersz
            
            ' ---------------------------------------------------------------------------
            ' ---------------------------------------------------------------------------
            
            
            
            
            
            ' this logic here is not fully about DHxx columns
            ' this is about column  no orderers
            ' it is partially affected by DHxx that is why it is here
            ' ---------------------------------------------------------------------
            ' ---------------------------------------------------------------------
            If CStr(msh.Cells(optWiersz, G_DHAS_COL + 1).Value) = "" Then
            
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
                    
                Else
                    
                    li.addRawInfoToLog "qty and packing data is not numeric! col: N0 Ordenes"
                
                End If
            Else
                ' msh.Cells(optWiersz, G_DHAS_COL + 1).Formula = "=C" & optWiersz & "/Y" & optWiersz
                msh.Cells(optWiersz, G_DHAS_COL + 1).Interior.Color = RGB(255, 255, 150)
            End If
            ' ---------------------------------------------------------------------
            ' ---------------------------------------------------------------------
            
        Else
            li.addRawInfoToLog "Incoterm: DAP or DDP"
        End If
        
        ' -----------------------------------------------------------------------
        ' -----------------------------------------------------------------------
    ElseIf dhType = "DHAS" Then
    
        ' !!!
        ' no operation
        ' this is becuase first "X" in register appears in DHEF for pickup
        ' and already make all logic that stays behind calc of DH
        ' !!!
    Else
        ' nop
    End If
End Sub


Private Sub assignStaticDatesIfThereIsAnyData(lsh As Worksheet, ByRef li As LineItem)


    If Trim(lsh.Cells(li.sourceRow, EVO.G_SRC_DHAS_COL).Value) <> "" Then
        li.delDate = CDate(lsh.Cells(li.sourceRow, EVO.G_SRC_DHAS_COL).Value)
        li.addRawInfoToLog "DHAS is manual!"
    End If
    
    If Trim(lsh.Cells(li.sourceRow, EVO.G_SRC_DHEF_COL).Value) <> "" Then
        li.puDate = CDate(lsh.Cells(li.sourceRow, EVO.G_SRC_DHEF_COL).Value)
        li.addRawInfoToLog "DHEF is manual!"
    End If
End Sub


Private Sub tryToAssignProperlyDHxxDatesWithDefaultCalculation(e As E_ECHANCIER_ONL_semaine_SCENARIO, _
    li As LineItem, _
    ByRef linesIter As Dictionary, _
    dhType As String, _
    lsh As Worksheet, _
    msh As Worksheet, _
    optWiersz As Long)
    
    
        ' e As E_ECHANCIER_ONL_semaine_SCENARIO
        ' is copied again just to be clear what is the params for this sub!
        
        msh.Cells(optWiersz, EVO.G_DHEF_COL).Interior.Color = RGB(255, 255, 255)
        msh.Cells(optWiersz, EVO.G_DHAS_COL).Interior.Color = RGB(255, 255, 255)
        
        If li.isCodEntregaAvailable = True Then
        
        
            If IsNumeric(lsh.Cells(li.sourceRow, EVO.G_T_TIME_COLUMN).Value) Then
    
    
                ' ============================================================================
                ' tryToAssignProperlyDHxxDatesWithDefaultCalculation
                ' !! full copy of params from upper sub
                ' ============================================================================
                If e = E_ECHANCIER_ONL_semaine_SCENARIO_DEL Then
                
                
                    ' this is the old logic when the echancier cw was treated as delivery date
                    
                    If CLng(li.delDate) = 0 Then
                    
                        li.dateAfterOffset = offsetDate(li.CodEntrega, li.monday)
                        li.delTime = CDate(lsh.Cells(li.sourceRow, EVO.G_DEL_TIME_COLUMN).Value)
                        li.delDate = li.dateAfterOffset + CDate(li.delTime)
                        
                    End If
                
                
                    If CLng(li.puDate) = 0 Then
                        
                        li.puDate = li.delDate - li.delTime - CDate(lsh.Cells(li.sourceRow, EVO.G_T_TIME_COLUMN).Value)
                        li.puTime = CDate(lsh.Cells(li.sourceRow, EVO.G_PU_TIME_COLUMN).Value)
                        li.puDate = li.puDate + li.puTime
                        
                    End If
                
                
                    ' operacja.recalculDatesIfSatOrSun li, E_ECHANCIER_ONL_semaine_SCENARIO_DEL
                    operacja.reculculIfBetweenDatesWeekend li, E_ECHANCIER_ONL_semaine_SCENARIO_DEL
                    
                    
                
                    If CStr(msh.Cells(optWiersz, G_DHEF_COL).Value) = "" Then
                        msh.Cells(optWiersz, G_DHEF_COL).Value = li.puDate
                        msh.Cells(optWiersz, G_DHEF_COL).Interior.Color = RGB(200, 255, 200)
                    Else
                        ' msh.Cells(optWiersz, G_DHEF_COL).Value = li.puDate
                        msh.Cells(optWiersz, G_DHEF_COL).Interior.Color = RGB(255, 255, 150)
                    End If
                    
                    If CStr(msh.Cells(optWiersz, G_DHAS_COL).Value) = "" Then
                        msh.Cells(optWiersz, G_DHAS_COL).Value = li.delDate
                        msh.Cells(optWiersz, G_DHAS_COL).Interior.Color = RGB(200, 255, 200)
                    Else
                        ' msh.Cells(optWiersz, G_DHAS_COL).Value = li.delDate
                        msh.Cells(optWiersz, G_DHAS_COL).Interior.Color = RGB(255, 255, 150)
                    End If
                    
                    
                    
                ElseIf e = E_ECHANCIER_ONL_semaine_SCENARIO_PU Then
                
                
                    ' #newapporachonDH
                    ' #ENCHACIER #CW #ECHANCIER ONL (semaine) #semaine
        
                
                    ' here we have totally new apporach, where enchancier cw is treated as
                    ' pikcup date - which means, that the parts in to be prepared a little
                    ' bit later - also there were a lot of confusion of the usage of this
                    ' tool - so maybe this new approach will help to provide better business
                    ' --------------------------------------------------------------------------------
                    ' --------------------------------------------------------------------------------
                    
                    If CLng(li.puDate) = 0 Then
                        
                        
                        ' ' logic for echancier as delivery:
                        'li.puDate = li.delDate - li.delTime - CDate(lsh.Cells(li.sourceRow, EVO.G_T_TIME_COLUMN).Value)
                        'li.puTime = CDate(lsh.Cells(li.sourceRow, EVO.G_PU_TIME_COLUMN).Value)
                        'li.puDate = li.puDate + li.puTime
                        
                        ' logic for echancier as pu:
                        li.dateAfterOffset = offsetDate(li.CodEntrega, li.monday)
                        li.puTime = CDate(lsh.Cells(li.sourceRow, EVO.G_PU_TIME_COLUMN).Value)
                        li.puDate = li.dateAfterOffset + CDate(li.puTime)
                        
                    End If
                    
                    If CLng(li.delDate) = 0 Then
                    
                        ' ' logic for echancier as delivery:
                        'li.dateAfterOffset = offsetDate(li.CodEntrega, li.monday)
                        'li.delTime = CDate(lsh.Cells(li.sourceRow, EVO.G_DEL_TIME_COLUMN).Value)
                        'li.delDate = li.dateAfterOffset + CDate(li.delTime)
                        
                        ' logic for echancier as pu:
                        li.delDate = li.puDate - li.puTime + CDate(lsh.Cells(li.sourceRow, EVO.G_T_TIME_COLUMN).Value)
                        li.delTime = CDate(lsh.Cells(li.sourceRow, EVO.G_DEL_TIME_COLUMN).Value)
                        li.delDate = li.delDate + li.delTime
                    End If
                
                
                
                    'operacja.recalculDatesIfSatOrSun li, E_ECHANCIER_ONL_semaine_SCENARIO_PU
                    operacja.reculculIfBetweenDatesWeekend li, E_ECHANCIER_ONL_semaine_SCENARIO_PU
                    
                    
                    Dim prvCmnt As String
                    
                    If CStr(msh.Cells(optWiersz, G_DHEF_COL).Value) = "" Then
                        msh.Cells(optWiersz, G_DHEF_COL).Value = li.puDate
                        msh.Cells(optWiersz, G_DHEF_COL).Interior.Color = RGB(200, 255, 200)
                    Else
                        ' msh.Cells(optWiersz, G_DHEF_COL).Value = li.puDate
                        'On Error Resume Next
                        'msh.Cells(optWiersz, G_DHEF_COL).Comment.Delete
                        prvCmnt = ""
                        If Not msh.Cells(optWiersz, G_DHEF_COL).Comment Is Nothing Then
                        
                            prvCmnt = msh.Cells(optWiersz, G_DHEF_COL).Comment.Text
                            msh.Cells(optWiersz, G_DHEF_COL).Comment.Delete
                        End If
                        msh.Cells(optWiersz, G_DHEF_COL).AddComment prvCmnt & Chr(10) & _
                            "not added: " & CStr(li.puDate)
                        msh.Cells(optWiersz, G_DHEF_COL).Comment.Shape.TextFrame.AutoSize = True
                        msh.Cells(optWiersz, G_DHEF_COL).Interior.Color = RGB(255, 255, 150)
                    End If
                    
                    If CStr(msh.Cells(optWiersz, G_DHAS_COL).Value) = "" Then
                        msh.Cells(optWiersz, G_DHAS_COL).Value = li.delDate
                        msh.Cells(optWiersz, G_DHAS_COL).Interior.Color = RGB(200, 255, 200)
                    Else
                        ' msh.Cells(optWiersz, G_DHAS_COL).Value = li.delDate
                        'On Error Resume Next
                        'msh.Cells(optWiersz, G_DHAS_COL).Comment.Delete
                        prvCmnt = ""
                        If Not msh.Cells(optWiersz, G_DHAS_COL).Comment Is Nothing Then
                        
                            prvCmnt = msh.Cells(optWiersz, G_DHAS_COL).Comment.Text
                            msh.Cells(optWiersz, G_DHAS_COL).Comment.Delete
                        End If
                        msh.Cells(optWiersz, G_DHAS_COL).AddComment prvCmnt & Chr(10) & _
                            "not added: " & CStr(li.delDate)
                        msh.Cells(optWiersz, G_DHAS_COL).Comment.Shape.TextFrame.AutoSize = True
                        
                        msh.Cells(optWiersz, G_DHAS_COL).Interior.Color = RGB(255, 255, 150)
                    End If
        
                    
                    ' --------------------------------------------------------------------------------
                    ' --------------------------------------------------------------------------------
                End If
            
            
                ' ============================================================================
                ' ============================================================================
                
                Else
                
                    ' transit time is not numeric
                    li.addRawInfoToLog "missing transit time!"
                    msh.Cells(optWiersz, EVO.G_DHEF_COL).Interior.Color = RGB(250, 200, 200)
                    msh.Cells(optWiersz, EVO.G_DHAS_COL).Interior.Color = RGB(250, 200, 200)
                End If
            
        Else
            ' li.isCodEntregaAvailable = false
            msh.Cells(optWiersz, EVO.G_DHEF_COL).Interior.Color = RGB(250, 200, 200)
            msh.Cells(optWiersz, EVO.G_DHAS_COL).Interior.Color = RGB(250, 200, 200)
            
            
            
            
            'msh.Cells(optWiersz, EVO.G_DHEF_COL).Value = ""
            'msh.Cells(optWiersz, EVO.G_DHAS_COL).Value = ""
            
        End If

End Sub


Public Sub init(m_master As Worksheet, m_feed As Worksheet, Optional dwa As E_COPY_HANDLER_INIT, Optional eAnswer As E_ECHANCIER_ONL_semaine_SCENARIO)


    ' just setting private field
    ' e as E_ECHANCIER_ONL_semaine_SCENARIO
    ' just choosing treatment for column in PUS file
    e = eAnswer


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
    
        Dim strTempExpCofor As String, strTempVenCofor As String
        Set expCofors = New Dictionary
    
        Do
            If r.EntireRow.Hidden Then
                ' no operation req for filter out data!
            Else
            
                ' main logic for dwa = 3 internal joke ;)
                ' ------------------------------------------------------------
                
                strTempExpCofor = CStr(CStr(r.Offset(0, EVO.G_COFOR_EXPEDITEUR_COL - 1).Value))
                strTempVenCofor = CStr(CStr(r.Offset(0, EVO.G_COFOR_VENDEUR_COL - 1).Value))
                
                If Not expCofors.Exists(strTempVenCofor & "_" & strTempExpCofor) Then
                
                
                    Set linesIter = Nothing
                    Set linesIter = New Dictionary
                    Set thisLine = New LineItem
                    
                    With thisLine
                        .pn = CStr(r.Value)
                        .cofor = strTempExpCofor
                        .COFOR_COFOR = strTempVenCofor & "_" & strTempExpCofor
                        ' !!!
                        Set .masterRng = r
                        '
                        .wiersz = CLng(r.row)
                        .tmc = CStr(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                        .yyyycw = CLng(operacja.calculateYYYYCW(CStr(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
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
                        
                        .condi = master.Cells(r.row, EVO.G_CONDI_COL).Value
                        ' -------------------------------------------------------------------
                    End With
                    
                    linesIter.Add r.row, thisLine
                    expCofors.Add strTempVenCofor & "_" & strTempExpCofor, linesIter
                
                Else
                
                    ' already exists
                    Set linesIter = expCofors(strTempVenCofor & "_" & strTempExpCofor)
                    Set thisLine = New LineItem
                
                    With thisLine
                        .pn = CStr(r.Value)
                        .cofor = strTempExpCofor
                        .COFOR_COFOR = strTempVenCofor & "_" & strTempExpCofor
                        ' !!!
                        Set .masterRng = r
                        '
                        .wiersz = CLng(r.row)
                        .tmc = CStr(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                        
                        .yyyycw = CLng(operacja.calculateYYYYCW(CStr(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
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
                        
                        .condi = master.Cells(r.row, EVO.G_CONDI_COL).Value
                        ' -------------------------------------------------------------------
                        
                    End With
                    
                    
                    ' just to add inner dictionary
                    ' because linesIter is already a ref inside dictionary!
                    linesIter.Add r.row, thisLine
                
                
                End If
                
                ' ------------------------------------------------------------
            End If
            Set r = r.Offset(1, 0)
        Loop Until CStr(r) = ""
    
        

    ElseIf dwa = E_COPY_HANDLER_BY_TMC_OPT Then
        Dim strTempPn As String
        Dim strTempCoforInner As String
        Dim strTempPnAndCofor As String
        Set pns2 = New Dictionary
        
        Do
        
            If r.EntireRow.Hidden Then
                ' no operation req for filter out data!
            Else
            
                strTempPn = CStr(r.Value)
                strTempCoforInner = CStr(CStr(r.Offset(0, EVO.G_COFOR_VENDEUR_COL - 1).Value))
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
                        .wiersz = CLng(r.row)
                        .tmc = CStr(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                        .yyyycw = CLng(operacja.calculateYYYYCW(CStr(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                        .monday = CDate(operacja.getMondayFromYYYYCW(CLng(.yyyycw)))
                        On Error Resume Next
                        .puDate = CDate(r.Offset(0, EVO.G_DHEF_COL - 1).Value)
                        On Error Resume Next
                        .delDate = CDate(r.Offset(0, EVO.G_DHAS_COL - 1).Value)
                        .isPast = CBool(operacja.isThisLineIsPast2(.puDate))
                        
                    End With
                    
                    linesIter.Add r.row, thisLine
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
                        .wiersz = CLng(r.row)
                        .tmc = CStr(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                        
                        .yyyycw = CLng(operacja.calculateYYYYCW(CStr(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                        .monday = CDate(operacja.getMondayFromYYYYCW(CLng(.yyyycw)))
                        On Error Resume Next
                        .puDate = CDate(r.Offset(0, EVO.G_DHEF_COL - 1).Value)
                        On Error Resume Next
                        .delDate = CDate(r.Offset(0, EVO.G_DHAS_COL - 1).Value)
                        .isPast = CBool(operacja.isThisLineIsPast2(.puDate))
                        
                    End With
                    
                    linesIter.Add r.row, thisLine
                    
                End If
            End If
            
            Set r = r.Offset(1, 0)
        Loop Until CStr(r) = ""

    ElseIf dwa = E_COPY_HANDLER_COPY_ONE Then

        
        Dim strTempCofor As String
        Set cofors = New Dictionary

        Do
        
        
            If r.EntireRow.Hidden Then
                ' no operation req for filter out data!
            Else
        
        
                strTempCofor = CStr(CStr(r.Offset(0, EVO.E_MASTER_cofor - 1).Value))
            
                If Not cofors.Exists(strTempCofor) Then
                
                    Set linesIter = Nothing
                    Set linesIter = New Dictionary
                    
                
                    Set thisLine = New LineItem
                    
                    With thisLine
                        .cofor = strTempCofor
                        .pn = CStr(CStr(r.Offset(0, EVO.E_MASTER_Reference - 1).Value))
                        ' !!!
                        Set .masterRng = r
                        '
                        .wiersz = CLng(r.row)
                        .tmc = CStr(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                        .yyyycw = CLng(operacja.calculateYYYYCW(CStr(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                        .monday = CDate(operacja.getMondayFromYYYYCW(CLng(.yyyycw)))
                        .isPast = CBool(operacja.isThisLineIsPast(CLng(.yyyycw)))
                        
                        ' new one starting from 0.85
                        .tmc2 = CStr(CStr(r.Offset(0, EVO.E_MASTER_3_Echeanciers_regroupés_par_TMC - 1).Value))
                    End With
                    
                    linesIter.Add r.row, thisLine
                    cofors.Add strTempCofor, linesIter
                Else
                
                    ' this cofor is listed - also it means that dictionary inside that cofor as well
                    
                    Set linesIter = cofors(strTempCofor)
                    
                    
                    Set thisLine = New LineItem
                    
                    With thisLine
                        .cofor = strTempCofor
                        .pn = CStr(CStr(r.Offset(0, EVO.E_MASTER_Reference - 1).Value))
                        ' !!!
                        Set .masterRng = r
                        '
                        .wiersz = CLng(r.row)
                        .tmc = CStr(CStr(r.Offset(0, EVO.E_MASTER_TMC - 1).Value))
                        .yyyycw = CLng(operacja.calculateYYYYCW(CStr(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
                        .monday = CDate(operacja.getMondayFromYYYYCW(CLng(.yyyycw)))
                        .isPast = CBool(operacja.isThisLineIsPast(CLng(.yyyycw)))
                        
                        
                        ' new one starting from 0.85
                        .tmc2 = CStr(CStr(r.Offset(0, EVO.E_MASTER_3_Echeanciers_regroupés_par_TMC - 1).Value))
                        
                    End With
                    
                    
                    linesIter.Add r.row, thisLine
                    
                    
                End If
            End If
        
            Set r = r.Offset(1, 0)
        Loop Until CStr(r) = ""
        
        
        
        
        ' afer all check math
        Debug.Print "count cofors: " & cofors.count
        
        ' check all lines
        Dim xl As Integer, k As Variant
        xl = 0
        For Each k In cofors.Keys
            
            Set linesIter = cofors(k)
            
            Debug.Print "for cofor: " & k & " we have: " & linesIter.count & " lines"
            xl = xl + linesIter.count
            
        Next
        
        ' -1 for loop -2 for labels
        Debug.Print "checksum on all lines: " & xl & " rows in file: " & (r.row - 1 - 2)
    
        
        ' for now without this checking!
        'If CLng(xl) = CLng(r.Row - 1 - 2) Then
        '    '  OK
        'Else
        '    MsgBox "Checksum on data is not right - there is sth not right with the BASE worksheet!", vbCritical
        'End If
    End If
End Sub




