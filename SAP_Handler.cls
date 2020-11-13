VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAP_Handler"
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


' OK - we need to change it into late binging becuase EVO users
' might do not have any SAP app

Private sapGuiAuto As Object
'Private sapGui As SAPFEWSELib.GuiApplication
'Private cnn As SAPFEWSELib.GuiConnection
'Private sess As SAPFEWSELib.GuiSession
'Private window As SAPFEWSELib.GuiModalWindow
'Private c As SAPFEWSELib.GuiCollection
'Private item As SAPFEWSELib.GuiComponent
'Private Btn As SAPFEWSELib.GuiButton
'Private table As SAPFEWSELib.GuiTableControl ' ISapTableControlTarget
'Private grid As SAPFEWSELib.GuiGridView
'Private gv As SAPFEWSELib.GuiGridView
'Private myTableRow As SAPFEWSELib.GuiTableRow
'Private coll As SAPFEWSELib.GuiCollection

'Private chbx As SAPFEWSELib.GuiCheckBox
'Private txt As SAPFEWSELib.GuiTextedit
    
Private sapGui As Object
Private cnn As Object
Private sess As Object
Private window As Object
Private c As Object
Private item As Object
Private Btn As Object
Private table As Object
Private grid As Object
Private gv As Object
Private myTableRow As Object
Private coll As Object
    
Private chbx As Object
Private txt As Object


' !!
Private cols As Variant

' params for table inside sap
Private y As Variant
Private x As Variant
Private x2 As Variant
Private x3 As Variant




Private Sub Class_Initialize()

    If EVO.GlobalSapModule.sapGuiAuto Is Nothing Then
        Set sapGuiAuto = GetObject("SAPGUI")
        Set EVO.GlobalSapModule.sapGuiAuto = sapGuiAuto
        Set sapGui = sapGuiAuto.GetScriptingEngine
        Set EVO.GlobalSapModule.sapGui = sapGui
    Else
        Set sapGuiAuto = EVO.GlobalSapModule.sapGuiAuto
        Set sapGui = EVO.GlobalSapModule.sapGui
    End If
    
    Set cnn = sapGui.Connections(0)
    
    Debug.Print cnn.ConnectionString
    Debug.Print cnn.Sessions.count

    
    Set sess = cnn.Children(0)
    Set item = sess.Children(0)
    Debug.Print item.name
    
    Debug.Print sess.Children.count
    
    
    ' Set item = sess.Children(0)
    Set item = sess.FindById("wnd[0]/usr")
    Debug.Print item.Children.count

    
    
    sess.FindById("wnd[0]").Maximize
End Sub

Private Sub Class_Terminate()
    Set sapGuiAuto = Nothing
    Set sapGui = Nothing
    Set cnn = Nothing
    Set sess = Nothing
    Set item = Nothing
End Sub

Public Sub runMainLogicFor__Y_PI1_80000391(sh1 As Worksheet, ByRef st_h As StatusHandler, ByRef xStHelper As Integer)


    Application.Calculation = xlCalculationManual
    
    ' we just taking all the data from this transaction - nothing special!"
    ' im not proud of this "hack"
    ' ----------------------------------------------------
    Dim x17 As Variant
    For x17 = 0 To 10
        On Error Resume Next
        sess.FindById("wnd[0]/tbar[0]/btn[12]").Press
    Next x17
    ' ----------------------------------------------------
    
    
    sess.FindById("wnd[0]/tbar[0]/okcd").Text = "Y_PI1_80000391"
    sess.FindById("wnd[0]").SendVKey 0
    sess.FindById("wnd[0]/tbar[1]/btn[8]").Press
    
    'sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetCurrentCell 53, "LFA1-NAME1"
    'sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").FirstVisibleRow = 31
    'sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectedRows = "53"
    
    Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    Set cols = grid.ColumnOrder
    
    
    Dim refOutput2 As Range  ' , lm As Long
    Set refOutput2 = sh1.Cells(2, 1)
    
    ' lm = CLng((ish.Range("A1").End(xlDown).row) / 50)
    
    For x = 0 To grid.RowCount - 1
        For y = 0 To grid.ColumnCount - 1
        
        
            refOutput2.Offset(0, y).Value = CStr(grid.getCellValue(x, cols(y)))
        Next y
        
        If x Mod 50 = 0 Then
        
            grid.FirstVisibleRow = x
            grid.CurrentCellRow = x
            
            ' delegacjaDlaProgresu st_h, xStHelper, 20 ' obsolete
            ' st_h.progress_increase
            
        End If
        
        
        Set refOutput2 = refOutput2.Offset(1, 0)
    Next x
    
    Application.Calculation = xlCalculationAutomatic
End Sub


Public Sub runMainLogicForSQ01__with_preDef(mcfg1 As E_SQ01_CONFIG, ByRef ish As Worksheet, ByRef osh As Worksheet, ByRef st_h As StatusHandler, ByRef xStHelper As Integer, whichLineForInput As Integer)


    ' from now on 2 possibilities
    
    Application.Calculation = xlCalculationManual


    If mcfg1 = E_SQ01_CONFIG_START_FROM_PASTING_ALREADY_PREPARED_INPUT_LIST Then
        
        ' avoiding initial stuff
    Else
    
        ' no need to clear anything becuase we just created new worksheet
        ' clearProduitFromInput ish = obsolete
        
        ' im not proud of this "hack"
        ' ----------------------------------------------------
        Dim x17 As Variant
        For x17 = 0 To 10
            On Error Resume Next
            sess.FindById("wnd[0]/tbar[0]/btn[12]").Press
        Next x17
        ' ----------------------------------------------------
    
    
        sess.FindById("wnd[0]/tbar[0]/okcd").Text = "SQ01"
        sess.FindById("wnd[0]").SendVKey 0
        sess.FindById("wnd[0]/mbar/menu[1]/menu[7]").Select
        
        'sess.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").SelectedRows = "0"
        'sess.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").DoubleClickCurrentCell
        
        
        runLoopsWithPreDef "wnd[1]/usr/cntlGRID1/shellcont/shell", E_SQ01_PREDEF_SYSTEM, whichLineForInput
        
        runLoopsWithPreDef "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", E_SQ01_PREDEF_LIST1, whichLineForInput
        
        
        sess.FindById("wnd[0]/tbar[1]/btn[17]").Press
        sess.FindById("wnd[1]").SendVKey 4
        
        runLoopsWithPreDef "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell", E_SQ01_PREDEF_STD1, whichLineForInput
        
        
        sess.FindById("wnd[1]/tbar[0]/btn[0]").Press
        
        ' wynik !
        ' z tego pobieramy liste czesci!
        sess.FindById("wnd[0]/tbar[1]/btn[8]").Press
        
        
        ' teraz glowny wrzut na input
        ' =========================================================================
        
        Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        Set cols = grid.ColumnOrder
        
        
        Dim refInput As Range
        Set refInput = ish.Range("A1")
        
        refInput.Offset(0, 0).Value = "PRODUIT"
        refInput.Offset(0, 1).Value = "DESIGNATION"
        refInput.Offset(0, 2).Value = "RU"
        
        Set refInput = refInput.Offset(1, 0)
        
        
        
        ' MAIN LOOP FOR FIRST LIST OF PN WITHOUT INDICE
        For x = 0 To grid.RowCount
            For y = 0 To grid.ColumnCount
            
            
                If CStr(cols(y)) Like "*_PRODUIT*" Then
                
                    ' Debug.Print Cstr(grid.GetCellValue(x2, cols(y2))) ' OK
                    refInput.Value = CStr(grid.getCellValue(x, cols(y)))
                    refInput.Offset(0, 1).Value = CStr(grid.getCellValue(x, cols(y + 1)))
                    refInput.Offset(0, 2).Value = CStr(grid.getCellValue(x, cols(y + 2)))
                    Exit For
                End If
                
            Next y
            
            
            

            If x Mod 50 = 0 Then
            
                grid.FirstVisibleRow = x
                grid.CurrentCellRow = x
                
                
                ' powazny side effect ...
                delegacjaDlaProgresu st_h, xStHelper, 20
                
            End If
            
            
            Set refInput = refInput.Offset(1, 0)
        Next x
        
        makeDoubleListWithDashAndStar ish.Range("A2")
        
        st_h.hide
        Set st_h = Nothing
        Set st_h = New StatusHandler
        st_h.init_statusbar CLng((ish.Range("A1").End(xlDown).row) / 50)
        st_h.show
        st_h.progress_increase
        
        ' go back! x2
        sess.FindById("wnd[0]/tbar[0]/btn[15]").Press
        sess.FindById("wnd[0]/tbar[0]/btn[15]").Press
        
        
        ' teraz etap 2
        ' mamy powiekszona liste czesci - musimy przejsc jeszcze raz przez wybor
        
        ' tym razem wybieramy konkretnie
        
        runLoopsWithPreDef "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", E_SQ01_PREDEF_LIST2, whichLineForInput
        
        
        sess.FindById("wnd[0]/tbar[1]/btn[17]").Press
        ' otwieramy liste z dostepnymi standardami do etapu drugiego!
        sess.FindById("wnd[1]").SendVKey 4
        
        sess.FindById("wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "0"
        sess.FindById("wnd[2]/tbar[0]/btn[2]").Press
        
        ' runLoopsWithPreDef "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell", E_SQ01_PREDEF_STD2, whichLineForInput
        session.FindById("wnd[1]/usr/ctxtRS38R-VARIANT").Text = getProperRefFromPreDef(E_SQ01_PREDEF_STD2, whichLineForInput)
        
        sess.FindById("wnd[1]/tbar[0]/btn[0]").Press
        
        
        ' teraz powinnismy byc na oknie ponownego forma wejsciowego dla listy quasi tp04
        
        
    End If
    
    Dim rngToCopy As Range
    Dim refOutput2 As Range  ' , lm As Long
    Set refOutput2 = osh.Cells(2, 1)
    
    ' A1048576 - always last one
    Set rngToCopy = tryToAssignFirstBatchOfData(ish)
    
    Do
    
    
        With sess
        
        
            ' Debug.Print TypeName(.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH"))
            
            
            Dim btn3 As Variant ' SAPFEWSELib.GuiButton
            Set btn3 = findMyButton() ' .FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH")
            ' simple solution - you need to take 3rd button from top
            'Debug.Print btn.Top
            
            'Set btn = .FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH")
            'Debug.Print btn.Top
            '
            'Set btn = .FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH")
            'Debug.Print btn.Top
            '
            'Set btn = .FindById("wnd[0]/usr/btn%_SP$00004_%_APP_%-VALU_PUSH")
            'Debug.Print btn.Top
            
    
            
            ' Code de suppression dans le
            '.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").Press
            ' please do not delete data from this field!
            ' .FindById("wnd[1]/tbar[0]/btn[16]").Press
            '.FindById("wnd[1]/tbar[0]/btn[0]").Press
            '.FindById("wnd[1]/tbar[0]/btn[8]").Press
        
            '.FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").Press
            '.FindById("wnd[1]/tbar[0]/btn[16]").Press
            '.FindById("wnd[1]/tbar[0]/btn[0]").Press
            '.FindById("wnd[1]/tbar[0]/btn[8]").Press
            
            ' Numéro d'article
            '.FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").Press
            '.FindById("wnd[1]/tbar[0]/btn[16]").Press
            
            btn3.Press
            .FindById("wnd[1]/tbar[0]/btn[16]").Press
               
            ' now we making it in the loop - to avoid OLE data missing issue
            ' Set rngToCopy = ish.Range(ish.Range("A2"), ish.Range("A2").End(xlDown))
            
            ' Debug.Print rngToCopy.Address
            rngToCopy.Copy
            
            .FindById("wnd[1]/tbar[0]/btn[24]").Press
            .FindById("wnd[1]/tbar[0]/btn[8]").Press
            .FindById("wnd[0]/tbar[1]/btn[8]").Press
        End With
    
        
        Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        Set cols = grid.ColumnOrder
        
        
        
        ' refOutput2
        
        If refOutput2.row > 2 Then
            ' so at least next iteration
            
            ' Debug.Print refOutput2.row
            
            
            ' very simple logic
            ' be careful - almost without any additional checking
            If Trim(refOutput2.Value) = "" Then
                Set refOutput2 = refOutput2.Offset(-1, 0)
            End If
            
            
        End If

        
        ' lm = CLng((ish.Range("A1").End(xlDown).row) / 50)
        
        For x = 0 To grid.RowCount - 1
            For y = 0 To grid.ColumnCount - 1
            
            
                refOutput2.Offset(0, y).Value = CStr(grid.getCellValue(x, cols(y)))
            Next y
            
            If x Mod 50 = 0 Then
            
                grid.FirstVisibleRow = x
                grid.CurrentCellRow = x
                
                ' delegacjaDlaProgresu st_h, xStHelper, 20 ' obsolete
                st_h.progress_increase
                
            End If
            
            
            Set refOutput2 = refOutput2.Offset(1, 0)
        Next x
        
        
        'step back in SIGAPP
        sess.FindById("wnd[0]/tbar[0]/btn[3]").Press
        ' new range to copy
        Set rngToCopy = tryToAssignNewRange(ish, rngToCopy)
        
        
        
        
        
    Loop Until youDryOutRangeOfDataForSq01(rngToCopy)
    
    st_h.hide
    
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Private Function tryToAssignFirstBatchOfData(sh As Worksheet) As Range
    Set tryToAssignFirstBatchOfData = sh.Range("A2:A500")
End Function



Private Function tryToAssignNewRange(sh As Worksheet, data As Range) As Range
    Set tryToAssignNewRange = Nothing
    
    
    Dim tmp As Range
    
    Set tmp = data.Offset(499, 0)
    
    ' Debug.Print tmp.Address
    
    Set tryToAssignNewRange = tmp
End Function

Private Function youDryOutRangeOfDataForSq01(data As Range) As Boolean

    youDryOutRangeOfDataForSq01 = False
    
    Dim ir As Range, i As Integer, tmp As Range
    
    i = 0
    For Each ir In data
        
        If Trim(ir.Value) <> "" Then
            i = i + 1
        End If
    Next ir
    
    If i > 0 Then
        youDryOutRangeOfDataForSq01 = False
    Else
        youDryOutRangeOfDataForSq01 = True
    End If
End Function

Private Function findMyButton() As Variant ' SAPFEWSELib.GuiButton
    
    Set findMyButton = Nothing
    
    ' top
    '51
    '73
    '95
    '29
    
    Dim tmpButton As Variant ' SAPFEWSELib.GuiButton
    
    Dim x As Variant
    For x = 1 To 4
        ' Set btn = .FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH")
        Set tmpButton = sess.FindById("wnd[0]/usr/btn%_SP$0000" & CStr(x) & "_%_APP_%-VALU_PUSH")
        
        If (tmpButton.Top > 70 And tmpButton.Top < 80) Or (tmpButton.Top = 73) Then
            ' this is it
            Set findMyButton = tmpButton
            Exit For
        End If
    Next x
End Function

Public Sub runMainLogicForSQ01(mcfg1 As E_SQ01_CONFIG, ByRef ish As Worksheet, ByRef osh As Worksheet, ByRef st_h As StatusHandler, ByRef xStHelper As Integer, whichLineForInput As Integer)


    If mcfg1 = E_SQ01_CONFIG_START_FROM_PASTING_ALREADY_PREPARED_INPUT_LIST Then
        
        ' avoiding initial stuff
    Else
    
        ' no need to clear anything becuase we just created new worksheet
        ' clearProduitFromInput ish = obsolete
        
        ' im not proud of this "hack"
        ' ----------------------------------------------------
        Dim x17 As Variant
        For x17 = 0 To 10
            On Error Resume Next
            sess.FindById("wnd[0]/tbar[0]/btn[12]").Press
        Next x17
        ' ----------------------------------------------------
    
    
        sess.FindById("wnd[0]/tbar[0]/okcd").Text = "SQ01"
        sess.FindById("wnd[0]").SendVKey 0
        sess.FindById("wnd[0]/mbar/menu[1]/menu[7]").Select
        
        'sess.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").SelectedRows = "0"
        'sess.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").DoubleClickCurrentCell
        
        
        runLoops "wnd[1]/usr/cntlGRID1/shellcont/shell"
        runLoops "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell"
        
        
        sess.FindById("wnd[0]/tbar[1]/btn[17]").Press
        sess.FindById("wnd[1]").SendVKey 4
        
        runLoops "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell"
        
        
        sess.FindById("wnd[1]/tbar[0]/btn[0]").Press
        
        ' wynik !
        ' z tego pobieramy liste czesci!
        sess.FindById("wnd[0]/tbar[1]/btn[8]").Press
        
        
        ' teraz glowny wrzut na input
        ' =========================================================================
        
        Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        Set cols = grid.ColumnOrder
        
        
        Dim refInput As Range
        Set refInput = ish.Range("A1")
        
        refInput.Offset(0, 0).Value = "PRODUIT"
        refInput.Offset(0, 1).Value = "DESIGNATION"
        refInput.Offset(0, 2).Value = "RU"
        
        Set refInput = refInput.Offset(1, 0)
        
        
        
        ' MAIN LOOP FOR FIRST LIST OF PN WITHOUT INDICE
        For x = 0 To grid.RowCount
            For y = 0 To grid.ColumnCount
            
            
                If CStr(cols(y)) Like "*_PRODUIT*" Then
                
                    ' Debug.Print Cstr(grid.GetCellValue(x2, cols(y2))) ' OK
                    refInput.Value = CStr(grid.getCellValue(x, cols(y)))
                    refInput.Offset(0, 1).Value = CStr(grid.getCellValue(x, cols(y + 1)))
                    refInput.Offset(0, 2).Value = CStr(grid.getCellValue(x, cols(y + 2)))
                    Exit For
                End If
                
            Next y
            
            
            

            If x Mod 50 = 0 Then
            
                grid.FirstVisibleRow = x
                grid.CurrentCellRow = x
                
                
                ' powazny side effect ...
                delegacjaDlaProgresu st_h, xStHelper, 20
                
            End If
            
            
            Set refInput = refInput.Offset(1, 0)
        Next x
        
        makeDoubleListWithDashAndStar ish.Range("A2")
        
        st_h.hide
        Set st_h = Nothing
        Set st_h = New StatusHandler
        st_h.init_statusbar CLng((ish.Range("A1").End(xlDown).row) / 50)
        st_h.show
        st_h.progress_increase
        
        
        ' go back! x2
        sess.FindById("wnd[0]/tbar[0]/btn[15]").Press
        sess.FindById("wnd[0]/tbar[0]/btn[15]").Press
        
        
        ' teraz etap 2
        ' mamy powiekszona liste czesci - musimy przejsc jeszcze raz przez wybor
        
        ' tym razem wybieramy konkretnie
        
        runLoops "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell"
        
        
        sess.FindById("wnd[0]/tbar[1]/btn[17]").Press
        ' otwieramy liste z dostepnymi standardami do etapu drugiego!
        sess.FindById("wnd[1]").SendVKey 4
        
        runLoops "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell"
        
        sess.FindById("wnd[1]/tbar[0]/btn[0]").Press
        
        
        ' teraz powinnismy byc na oknie ponownego forma wejsciowego dla listy quasi tp04
        
        
    End If
    
    Dim rngToCopy As Range
    
    With sess
        .FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").Press
        .FindById("wnd[1]/tbar[0]/btn[16]").Press
        Set rngToCopy = ish.Range(ish.Range("A2"), ish.Range("A2").End(xlDown))
        rngToCopy.Copy
        
        .FindById("wnd[1]/tbar[0]/btn[24]").Press
        .FindById("wnd[1]/tbar[0]/btn[8]").Press
        .FindById("wnd[0]/tbar[1]/btn[8]").Press
    End With

    
    Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    Set cols = grid.ColumnOrder
    
    
    Dim refOutput2 As Range ', 'lm As Long
    Set refOutput2 = osh.Cells(2, 1)
    
    
    'lm = CLng((ish.Range("A1").End(xlDown).row) / 50)
    
    
    
    For x = 0 To grid.RowCount - 1
        For y = 0 To grid.ColumnCount - 1
        
        
            refOutput2.Offset(0, y).Value = CStr(grid.getCellValue(x, cols(y)))
        Next y
        
        If x Mod 50 = 0 Then
        
            grid.FirstVisibleRow = x
            grid.CurrentCellRow = x
            
            ' delegacjaDlaProgresu st_h, xStHelper, 20
            st_h.progress_increase
            
        End If
        
        
        Set refOutput2 = refOutput2.Offset(1, 0)
    Next x
    
End Sub


Private Function fillFormAndWaitForDecision(cfg1 As Integer) As String


    fillFormAndWaitForDecision = ""
    
    SQ01ConfigForm.ListBox1.Clear
    SQ01ConfigForm.ListBox1.MultiSelect = fmMultiSelectSingle
    
    If cfg1 = 1 Then
        ' pierwsze wybieranie!
        ' -----------------------------------------------------
        
        For x = 0 To grid.RowCount - 1
            SQ01ConfigForm.ListBox1.addItem CStr(grid.getCellValue(x, cols(0)))
        Next x
        
        ' -----------------------------------------------------
    End If
    
    SQ01ConfigForm.show
    
    On Error Resume Next
    fillFormAndWaitForDecision = SQ01ConfigForm.whatYouChoose
End Function




Private Sub makeDoubleListWithDashAndStar(refR As Range)
    
    Dim doubleR As Range, lastOrgRng As Range
    Set doubleR = refR.End(xlDown).End(xlDown).End(xlUp).Offset(1, 0)
    
    If refR.Offset(1, 0) <> "" Then
        
        Set lastOrgRng = refR.End(xlDown)
    Else
        Set lastOrgRng = refR
    End If
    
    Dim ir As Range
    For Each ir In refR.Parent.Range(refR, lastOrgRng)
        doubleR.Value = ir.Value & "-*"
        doubleR.Offset(0, 1).Value = "*-*"
        Set doubleR = doubleR.Offset(1, 0)
    Next ir
End Sub




Private Sub runLoops(stringReference As String)


    Dim wynik As String

    Set grid = sess.FindById(CStr(stringReference))
    Set cols = grid.ColumnOrder
    
    wynik = fillFormAndWaitForDecision(1)
    
    For x = 0 To grid.RowCount - 1
        If CStr(wynik) = CStr(grid.getCellValue(x, cols(0))) Then
        
            With grid
                .CurrentCellRow = x
                .SelectedRows = "" & CStr(x)
                .DoubleClickCurrentCell
            End With
            
            Exit For
            
        End If
    Next x

End Sub


Private Sub runLoopsWithPreDef(stringReference As String, preDefEnum As E_SQ01_PREDEF, preDefLine As Integer)


    Dim wynik As String

    Set grid = sess.FindById(CStr(stringReference))
    Set cols = grid.ColumnOrder
    
    wynik = getPreDef(preDefEnum, preDefLine)
    
    For x = 0 To grid.RowCount - 1
        If CStr(wynik) = CStr(grid.getCellValue(x, cols(0))) Then
        
            With grid
                .CurrentCellRow = x
                .SelectedRows = "" & CStr(x)
                .DoubleClickCurrentCell
            End With
            
            Exit For
            
        End If
    Next x

End Sub

Private Function getProperRefFromPreDef(preDefEnum As E_SQ01_PREDEF, preDefLine As Integer) As String
    getProperRefFromPreDef = getPreDef(preDefEnum, preDefLine)
End Function


Private Function getPreDef(preDefEnum As E_SQ01_PREDEF, m_getPreDef As Integer) As String

    If preDefEnum = E_SQ01_PREDEF_SYSTEM Then
        
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").Offset(m_getPreDef - 1, 0).Value
    ElseIf preDefEnum = E_SQ01_PREDEF_LIST1 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").Offset(m_getPreDef - 1, 1).Value
    ElseIf preDefEnum = E_SQ01_PREDEF_STD1 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").Offset(m_getPreDef - 1, 2).Value
    ElseIf preDefEnum = E_SQ01_PREDEF_LIST2 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").Offset(m_getPreDef - 1, 3).Value
    ElseIf preDefEnum = E_SQ01_PREDEF_STD2 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").Offset(getPreDef - 1, 4).Value
    Else
        MsgBox "Fatal error in pre-def logic inside foo getPreDef!", vbCritical
        End
    End If
End Function
