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

Private sapGuiAuto As Object
Private sapGui As SAPFEWSELib.GuiApplication
Private cnn As SAPFEWSELib.GuiConnection
Private sess As SAPFEWSELib.GuiSession
Private window As SAPFEWSELib.GuiModalWindow
Private c As SAPFEWSELib.GuiCollection
Private item As SAPFEWSELib.GuiComponent
Private btn As SAPFEWSELib.GuiButton
Private table As SAPFEWSELib.GuiTableControl ' ISapTableControlTarget
Private grid As SAPFEWSELib.GuiGridView
Private gv As SAPFEWSELib.GuiGridView
Private myTableRow As SAPFEWSELib.GuiTableRow
Private coll As SAPFEWSELib.GuiCollection
    
    
Private chbx As SAPFEWSELib.GuiCheckBox
Private txt As SAPFEWSELib.GuiTextedit


' !!
Private cols As Variant

' params for table inside sap
Private y As Variant
Private x As Variant
Private x2 As Variant
Private x3 As Variant

Private Sub Class_Initialize()
    Set sapGuiAuto = GetObject("SAPGUI")
    Set sapGui = sapGuiAuto.GetScriptingEngine
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
End Sub

Private Sub Class_Terminate()
    Set sapGuiAuto = Nothing
    Set sapGui = Nothing
    Set cnn = Nothing
    Set sess = Nothing
    Set item = Nothing
End Sub


Public Sub runMainLogicForSQ01__with_preDef(mcfg1 As E_SQ01_CONFIG, ByRef ish As Worksheet, ByRef osh As Worksheet, ByRef st_h As StatusHandler, ByRef xStHelper As Integer)


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
        
        
        runLoopsWithPreDef "wnd[1]/usr/cntlGRID1/shellcont/shell", E_SQ01_PREDEF_SYSTEM
        
        runLoopsWithPreDef "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", E_SQ01_PREDEF_LIST1
        
        
        sess.FindById("wnd[0]/tbar[1]/btn[17]").Press
        sess.FindById("wnd[1]").SendVKey 4
        
        runLoopsWithPreDef "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell", E_SQ01_PREDEF_STD1
        
        
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
                    refInput.Value = CStr(grid.GetCellValue(x, cols(y)))
                    refInput.Offset(0, 1).Value = CStr(grid.GetCellValue(x, cols(y + 1)))
                    refInput.Offset(0, 2).Value = CStr(grid.GetCellValue(x, cols(y + 2)))
                    Exit For
                End If
                
            Next y
            
            
            

            If x Mod 100 = 0 Then
            
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
        
        runLoopsWithPreDef "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", E_SQ01_PREDEF_LIST2
        
        
        sess.FindById("wnd[0]/tbar[1]/btn[17]").Press
        ' otwieramy liste z dostepnymi standardami do etapu drugiego!
        sess.FindById("wnd[1]").SendVKey 4
        
        runLoopsWithPreDef "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell", E_SQ01_PREDEF_STD2
        
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
    
    
    Dim refOutput2 As Range  ' , lm As Long
    Set refOutput2 = osh.Cells(2, 1)
    
    ' lm = CLng((ish.Range("A1").End(xlDown).row) / 50)
    
    For x = 0 To grid.RowCount - 1
        For y = 0 To grid.ColumnCount - 1
        
        
            refOutput2.Offset(0, y).Value = CStr(grid.GetCellValue(x, cols(y)))
        Next y
        
        If x Mod 50 = 0 Then
        
            grid.FirstVisibleRow = x
            grid.CurrentCellRow = x
            
            ' delegacjaDlaProgresu st_h, xStHelper, 20 ' obsolete
            st_h.progress_increase
            
        End If
        
        
        Set refOutput2 = refOutput2.Offset(1, 0)
    Next x
    
End Sub



Public Sub runMainLogicForSQ01(mcfg1 As E_SQ01_CONFIG, ByRef ish As Worksheet, ByRef osh As Worksheet, ByRef st_h As StatusHandler, ByRef xStHelper As Integer)


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
                    refInput.Value = CStr(grid.GetCellValue(x, cols(y)))
                    refInput.Offset(0, 1).Value = CStr(grid.GetCellValue(x, cols(y + 1)))
                    refInput.Offset(0, 2).Value = CStr(grid.GetCellValue(x, cols(y + 2)))
                    Exit For
                End If
                
            Next y
            
            
            

            If x Mod 100 = 0 Then
            
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
        
        
            refOutput2.Offset(0, y).Value = CStr(grid.GetCellValue(x, cols(y)))
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
    
    If cfg1 = 1 Then
        ' pierwsze wybieranie!
        ' -----------------------------------------------------
        
        For x = 0 To grid.RowCount - 1
            SQ01ConfigForm.ListBox1.AddItem CStr(grid.GetCellValue(x, cols(0)))
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
        If CStr(wynik) = CStr(grid.GetCellValue(x, cols(0))) Then
        
            With grid
                .CurrentCellRow = x
                .SelectedRows = "" & CStr(x)
                .DoubleClickCurrentCell
            End With
            
            Exit For
            
        End If
    Next x

End Sub


Private Sub runLoopsWithPreDef(stringReference As String, preDefEnum As E_SQ01_PREDEF)


    Dim wynik As String

    Set grid = sess.FindById(CStr(stringReference))
    Set cols = grid.ColumnOrder
    
    wynik = getPreDef(preDefEnum)
    
    For x = 0 To grid.RowCount - 1
        If CStr(wynik) = CStr(grid.GetCellValue(x, cols(0))) Then
        
            With grid
                .CurrentCellRow = x
                .SelectedRows = "" & CStr(x)
                .DoubleClickCurrentCell
            End With
            
            Exit For
            
        End If
    Next x

End Sub


Private Function getPreDef(preDefEnum As E_SQ01_PREDEF) As String

    If preDefEnum = E_SQ01_PREDEF_SYSTEM Then
        
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").Value
    ElseIf preDefEnum = E_SQ01_PREDEF_LIST1 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").Offset(0, 1).Value
    ElseIf preDefEnum = E_SQ01_PREDEF_STD1 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").Offset(0, 2).Value
    ElseIf preDefEnum = E_SQ01_PREDEF_LIST2 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").Offset(0, 3).Value
    ElseIf preDefEnum = E_SQ01_PREDEF_STD2 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").Offset(0, 4).Value
    Else
        MsgBox "Fatal error in pre-def logic inside foo getPreDef!", vbCritical
        End
    End If
End Function