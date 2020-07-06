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

Public Sub getSq01DataWithPreDefParams(ictrl As IRibbonControl)
    getDataFromSq01WithPreDefinedParams
End Sub


Public Sub getDataFromSq01WithPreDefinedParams()
    
    ' PRE_DEF_RUN_FOR_SQ01
    Dim refReg As Range, ctrl As TextBox, x As Variant
    Set refReg = ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01")
    
    With SQ01PreDefForm
        For x = 1 To 5
            On Error Resume Next
            .Controls("TextBox" & CStr(x)).Value = refReg.Offset(0, x - 1).Value
        Next x
        
        .show
    End With
    
    ' try to run with those predefs
    innerMainForSq01 True
End Sub

Public Sub getDataFromSq01(ictrl As IRibbonControl)
    
    ' Debug.Print "Welcome in SQ01 scope!"
    
    innerMainForSq01 False
End Sub


Private Sub innerMainForSq01(Optional preDef As Boolean)


    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    
    Dim numHandler As NumberHandler
    Set numHandler = New NumberHandler


    Dim st_h As StatusHandler, xStHelper As Integer
    Set st_h = New StatusHandler
    st_h.init_statusbar 20
    st_h.show
    
    
    delegacjaDlaProgresu st_h, xStHelper, 20
    
    
    Dim ish As Worksheet, osh As Worksheet, irng As Range, orng As Range
    Set ish = ThisWorkbook.Sheets.Add
    Set osh = ThisWorkbook.Sheets.Add
    ish.name = EVO.TryToRenameModule.tryToRenameWorksheet(ish, "input_sq01")
    osh.name = EVO.TryToRenameModule.tryToRenameWorksheet(osh, "ouput_sq01")
    
    
    ' LABLES ------------------------------------------------
    
    fillLabels osh.Range("A1")
    
    ' -------------------------------------------------------
    
    Dim sap__handler As New SAP_Handler
    If preDef Then
        sap__handler.runMainLogicForSQ01__with_preDef E_SQ01_CONFIG_MAKE_ALL, ish, osh, st_h, xStHelper
    Else
        sap__handler.runMainLogicForSQ01 E_SQ01_CONFIG_MAKE_ALL, ish, osh, st_h, xStHelper
    End If
    
    ' COPY AND PASTE AS VALUES ------------------------------
    
    copyAndPasteAsValues osh.Range("A1").Offset(0, E_FROM_SQ01_QUASI_TP04_ARTICLE - 1)
    
    ' -------------------------------------------------------
    
    
    
    st_h.hide
    Set st_h = Nothing
    
    
    
    Set numHandler = Nothing
    
    
    
    
'    Dim answer As Variant
'    answer = MsgBox("Raw output from SQ01 (quasi TP04) ready! Do you want to continue?", vbYesNo + vbQuestion)
'
'    If answer = vbYes Then
'        runAdjusterForDataFromSq01 osh
'    End If
    
    
    MsgBox "GOTOWE!", vbInformation
    
    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
End Sub



Private Sub copyAndPasteAsValues(refRng As Range)


    Dim allRange As Range
    
    If refRng.Offset(1, 0).Value <> "" Then
    
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


Private Sub fillLabels(labelRefRange As Range)


    Dim refLabelInRegister As Range, x As Variant
    Set refLabelInRegister = ThisWorkbook.Sheets("forValidation").Range(G_REF_MOUNT_SQ1_OUT)


    With labelRefRange
        
        For x = EVO.E_FROM_SQ01_QUASI_TP04_DOMAIN To EVO.E_FROM_SQ01_QUASI_TP04_CURRENCY
            
            .Offset(0, x - 1).Value = refLabelInRegister.Offset(0, x - 1).Value
        Next x
        
        
    End With
    
End Sub
