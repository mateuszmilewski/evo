VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ComboFormGreenLightReport 
   Caption         =   "Combo Form"
   ClientHeight    =   2745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14820
   OleObjectBlob   =   "ComboFormGreenLightReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ComboFormGreenLightReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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


Private currentStep As E_GREEN_LIGHT_REPORT_STEP
Private internalSuppliersSheet As Worksheet
Private mb51_output As Worksheet
Private tango_output As Worksheet
Private lean_tango As Worksheet
Private managersDaSh As Worksheet


' output from sq01
Private sq01_sh1 As Worksheet
Private sq01_sh2 As Worksheet
Private concatSh As Worksheet
Private adjustedSh As Worksheet

Private Sub ComboBoxPRE_DEF_Change()

    Dim cs As Controls
    Set cs = Me.Controls
    
    Dim tbxMag As Control
    Dim c As Control


    Dim tmp As String
    tmp = CStr(Me.ComboBoxPRE_DEF.Value)
    
    Dim rr As Range, sq01PatternRef As Range
    Set rr = ThisWorkbook.Sheets("register").Range("AD2")
    Set sq01PatternRef = ThisWorkbook.Sheets("register").Range("A50")
    
    Do
        If rr.Value = "F" Then
            If CStr(rr.offset(0, 1).Value) = CStr(tmp) Then
                
                With sq01PatternRef
                    Me.TextBox11.Value = .Value
                    Me.TextBox12.Value = .offset(0, 1).Value
                    Me.TextBox13.Value = _
                        Replace(CStr(.offset(0, 2).Value), "XXX", CStr(rr.offset(0, 2).Value))
                    Me.TextBox14.Value = _
                        Replace(CStr(.offset(0, 3).Value), "XXX", CStr(rr.offset(0, 2).Value))
                    Me.TextBox15.Value = _
                        Replace(CStr(.offset(0, 4).Value), "XXX", CStr(rr.offset(0, 2).Value))
                End With
                
                With sq01PatternRef
                    Me.TextBox21.Value = .Value
                    Me.TextBox22.Value = .offset(0, 1).Value
                    Me.TextBox23.Value = _
                        Replace(CStr(.offset(0, 2).Value), "XXX", CStr(rr.offset(0, 3).Value))
                    Me.TextBox24.Value = _
                        Replace(CStr(.offset(0, 3).Value), "XXX", CStr(rr.offset(0, 3).Value))
                    Me.TextBox25.Value = _
                        Replace(CStr(.offset(0, 4).Value), "XXX", CStr(rr.offset(0, 3).Value))
                End With
                
                
                ' P81  for P2QO for example!
                Me.TxtBoxPricePattern.Value = rr.offset(0, 4).Value
                Me.TxtBoxProjectNameAlias.Value = CStr(tmp)
                
                Exit Do
                
            End If
        End If
        
        Set rr = rr.offset(1, 0)
        
    Loop Until Trim(rr.Value) = ""
    
    

    
    ' PRE_DEF_RUN_FOR_SQ01
    
    Application.EnableEvents = False
    
    Dim x As Variant

    For x = 0 To 4
        ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").offset(0, x).Value = _
            Me.Controls("TextBox1" & CStr(x + 1)).Text
           
        On Error Resume Next
        ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").offset(1, x).Value = _
            Me.Controls("TextBox2" & CStr(x + 1)).Text
    Next x
    
    Application.EnableEvents = True

End Sub

Private Sub SubmitBtn_Click()
    
    ' inside main run for the green light report!
    ' ==============================================================
    
    hide
    
    Set sq01_sh1 = Nothing
    Set sq01_sh2 = Nothing
    Set concatSh = Nothing
    Set adjustedSh = Nothing
    
    Set internalSuppliersSheet = Nothing
    Set mb51_output = Nothing
    Set lean_tango = Nothing
    Set managersDaSh = Nothing
    Set EVO.GlobalSapModule.sapGuiAuto = Nothing
    
    
    Dim alias As String, pusName As String, divForInterrocom As String, tp04_ready As String
    alias = Me.TxtBoxProjectNameAlias.Value
    pusName = Me.ComboBoxPUS.Value
    divForInterrocom = Me.TxtBoxPricePattern.Value
    tp04_ready = Me.ComboBoxTP04Ready.Value
    
    
    Dim v As New Validator
    If v.checkIfComboFormIsFilledProperly(Me, "GREEN_LIGHT") Then
    
        If Me.ComboBoxTangoSource.Value = "" Then
        
        
            ' open file dialog and take proper iterrocom file!
            Dim ans2 As Variant
            
            ans2 = MsgBox("You do not choose any for interrocom files for EVO, do you want to find a fresh external file for report?", vbInformation + vbYesNo)
            
            
            If ans2 = vbYes Then
                
                Set tango_output = Nothing
            
                Set tango_output = tryToFindProperInterrocomFileThen()
                
                If tango_output Is Nothing Then
                
                    MsgBox "NO! You can't run report without Tango data in proper standard!", vbCritical
                    End
                Else
                
                    ok_runInterrocomAdjustment tango_output, lean_tango, alias
                    
                    closeExternalFile tango_output
                    
                    ThisWorkbook.Activate
                End If
            End If
        Else
            ' just assign nicely to variable
            Set lean_tango = ThisWorkbook.Sheets(CStr(Me.ComboBoxTangoSource.Value))
        End If
        
        Dim ans As Variant
        
        Dim c As Control
        Dim cs As Controls
        Set cs = Me.Controls
        
        
        Dim i_sq01 As SQ01_InputItem
        
        
        Dim d As New Dictionary
        ' key will be number from textbox
        
        Dim enumItem As Long
        enumItem = 1
        
        Dim key As String
        For Each c In cs
        
            If c.name Like "TextBox??" Then
        
                key = Left(c.name, 8)
                
                If Not d.Exists(key) Then
                    
                    Set i_sq01 = New SQ01_InputItem
                    tryToAddValueInto i_sq01, c
                    
                    d.Add key, i_sq01
                Else
                    Set i_sq01 = d(key)
                    tryToAddValueInto i_sq01, c
                End If
            End If
            
        Next c
        
        
        
        ' this test went just fine OK
        ' ------------------------------------
        'Dim k2 As Variant
        'For Each k2 In d.Keys
        '    Debug.Print d(k2).standard2
        'Next
        ' ------------------------------------
        
        If Me.ComboBoxInternalSupplier.Value = "" Then
            'MsgBox "You do not choose any worksheet with internal supplier data... tool need extra time for downloading it from sigapp now", vbInformation
            'ans = MsgBox("You are sure you want to make it? Maybe check again if there is any N_ worksheet with internal suppliers list already...", vbInformation + vbYesNo)
            ans = vbYes
            If ans = vbYes Then
                isolatedLogicForInternalSuppliers internalSuppliersSheet
            Else
                End
            End If
        Else
            Debug.Print "You choose internal suppliers source - no need to donwload it again!"
            Set internalSuppliersSheet = ThisWorkbook.Sheets(CStr(Me.ComboBoxInternalSupplier.Value))
        End If
        
        
    
        
        ThisWorkbook.Activate
        
        If Me.ComboBoxTP04Ready.Value = "" Then
        
            currentStep = E_GREEN_LIGHT_REPORT_STEP_GET_SQ01
            innerMainForSq01 True, Me.TextBox13.Value, Me.TextBox23.Value, sq01_sh1, sq01_sh2, True
        
            currentStep = E_GREEN_LIGHT_REPORT_STEP_CONCAT
            concatAndStd sq01_sh1, sq01_sh2, concatSh
        Else
            
            Set concatSh = ThisWorkbook.Sheets(Me.ComboBoxTP04Ready.Value)
        End If
        
        concatSh.Activate
        
        currentStep = E_GREEN_LIGHT_REPORT_STEP_ADJUST
        innerAfterSQ01Logic pusName, concatSh.name, adjustedSh
        
        adjustedSh.Activate
        
        'sepcial place after main logic because main logic for mb51 provide feed for managers da
        ' managers da new one!
        If Me.ComboBoxManagersDA.Value = "" Then
            'MsgBox "There is no source for managers da fields... tool need extra time for downloading", vbInformation
            'ans = MsgBox("You are sure you want to make it? Maybe check again if there is any MANAGERS_DA_ worksheet with internal suppliers list already...", vbInformation + vbYesNo)
            ans = vbYes
            If ans = vbYes Then
                
                ' mb51_output
                innerGetManagersDa adjustedSh, managersDaSh
            Else
                End
            End If
        Else
            Debug.Print "You choose managersda source - no need to donwload it again!"
            Set managersDaSh = ThisWorkbook.Sheets(CStr(Me.ComboBoxManagersDA.Value))
        End If
        
        fillGreenLightManagersDaColumn adjustedSh, managersDaSh
        
        adjustedSh.Activate
        
        
        currentStep = E_GREEN_LIGHT_REPORT_STEP_GET_INTERNAL_SUPPLIERS
        runMatchingLogicOnInternalSuppliers adjustedSh, internalSuppliersSheet
        
        
        currentStep = E_GREEN_LIGHT_REPORT_MATCH_WITH_INTERROCOM
        runMatchingLogicOnTango adjustedSh, lean_tango, True, divForInterrocom
        
        
        currentStep = E_GREEN_LIGHT_REPORT_FINAL_TOUCH
        innerFinalTouchOnGreenLightReport adjustedSh, True
        
        ' ==============================================================
    Else
        MsgBox "wrong input!", vbInformation
    End If
End Sub

Private Sub closeExternalFile(extSh As Worksheet)

    Dim wrk As Workbook
    Set wrk = extSh.Parent
    
    wrk.Close False
End Sub

Private Sub TestBtn_Click()
    
    If validatePUSfile() Then
        Me.SubmitBtn.Enabled = True
    Else
        Me.SubmitBtn.Enabled = False
        MsgBox "Validation failure!", vbCritical
    End If
End Sub


Private Function validatePUSfile() As Boolean

    validatePUSfile = False
    Dim v As New Validator
    'Dim sh As Worksheet
    'Set sh = Workbooks(Me.ComboBoxPUS.Value).Sheets("BASE")
    
    Dim ans As Boolean
    ans = False
    If Me.ComboBoxPUS.Value <> "" Then
        ans = v.getWorksheetForValidation(Me.ComboBoxPUS.Value, E_MASTER_PUS)
    End If
    
    validatePUSfile = ans
    
End Function


Private Sub tryToAddValueInto(ByRef o As SQ01_InputItem, ByRef c As Control)
    
    If c.name Like "TextBox?1" Then
        o.group = CStr(c.Value)
    ElseIf c.name Like "TextBox?2" Then
        o.query1 = CStr(c.Value)
    ElseIf c.name Like "TextBox?3" Then
        o.standard1 = CStr(c.Value)
    ElseIf c.name Like "TextBox?4" Then
        o.query2 = CStr(c.Value)
    ElseIf c.name Like "TextBox?5" Then
        o.standard2 = CStr(c.Value)
    End If
End Sub


Private Function tryToFindProperInterrocomFileThen() As Worksheet
    Set tryToFindProperInterrocomFileThen = Nothing
    
    ' ==================================================================
    Dim strFile As String, file As Workbook
    
    strFile = CStr(Application.GetOpenFilename(, , "Get file with Interrocom standard!", "GET INTERROCOM", False))
    
    If strFile <> "" Then
        
        Set file = Workbooks.Open(strFile)
        '
        Do
            DoEvents
        Loop While file Is Nothing
        ''
        
        Set tryToFindProperInterrocomFileThen = file.ActiveSheet
        
        If butAreYouInInterrocomStandardQuestionMark(tryToFindProperInterrocomFileThen) Then
            
        Else
            Set tryToFindProperInterrocomFileThen = Nothing
        End If
    End If
    ' ==================================================================
End Function


Private Function butAreYouInInterrocomStandardQuestionMark(sh As Worksheet) As Boolean
    butAreYouInInterrocomStandardQuestionMark = False
    
    ' from interrocom module - dry - kinda
    butAreYouInInterrocomStandardQuestionMark = Not notValidInterrocomFile(sh.Parent.name)
End Function

