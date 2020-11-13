VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ComboFormReceptionReport 
   Caption         =   "ComboFormReceptionReport"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13995
   OleObjectBlob   =   "ComboFormReceptionReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ComboFormReceptionReport"
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

Private currentStep As E_RECEPTION_REPORT_STEP
Private internalSuppliersSheet As Worksheet
Private mb51_output As Worksheet
Private tango_output As Worksheet
Private lean_tango As Worksheet
Private managersDaSh As Worksheet

' poniewaz mam nielichy problem z logika dzialania
' gdu textbox Mag jest pusty, to ogranicze silowo
' aby wlasciwie ograniczyc


Public Sub innerAddLine()
    

    Dim cs As Controls
    Set cs = Me.Controls
    
    ' Debug.Print cs.Count
    
    
    Dim tbxMag As Control
    Dim tbxDu As Control
    Dim tbxAu As Control
    Dim tbxMvt1 As Control
    Dim tbxMvt2 As Control
    
    

    
    Dim howManyLinesAlready As Integer
    ' howManyLinesAlready = ((cs.Count - 12) / 5) + 1
    howManyLinesAlready = 0
    Dim c As Control
    For Each c In cs
    
        If c.name Like "TextBoxMag*" Then
            howManyLinesAlready = howManyLinesAlready + 1
        End If
    Next c
        
    
    ' artificial limitation - only two possible!
    If howManyLinesAlready < 3 Then
        
            
        Set tbxMag = cs.Add("Forms.TextBox.1", "TextBoxMag0" & CStr(howManyLinesAlready + 1), True)
        tbxMag.Top = 54 + 18 * (howManyLinesAlready)
        tbxMag.Left = 6
        tbxMag.Value = ""
        
        Set tbxDu = cs.Add("Forms.TextBox.1", "TextBoxDu0" & CStr(howManyLinesAlready + 1), True)
        tbxDu.Top = 54 + 18 * (howManyLinesAlready)
        tbxDu.Left = 84
        tbxDu.Value = Me.TextBoxDu01.Value
        
        Set tbxAu = cs.Add("Forms.TextBox.1", "TextBoxAu0" & CStr(howManyLinesAlready + 1), True)
        tbxAu.Top = 54 + 18 * howManyLinesAlready
        tbxAu.Left = 162
        tbxAu.Value = Me.TextBoxAu01.Value
        
        
        
        Set tbxMvt1 = cs.Add("Forms.TextBox.1", "TextBoxMvt1_0" & CStr(howManyLinesAlready + 1), True)
        tbxMvt1.Top = 54 + 18 * howManyLinesAlready
        tbxMvt1.Left = 240
        
        tbxMvt1.Value = "101"
        
        Set tbxMvt2 = cs.Add("Forms.TextBox.1", "TextBoxMvt2_0" & CStr(howManyLinesAlready + 1), True)
        tbxMvt2.Top = 54 + 18 * howManyLinesAlready
        tbxMvt2.Left = 318
        tbxMvt2.Value = "102"
        
        Me.Height = Me.Height + 18
    End If

End Sub


Private Sub AddLineBtn_Click()

    Dim c As Control
    Set c = Nothing
    On Error Resume Next
    Set c = Me.Controls("TextBoxMag02")
        
    If c Is Nothing Then innerAddLine

End Sub

Private Sub ComboBoxPRE_DEF_Change()


    Dim cs As Controls
    Set cs = Me.Controls
    
    Dim tbxMag As Control
    Dim c As Control


    Dim tmp As String
    tmp = CStr(Me.ComboBoxPRE_DEF.Value)
    
    Dim rr As Range
    Set rr = ThisWorkbook.Sheets("register").Range("AD2")
    
    Do
        If CStr(rr.Offset(0, 1).Value) = CStr(tmp) Then
            On Error Resume Next
            Me.TextBoxMag01.Value = rr.Offset(0, 2).Value
            'On Error Resume Next
            'Me.TextBoxMag02.Value = rr.Offset(0, 3).Value
            For Each c In cs
            
                If c.name = "TextBoxMag02" Then
                    c.Value = rr.Offset(0, 3).Value
                    Exit For
                End If
            Next c
            
            
            ' P81  for P2QO for example!
            Me.TxtBoxPricePattern.Value = rr.Offset(0, 4).Value
            Me.TxtBoxProjectNameAlias.Value = CStr(tmp)
            
            Exit Do
            
        End If
        Set rr = rr.Offset(1, 0)
    Loop Until Trim(rr.Value) = ""
    
End Sub


Private Sub SelectAll_Click()
    Me.TextBoxYYYYCW.Value = "*"
End Sub

Private Sub SubmitBtn_Click()

    hide
    
    Set internalSuppliersSheet = Nothing
    Set mb51_output = Nothing
    Set lean_tango = Nothing
    Set managersDaSh = Nothing
    Set EVO.GlobalSapModule.sapGuiAuto = Nothing
    
    
    Dim alias As String
    alias = Me.TxtBoxProjectNameAlias.Value
    Dim divForInterrocom As String
    divForInterrocom = Me.TxtBoxPricePattern.Value
    
    
    Dim v As New Validator
    If v.checkIfComboFormIsFilledProperly(Me) Then
    
    
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
        
    
        Dim c As Control
        Dim cs As Controls
        Set cs = Me.Controls
        
        
        Dim i_mb51 As MB51_InputItem
        
        
        Dim d As New Dictionary
        ' key will be number from textbox
        
        Dim enumItem As Long
        enumItem = 1
        
        Dim key As String
        For Each c In cs
        
            If c.name Like "TextBox*" Then
        
                key = Right(c.name, 2)
                
                If Not d.Exists(key) Then
                    
                    Set i_mb51 = New MB51_InputItem
                    tryToAddValueInto i_mb51, c
                    
                    d.Add key, i_mb51
                Else
                    Set i_mb51 = d(key)
                    tryToAddValueInto i_mb51, c
                End If
            End If
            
        Next c
        
        Dim ans As Variant
        
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
        
        currentStep = E_RECEPTION_REPORT_STEP_GET_DATA_FROM_MB51
        runMainMB51Logic d, True, mb51_output
        
        
        'sepcial place after main logic because main logic for mb51 provide feed for managers da
        ' managers da new one!
        If Me.ComboBoxManagersDA.Value = "" Then
            'MsgBox "There is no source for managers da fields... tool need extra time for downloading", vbInformation
            'ans = MsgBox("You are sure you want to make it? Maybe check again if there is any MANAGERS_DA_ worksheet with internal suppliers list already...", vbInformation + vbYesNo)
            ans = vbYes
            If ans = vbYes Then
                
                ' mb51_output
                innerGetManagersDa mb51_output, managersDaSh
            Else
                End
            End If
        Else
            Debug.Print "You choose internal suppliers source - no need to donwload it again!"
            Set managersDaSh = ThisWorkbook.Sheets(CStr(Me.ComboBoxManagersDA.Value))
        End If
        
        fillReceptionManagersDaColumn mb51_output, managersDaSh
        
        mb51_output.Activate
        
        currentStep = E_RECEPTION_REPORT_STEP_GET_INTERNAL_SUPPLIERS
        runMatchingLogicOnInternalSuppliers mb51_output, internalSuppliersSheet
        
        currentStep = E_RECEPTION_REPORT_STEP_MATCH_WITH_INTERROCOM
        runMatchingLogicOnTango mb51_output, lean_tango, True, divForInterrocom
        
        currentStep = E_RECEPTION_REPORT_STEP_FINAL_TOUCH
        innerFinalTouchOnReceptionReport mb51_output, True, Me.TextBoxYYYYCW.Value
        
        
        MsgBox "Ready!"
    Else
        MsgBox "Hey! What are you doing!? PLeAse fill inpput data correctly!", vbQuestion
    End If
    
End Sub

Private Sub closeExternalFile(extSh As Worksheet)

    Dim wrk As Workbook
    Set wrk = extSh.Parent
    
    wrk.Close False
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


Private Sub tryToAddValueInto(ByRef o As MB51_InputItem, ByRef c As Control)
    
    If c.name Like "TextBoxMag*" Then
        o.mag = CStr(c.Value)
    ElseIf c.name Like "TextBoxDu*" Then
        o.du = CStr(c.Value)
    ElseIf c.name Like "TextBoxAu*" Then
        o.au = CStr(c.Value)
    ElseIf c.name Like "TextBoxMvt1*" Then
        o.mvt1 = CStr(c.Value)
    ElseIf c.name Like "TextBoxMvt2*" Then
        o.mvt2 = CStr(c.Value)
    End If
End Sub

Private Sub UserForm_Deactivate()
    End
End Sub
