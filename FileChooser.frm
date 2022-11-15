VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileChooser 
   Caption         =   "Define your files"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9210
   OleObjectBlob   =   "FileChooser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileChooser"
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




Public scenarioType As E_FORM_SCENARIO_TYPE
Public tmpCaption1 As String
Public tmpCaption2 As String



Private Sub BtnCopy_Click()
    hide
    
    
    If Me.scenarioType = E_FORM_SCENARIO_COPY_PASTE Then
        innerCopyData Me.ComboBoxMaster.Value, Me.ComboBoxFeed.Value
    ElseIf Me.scenarioType = E_FORM_SCENARIO_CREATE_PIVOT_SCENARIO Then
        innerCreateSourceForPivot Me.ComboBoxMaster.Value, Me.ComboBoxFeed.Value
    ElseIf Me.scenarioType = E_FORM_SCENARIO_KPI Then
        innerAfterFormCreateKPI Me.ComboBoxMaster.Value, Me.ComboBoxFeed.Value
    ElseIf Me.scenarioType = E_FORM_SCENATIO_PRICE_MATCHING_FOR_TP04 Then
        innerAfterTP04Logic Me.ComboBoxMaster.Value, Me.ComboBoxFeed.Value
    ElseIf Me.scenarioType = E_FORM_SCENATIO_PRICE_MATCHING_FOR_SQ01 Then
        innerAfterSQ01Logic Me.ComboBoxMaster.Value, Me.ComboBoxFeed.Value
    ElseIf Me.scenarioType = E_FORM_SCENARIO_CLEAR_COLORS Then
        ' nop
    ElseIf Me.scenarioType = E_FORM_SCENARIO_CLEAR_COMMENTS Then
        ' nop
    ElseIf Me.scenarioType = E_FORM_SCENARIO_OPT_BY_BB Then
        ' nop
    ElseIf Me.scenarioType = E_FORM_SCENARIO_OPT_BY_TMC Then
        ' nop
    Else
        MsgBox "Not possible - no such scenario defined!"
    End If
End Sub

Private Sub BtnValid_Click()

    ' some basic validation first
    
    
    Dim v As Validator
    Set v = New Validator
    
    
    Dim answer As Boolean
    answer = True
    
    
    If (Me.scenarioType = E_FORM_SCENARIO_CLEAR_COLORS) Or (Me.scenarioType = E_FORM_SCENARIO_CLEAR_COMMENTS) Then
    
    
        answer = v.getWorksheetForValidation(Me.ComboBoxMaster.Value, E_MASTER_PUS)
        
        'MsgBox "Chosen files valideted! OK!", vbInformation
        Me.BtnCopy.Enabled = True
        
        If answer Then
            ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M1").Value = Me.ComboBoxMaster.Value
        Else
            MsgBox "Chosen files are not in standard!", vbCritical
        End If
    
    ElseIf (Me.scenarioType = E_FORM_SCENARIO_OPT_BY_BB) Or (Me.scenarioType = E_FORM_SCENARIO_OPT_BY_TMC) Then
    
        answer = v.getWorksheetForValidation(Me.ComboBoxMaster.Value, E_MASTER_PUS)
        
        'MsgBox "Chosen files valideted! OK!", vbInformation
        Me.BtnCopy.Enabled = True
        
        If answer Then
            ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M1").Value = Me.ComboBoxMaster.Value
        Else
            MsgBox "Chosen files are not in standard!", vbCritical
        End If
        
    ElseIf Me.scenarioType = E_FORM_SCENATIO_PRICE_MATCHING_FOR_SQ01 Then
    
        answer = answer And v.getWorksheetForValidation(Me.ComboBoxFeed.Value, E_SQ01)
        answer = answer And v.getWorksheetForValidation(Me.ComboBoxMaster.Value, E_MASTER_PUS)
        
        
        If answer Then
            MsgBox "Chosen files valideted! OK!", vbInformation
            Me.BtnCopy.Enabled = True
            
            
            ' put just names of those workbooks for next subs
            ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M1").Value = Me.ComboBoxMaster.Value
            ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M2").Value = Me.ComboBoxFeed.Value
        Else
            MsgBox "Chosen files are not in standard!", vbCritical
        End If
        
    ElseIf Me.scenarioType = E_FORM_SCENATIO_PRICE_MATCHING_FOR_TP04 Then
    
    
        answer = answer And v.getWorksheetForValidation(Me.ComboBoxFeed.Value, E_TP04_01)
        answer = answer And v.getWorksheetForValidation(Me.ComboBoxMaster.Value, E_MASTER_PUS)
        
        If answer Then
            'MsgBox "Chosen files valideted! OK!", vbInformation
            Me.BtnCopy.Enabled = True
            
            
            ' put just names of those workbooks for next subs
            ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M1").Value = Me.ComboBoxMaster.Value
            ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M2").Value = Me.ComboBoxFeed.Value
        Else
            MsgBox "Chosen files are not in standard!", vbCritical
        End If
    
        
    ElseIf Me.scenarioType < E_FORM_SCENATIO_PRICE_MATCHING_FOR_TP04 Then
    
        
        
        answer = answer And v.getWorksheetForValidation(Me.ComboBoxFeed.Value, E_FEED_CPL)
        answer = answer And v.getWorksheetForValidation(Me.ComboBoxMaster.Value, E_MASTER_PUS)
        
        If answer Then
            'MsgBox "Chosen files valideted! OK!", vbInformation
            Me.BtnCopy.Enabled = True
            
            
            ' put just names of those workbooks for next subs
            ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M1").Value = Me.ComboBoxMaster.Value
            ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M2").Value = Me.ComboBoxFeed.Value
        Else
            MsgBox "Chosen files are not in standard!", vbCritical
        End If
    
    Else
        MsgBox "Not such scenario to validation!", vbCritical
    End If
    
End Sub

