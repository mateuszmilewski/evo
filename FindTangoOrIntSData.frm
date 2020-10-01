VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FindTangoOrIntSData 
   Caption         =   "Find Tango Data"
   ClientHeight    =   1245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4365
   OleObjectBlob   =   "FindTangoOrIntSData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FindTangoOrIntSData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnMatch_Click()

    If checkInputs(ActiveSheet, ThisWorkbook.Sheets(Me.ComboBox1.Value)) Then
    
        hide
    
        If Me.Caption Like "*TANGO*" Then
            runMatchingLogicOnTango ActiveSheet, ThisWorkbook.Sheets(Me.ComboBox1.Value)
        Else
            runMatchingLogicOnInternalSuppliers ActiveSheet, ThisWorkbook.Sheets(Me.ComboBox1.Value)
        End If
    Else
        MsgBox "Worksheets that you choose are in wrong standard!", vbCritical
    End If
End Sub


Private Function checkInputs(sh As Worksheet, interrocomData As Worksheet) As Boolean
    checkInputs = False
    
    If Me.Caption Like "*TANGO*" Then
        If UCase(interrocomData.name) Like "INTERROCOM_*" Then
            If sh.name Like "TP04*" Or sh.name Like "MB51*" Then
                checkInputs = True
            End If
        End If
    Else
    
        If sh.name Like "TP04*" Or sh.name Like "MB51*" Then
            ' fake name interrocom can be interrocom and internal supplier as well
            ' but quality DRY
            If UCase(interrocomData.name) Like "N_*" Then
                checkInputs = True
            End If
        End If
    End If
End Function
