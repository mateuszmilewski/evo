VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ECHANCIER_ONL_CW_TREAT_FORM 
   Caption         =   "ECHANCIER ONL (semaine) treatment"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ECHANCIER_ONL_CW_TREAT_FORM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ECHANCIER_ONL_CW_TREAT_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public whatYouChoose As E_ECHANCIER_ONL_semaine_SCENARIO

Private Sub SubmitBtn_Click()
    hide
    
    
    If Me.OptionButtonDEL.Value Then
        whatYouChoose = E_ECHANCIER_ONL_semaine_SCENARIO_DEL
    ElseIf Me.OptionButtonPU.Value Then
        whatYouChoose = E_ECHANCIER_ONL_semaine_SCENARIO_PU
    End If
End Sub
