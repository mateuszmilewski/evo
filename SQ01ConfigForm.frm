VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SQ01ConfigForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "SQ01ConfigForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SQ01ConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public whatYouChoose As String

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    hide
    whatYouChoose = ListBox1.Value

End Sub
