VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StatusForm 
   Caption         =   "Status"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9390
   OleObjectBlob   =   "StatusForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StatusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' jedyna metoda przerwania jesli chodzi o ten form
' normalnie ma pracowac przez caly czas bez bolu
Private Sub BtnPrzerwij_Click()
    Application.EnableEvents = True
    End
End Sub

