VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FinalScope 
   Caption         =   "Define Scope"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3285
   OleObjectBlob   =   "FinalScope.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FinalScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public c As Collection

Private Sub Btn_Click()

    hide
    Set c = New Collection


    With Me.ListBox1
    
        Dim x As Variant
        For x = 0 To .ListCount - 1
            
            If .Selected(x) Then
                c.Add .list(x)
            End If
        Next x
    End With

End Sub
