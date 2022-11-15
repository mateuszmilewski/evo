VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PoorForm 
   Caption         =   "80' form for caret position"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5730
   OleObjectBlob   =   "PoorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PoorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListBox1_Click()
    Debug.Print
    
    Dim x As Variant
    For x = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(x) Then
            Me.TextBoxPos.Value = CStr(Int(x) + 1)
        End If
    Next x
End Sub

Private Sub SubmitBtn_Click()
    hide
End Sub
