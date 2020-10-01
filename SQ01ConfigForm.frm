VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SQ01ConfigForm 
   Caption         =   "SQ01"
   ClientHeight    =   3510
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
Public coll As Collection



Private Sub BtnSubmit_Click()
    innerSubmitDry
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    innerSubmitDry
End Sub


Private Sub innerSubmitDry()
    hide
    
    Set coll = New Collection
    
    If Me.ListBox1.MultiSelect = fmMultiSelectMulti Then
    
        Dim x As Variant
        For x = 0 To Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(x) Then
                coll.Add CStr(Me.ListBox1.list(x))
            End If
        Next x
        
    ElseIf Me.ListBox1.MultiSelect = fmMultiSelectSingle Then
        whatYouChoose = ListBox1.Value
    End If
    
    
    Me.ListBox1.MultiSelect = fmMultiSelectSingle
End Sub
