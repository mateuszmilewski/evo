VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SQ01PreDefForm 
   Caption         =   "PRE-DEFINED INPUT"
   ClientHeight    =   1830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11700
   OleObjectBlob   =   "SQ01PreDefForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SQ01PreDefForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SaveAndRun_Click()
    
    hide
    
    innerSave
End Sub


Private Sub innerSave()
    
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


