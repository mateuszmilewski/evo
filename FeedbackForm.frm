VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FeedbackForm 
   Caption         =   "Feedback Form - choose Reception report"
   ClientHeight    =   1275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "FeedbackForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FeedbackForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public shName As String

Private Sub SubmitBtn_Click()
    hide
    shName = Me.ComboBox1.Value
End Sub
