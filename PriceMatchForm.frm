VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PriceMatchForm 
   Caption         =   "PRICE MATCH"
   ClientHeight    =   1560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "PriceMatchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PriceMatchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnFromExternalFile_Click()


    ' Debug.Print "tp04Match!"
    Me.hide
    
    
    Dim w As Workbook, tmpCaption As String
    
    With FileChooser
    
        tmpCaption = .LabelForSecFile.Caption
        .LabelForSecFile.Caption = "Export from TP04"
    
        .scenarioType = E_FORM_SCENATIO_PRICE_MATCHING_FOR_TP04
        .BtnCopy.Enabled = False
        .BtnValid.Enabled = True
        
        .ComboBoxFeed.Clear
        .ComboBoxMaster.Clear
        
        
        For Each w In Application.Workbooks
            .ComboBoxFeed.addItem w.name
            .ComboBoxMaster.addItem w.name
        Next w
        
        
        .show
    End With
    
    MsgBox "GOTOWE!", vbInformation
End Sub

Private Sub BtnSq01Data_Click()
    ' Debug.Print "tp04Match!"
    Me.hide
    
    
    innerSq01
    
    MsgBox "GOTOWE!", vbInformation
End Sub
