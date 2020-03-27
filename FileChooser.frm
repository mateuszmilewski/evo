VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileChooser 
   Caption         =   "Define your files"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5850
   OleObjectBlob   =   "FileChooser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCopy_Click()
    hide
    innerCopyData Me.ComboBoxMaster.Value, Me.ComboBoxFeed.Value
End Sub

Private Sub BtnValid_Click()

    ' some basic validation first
    
    Dim v As Validator
    Set v = New Validator
    
    
    Dim answer As Boolean
    
    answer = True
    
    answer = answer And v.getWorksheetForValidation(Me.ComboBoxFeed.Value, E_FEED_CPL)
    answer = answer And v.getWorksheetForValidation(Me.ComboBoxMaster.Value, E_MASTER_PUS)
    
    If answer Then
        MsgBox "Chosen files valideted! OK!", vbInformation
        Me.BtnCopy.Enabled = True
        
        
        ' put just names of those workbooks for next subs
        ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M1").Value = Me.ComboBoxMaster.Value
        ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("M2").Value = Me.ComboBoxFeed.Value
    Else
        MsgBox "Chosen files are not in standard!", vbCritical
    End If
    
End Sub
