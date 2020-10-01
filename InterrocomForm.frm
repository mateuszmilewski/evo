VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InterrocomForm 
   Caption         =   "Get Interrocom file"
   ClientHeight    =   1935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5970
   OleObjectBlob   =   "InterrocomForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InterrocomForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InterrocomBtn_Click()
    
    
    ' check if value in box is valid
    
    If notValidInterrocomFile(Me.ComboBox1.Value) Then
    
        MsgBox "Chosen file is not valid!", vbExclamation
    Else
        hide
        
        Debug.Print "Chosen file is valid!"
        
        ' inside interrocom module
        'validation ok, so go with logic
        Dim sh As Worksheet
        Set sh = Workbooks(Me.ComboBox1.Value).ActiveSheet
        ok_runInterrocomAdjustment sh, Nothing, ""
    End If
    
    ThisWorkbook.Activate
End Sub




Private Sub OpenFile_Click()

    Dim sh1 As Worksheet
    Set sh1 = tryToFindProperInterrocomFileThen()
    
    If Not sh1 Is Nothing Then
        Me.ComboBox1.addItem sh1.name
        Me.ComboBox1.Value = sh1.Parent.name
        ThisWorkbook.Activate
    End If
End Sub


Private Function tryToFindProperInterrocomFileThen() As Worksheet
    Set tryToFindProperInterrocomFileThen = Nothing
    
    ' ==================================================================
    Dim strFile As String, file As Workbook
    
    strFile = CStr(Application.GetOpenFilename(, , "Get file with Interrocom standard!", "GET INTERROCOM", False))
    
    If strFile <> "" Then
        
        Set file = Workbooks.Open(strFile)
        '
        Do
            DoEvents
        Loop While file Is Nothing
        ''
        
        Set tryToFindProperInterrocomFileThen = file.ActiveSheet
        
        If butAreYouInInterrocomStandardQuestionMark(tryToFindProperInterrocomFileThen) Then
            
        Else
            Set tryToFindProperInterrocomFileThen = Nothing
        End If
    End If
    ' ==================================================================
End Function



Private Function butAreYouInInterrocomStandardQuestionMark(sh As Worksheet) As Boolean
    butAreYouInInterrocomStandardQuestionMark = False
    
    ' from interrocom module - dry - kinda
    butAreYouInInterrocomStandardQuestionMark = Not notValidInterrocomFile(sh.Parent.name)
End Function
