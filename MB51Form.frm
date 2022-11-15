VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MB51Form 
   Caption         =   "MB51"
   ClientHeight    =   1290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9495
   OleObjectBlob   =   "MB51Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MB51Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddLineBtn_Click()

    Dim cs As Controls
    Set cs = Me.Controls
    
    ' Debug.Print cs.Count
    
    
    Dim tbxMag As Control
    Dim tbxDu As Control
    Dim tbxAu As Control
    Dim tbxMvt1 As Control
    Dim tbxMvt2 As Control
    
    
    If cs.count = 12 Then
        ' just a begining
        
        Set tbxMag = cs.Add("Forms.TextBox.1", "TextBoxMag02", True)
        tbxMag.Top = 36 + 18
        tbxMag.Left = 6
        
        Set tbxMag = cs.Add("Forms.TextBox.1", "TextBoxDu02", True)
        tbxMag.Top = 36 + 18
        tbxMag.Left = 84
        
        tbxMag.Value = Format(Date - 14, "dd.mm.yyyy")
        
        
        Set tbxMag = cs.Add("Forms.TextBox.1", "TextBoxAu02", True)
        tbxMag.Top = 36 + 18
        tbxMag.Left = 162
        
        tbxMag.Value = Format(Date, "dd.mm.yyyy")
        
        
        Set tbxMvt1 = cs.Add("Forms.TextBox.1", "TextBoxMvt1_02", True)
        tbxMvt1.Top = 36 + 18
        tbxMvt1.Left = 240
        
        tbxMvt1.Value = "101"
        
        
        Set tbxMvt2 = cs.Add("Forms.TextBox.1", "TextBoxMvt2_02", True)
        tbxMvt2.Top = 36 + 18
        tbxMvt2.Left = 318
        
        tbxMvt2.Value = "102"
        
        
        
        Me.Height = Me.Height + 40
    Else
    
        Dim howManyLinesAlready As Integer
        howManyLinesAlready = ((cs.count - 12) / 5) + 1
        
        If howManyLinesAlready < 10 Then
        
            
            Set tbxMag = cs.Add("Forms.TextBox.1", "TextBoxMag0" & CStr(howManyLinesAlready + 1), True)
            tbxMag.Top = 36 + 18 * (howManyLinesAlready)
            tbxMag.Left = 6
            
            Set tbxMag = cs.Add("Forms.TextBox.1", "TextBoxDu0" & CStr(howManyLinesAlready + 1), True)
            tbxMag.Top = 36 + 18 * (howManyLinesAlready)
            tbxMag.Left = 84
            
            tbxMag.Value = Format(Date - 14, "dd.mm.yyyy")
            
            Set tbxMag = cs.Add("Forms.TextBox.1", "TextBoxAu0" & CStr(howManyLinesAlready + 1), True)
            tbxMag.Top = 36 + 18 * howManyLinesAlready
            tbxMag.Left = 162
            
            tbxMag.Value = Format(Date, "dd.mm.yyyy")
            
            Set tbxMvt1 = cs.Add("Forms.TextBox.1", "TextBoxMvt1_0" & CStr(howManyLinesAlready + 1), True)
            tbxMvt1.Top = 36 + 18 * howManyLinesAlready
            tbxMvt1.Left = 240
            
            tbxMvt1.Value = "101"
            
            Set tbxMag = cs.Add("Forms.TextBox.1", "TextBoxMvt2_0" & CStr(howManyLinesAlready + 1), True)
            tbxMag.Top = 36 + 18 * howManyLinesAlready
            tbxMag.Left = 318
            
            tbxMag.Value = "102"
            
            Me.Height = Me.Height + 18
        Else
            MsgBox "MAX INPUT: 9 lines!"
        End If
    End If
End Sub

Private Sub SubmitBtn_Click()

    hide

    Dim c As Control
    Dim cs As Controls
    Set cs = Me.Controls
    
    
    Dim i_mb51 As MB51_InputItem
    
    
    Dim d As New Dictionary
    ' key will be number from textbox
    
    Dim enumItem As Long
    enumItem = 1
    
    Dim key As String
    For Each c In cs
    
        If c.name Like "TextBox*" Then
    
            key = Right(c.name, 2)
            
            If Not d.Exists(key) Then
                
                Set i_mb51 = New MB51_InputItem
                tryToAddValueInto i_mb51, c
                
                If i_mb51.mag <> "" Then
                    d.Add key, i_mb51
                End If
            Else
                Set i_mb51 = d(key)
                If i_mb51.mag <> "" Then
                    tryToAddValueInto i_mb51, c
                End If
            End If
        End If
        
    Next c
    
    
    runMainMB51Logic d, False
End Sub


Private Sub tryToAddValueInto(ByRef o As MB51_InputItem, ByRef c As Control)
    
    If c.name Like "TextBoxMag*" Then
        o.mag = CStr(c.Value)
    ElseIf c.name Like "TextBoxDu*" Then
        o.du = CStr(c.Value)
    ElseIf c.name Like "TextBoxAu*" Then
        o.au = CStr(c.Value)
    ElseIf c.name Like "TextBoxMvt1*" Then
        o.mvt1 = CStr(c.Value)
    ElseIf c.name Like "TextBoxMvt2*" Then
        o.mvt2 = CStr(c.Value)
    End If
End Sub
