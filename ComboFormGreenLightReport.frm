VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ComboFormGreenLightReport 
   Caption         =   "UserForm1"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14820
   OleObjectBlob   =   "ComboFormGreenLightReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ComboFormGreenLightReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBoxPRE_DEF_Change()

    Dim cs As Controls
    Set cs = Me.Controls
    
    Dim tbxMag As Control
    Dim c As Control


    Dim tmp As String
    tmp = CStr(Me.ComboBoxPRE_DEF.Value)
    
    Dim rr As Range, sq01PatternRef As Range
    Set rr = ThisWorkbook.Sheets("register").Range("AD2")
    Set sq01PatternRef = ThisWorkbook.Sheets("register").Range("A50")
    
    Do
        If rr.Value = "F" Then
            If CStr(rr.Offset(0, 1).Value) = CStr(tmp) Then
                
                With sq01PatternRef
                    Me.TextBox11.Value = .Value
                    Me.TextBox12.Value = .Offset(0, 1).Value
                    Me.TextBox13.Value = _
                        Replace(CStr(.Offset(0, 2).Value), "XXX", CStr(rr.Offset(0, 2).Value))
                    Me.TextBox14.Value = _
                        Replace(CStr(.Offset(0, 3).Value), "XXX", CStr(rr.Offset(0, 2).Value))
                    Me.TextBox15.Value = _
                        Replace(CStr(.Offset(0, 4).Value), "XXX", CStr(rr.Offset(0, 2).Value))
                End With
                
                With sq01PatternRef
                    Me.TextBox21.Value = .Value
                    Me.TextBox22.Value = .Offset(0, 1).Value
                    Me.TextBox23.Value = _
                        Replace(CStr(.Offset(0, 2).Value), "XXX", CStr(rr.Offset(0, 3).Value))
                    Me.TextBox24.Value = _
                        Replace(CStr(.Offset(0, 3).Value), "XXX", CStr(rr.Offset(0, 3).Value))
                    Me.TextBox25.Value = _
                        Replace(CStr(.Offset(0, 4).Value), "XXX", CStr(rr.Offset(0, 3).Value))
                End With
                
                
                ' P81  for P2QO for example!
                Me.TxtBoxPricePattern.Value = rr.Offset(0, 4).Value
                Me.TxtBoxProjectNameAlias.Value = CStr(tmp)
                
                Exit Do
                
            End If
        End If
        
        Set rr = rr.Offset(1, 0)
        
    Loop Until Trim(rr.Value) = ""
End Sub

