VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModelessLeaf 
   Caption         =   "Modeless Leaf"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   OleObjectBlob   =   "ModelessLeaf.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModelessLeaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public repsh As Worksheet
Public outWrk As Workbook




Public Sub Setup(sh1 As Worksheet)

    Set repsh = sh1
    
    
    TextBoxSource.Value = repsh.name
    
    
    Dim scopeDictionary As New Dictionary
    Dim rRef As Range, key As Variant
    
    If repsh.name Like "GREEN_LIGHT_*" Then
        
        ' green light approach
        Set rRef = repsh.Cells(2, 1)
        Do
            
            key = rRef.Offset(0, EVO.E_GREEN_LIGHT_ECHANCIER_ONL_semaine - 1).Value
            If Not scopeDictionary.Exists(key) Then
                scopeDictionary.Add key, 1
            End If
            
            Set rRef = rRef.Offset(1, 0)
        Loop Until Trim(rRef.Value) = ""
        
        
    ElseIf repsh.name Like "RECEPTION_*" Then
    
        ' reception approach
        Set rRef = repsh.Cells(2, 1)
        Do
            key = rRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sem - 1).Value
            If Not scopeDictionary.Exists(key) Then
                scopeDictionary.Add key, 1
            End If
            Set rRef = rRef.Offset(1, 0)
        Loop Until Trim(rRef.Value) = ""
    End If
    
    
    Me.ListBoxScope.Clear
    For Each key In scopeDictionary.Keys
        Me.ListBoxScope.addItem key
    Next
End Sub

Private Sub ExportBtn_Click()

    
    
    ' put into new workbook
    Dim mlh As New ModelessLeafHandler
    If repsh.name Like "GREEN_LIGHT_*" Then
        mlh.setMode True, False, Me
    ElseIf repsh.name Like "RECEPTION_*" Then
        mlh.setMode False, True, Me
    End If
    
    mlh.createLeaf
End Sub

Private Sub ListBoxScope_Click()
    

        Dim pnCount As New Dictionary
        
        Dim internalCount As New Dictionary, costInternal As Double
        Dim countNoTango As New Dictionary, costNoTango As Double
        
        Dim countTango As New Dictionary
        Dim countTangoOK As New Dictionary
        Dim countTangoNOK As New Dictionary
        Dim calcOKNOK As Double
        Dim costTango As Double, costTarget As Double, costGap As Double
        
        
        Dim key As Variant
        
        
        Dim rRef As Range
        Set rRef = repsh.Cells(2, 1)
        
        Do

            If repsh.name Like "GREEN_LIGHT_*" Then
            
                
                If rRef.Offset(0, EVO.E_GREEN_LIGHT_ECHANCIER_ONL_semaine - 1).Value = Me.ListBoxScope.Value Then
                
                    key = rRef.Offset(0, EVO.E_GREEN_LIGHT_Reference - 1).Value
                    If Not pnCount.Exists(key) Then pnCount.Add key, 1
                    
                    ' internal
                    If rRef.Offset(0, EVO.E_GREEN_LIGHT_IS_INTERNAL - 1).Value = "internal" Then
                        If Not internalCount.Exists(key) Then internalCount.Add key, 1
                        
                        costInternal = costInternal + CDbl(rRef.Offset(0, EVO.E_GREEN_LIGHT_Spending_sigapp - 1).Value)
                    Else
                        
                        ' most important scope
                        ' --------------------------------------------------------------
                        If rRef.Offset(0, EVO.E_GREEN_LIGHT_TANGO_OKNOK - 1).Value = "NO TANGO PRICE" Then
                        
                            ' scope only for no tango price PNs
                            If Not countNoTango.Exists(key) Then countNoTango.Add key, 1
                            costNoTango = costNoTango + CDbl(rRef.Offset(0, EVO.E_GREEN_LIGHT_Spending_sigapp - 1).Value)
                        Else
                        
                            
                            
                            If Not countTango.Exists(key) Then countTango.Add key, 1
                            
                            costTango = costTango + CDbl(rRef.Offset(0, EVO.E_GREEN_LIGHT_Spending_sigapp - 1).Value)
                            costTarget = costTarget + CDbl(rRef.Offset(0, EVO.E_GREEN_LIGHT_Spending_Target - 1).Value)
                            
                        
                            If rRef.Offset(0, EVO.E_GREEN_LIGHT_TANGO_OKNOK - 1).Value = "OK" Then
                                If Not countTangoOK.Exists(key) Then countTangoOK.Add key, 1
                            ElseIf rRef.Offset(0, EVO.E_GREEN_LIGHT_TANGO_OKNOK - 1).Value = "NOK" Then
                                If Not countTangoNOK.Exists(key) Then countTangoNOK.Add key, 1
                            End If
                        End If
                        
                        ' --------------------------------------------------------------
                    End If
                    
                End If
                
            
            ElseIf repsh.name Like "RECEPTION_*" Then
            
                     
                If rRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sem - 1).Value = Me.ListBoxScope.Value Then
                
                
                    key = rRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_article - 1).Value
                    If Not pnCount.Exists(key) Then pnCount.Add key, 1
                    
                    ' internal
                    If rRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Interne - 1).Value = "internal" Then
                        If Not internalCount.Exists(key) Then internalCount.Add key, 1
                        
                        costInternal = costInternal + CDbl(rRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sigapp - 1).Value)
                    Else
                        
                        ' all section which is not internal
                        ' most important scope
                        ' --------------------------------------------------------------
                        
                        
                        If Trim(rRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_Tango - 1).Value) = "" Then
                        
                            ' scope only for no tango price PNs
                            If Not countNoTango.Exists(key) Then countNoTango.Add key, 1
                            costNoTango = costNoTango + CDbl(rRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sigapp - 1).Value)
                        Else
                        
                            calcOKNOK = CDbl(rRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Ecart - 1).Value)
                        
                            
                            If Not countTango.Exists(key) Then countTango.Add key, 1
                            
                            costTango = costTango + CDbl(rRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sigapp - 1).Value)
                            costTarget = costTarget + CDbl(rRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_cible - 1).Value)
                            
                        
                            If calcOKNOK < 1.1 Then
                                If Not countTangoOK.Exists(key) Then countTangoOK.Add key, 1
                            Else
                                If Not countTangoNOK.Exists(key) Then countTangoNOK.Add key, 1
                            End If
                        End If
                        ' --------------------------------------------------------------
                    End If
                    
                    
                End If
            
            End If
            
            
            
            Set rRef = rRef.Offset(1, 0)
            
        Loop Until Trim(rRef.Value) = ""
        
        Me.TextBoxCPN.Value = ""
        Me.TextBoxCPN.Value = pnCount.Count
        Me.TextBox_CountInternal.Value = ""
        Me.TextBox_CountInternal.Value = internalCount.Count
        
        Me.TextBox_CostInternal.Value = ""
        Me.TextBox_CostInternal.Value = costInternal
        
        Me.TextBox_CountNoTango.Value = countNoTango.Count
        Me.TextBox_CostNoTango.Value = costNoTango
        
        
        
        Me.TextBox_CostTango.Value = costTango
        Me.TextBox_CountTango = countTangoOK.Count + countTangoNOK.Count
        Me.TextBox_CountTangoNOK = countTangoNOK.Count
        
        Me.TextBox_CostTarget.Value = costTarget
        
        If costTarget > 0 Then Me.TextBox_RATE = CDbl(costTango / costTarget)
        
        ' final gap
        costGap = costTango - costTarget
        Me.TextBox_CostGap.Value = costGap
        
End Sub
