VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModelessLeaf 
   Caption         =   "Modeless Leaf"
   ClientHeight    =   7980
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
            
            key = rRef.offset(0, EVO.E_GREEN_LIGHT_ECHANCIER_ONL_semaine - 1).Value
            If Not scopeDictionary.Exists(key) Then
                scopeDictionary.Add key, 1
            End If
            
            Set rRef = rRef.offset(1, 0)
        Loop Until Trim(rRef.Value) = ""
        
        
    ElseIf repsh.name Like "RECEPTION_*" Then
    
        ' reception approach
        Set rRef = repsh.Cells(2, 1)
        Do
            key = rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sem - 1).Value
            If Not scopeDictionary.Exists(key) Then
                scopeDictionary.Add key, 1
            End If
            Set rRef = rRef.offset(1, 0)
        Loop Until Trim(rRef.Value) = ""
    End If
    
    
    Me.ListBoxScope.Clear
    
    Set scopeDictionary = funcSortDic(scopeDictionary)
    
    For Each key In scopeDictionary.Keys
        Me.ListBoxScope.addItem key
    Next
    
    
    
    ' also fill active workbooks
    ' --------------------------------
    
    Dim tw As Workbook
    Me.ComboBoxConnectedWith.Clear
    For Each tw In Application.Workbooks
        Me.ComboBoxConnectedWith.addItem tw.name
    Next tw
    
    ' --------------------------------
End Sub




Private Function funcSortDic(d1 As Dictionary) As Dictionary

    
    If d1.count > 1 Then
    
        Dim k As Variant, iter As Long, hcount As Long
        hcount = d1.count
        Dim str() As String
        ReDim str(hcount) As String
        
        
        Dim stringPattern As String
        
        iter = 0
        For Each k In d1.Keys
        
            If iter = 0 Then
                stringPattern = k
            End If
            
            str(iter) = k
            iter = iter + 1
        Next
        
        
        
        If stringPattern Like "??-CW*" Then
            ' special case, becuase frog eaters love to make it overcomplicated:
            ' they made it: 21-CW4 instead of 21-CW04 , so i can't make simple sorting
            ' stpd fckrs...
            str = specialSortArrayAtoZThanksToStpdFrogEaters(str)
        Else
            ' normal sorting - of course not gonna happen
            str = SortArrayAtoZ(str)
        End If
        
        Set d1 = Nothing
        Set d1 = New Dictionary
        
        For iter = 0 To UBound(str)
            
            d1.Add str(iter), 1
        Next iter
        
    Else
        ' nop req
    End If
    
    Set funcSortDic = d1
End Function

Private Function specialSortArrayAtoZThanksToStpdFrogEaters(myArray As Variant)

    Dim i As Long
    Dim j As Long
    Dim Temp
    
    Dim str1 As String, str2 As String
    
    'Sort the Array A-Z
    For i = LBound(myArray) To UBound(myArray) - 1
        For j = i + 1 To UBound(myArray)
        
            If UCase(myArray(i)) Like "??-CW??" Then str1 = UCase(myArray(i))
            If UCase(myArray(i)) Like "??-CW?" Then str1 = Left(UCase(myArray(i)), 5) & "0" & Right(UCase(myArray(i)), 1)
            
            If UCase(myArray(j)) Like "??-CW??" Then str2 = UCase(myArray(j))
            If UCase(myArray(j)) Like "??-CW?" Then str2 = Left(UCase(myArray(j)), 5) & "0" & Right(UCase(myArray(j)), 1)
            
            If str1 > str2 Then
                Temp = myArray(j)
                myArray(j) = myArray(i)
                myArray(i) = Temp
            End If
        Next j
    Next i
    
    specialSortArrayAtoZThanksToStpdFrogEaters = myArray

End Function

Private Function SortArrayAtoZ(myArray As Variant)

    Dim i As Long
    Dim j As Long
    Dim Temp
    
    'Sort the Array A-Z
    For i = LBound(myArray) To UBound(myArray)
        For j = i + 1 To UBound(myArray)
            If UCase(myArray(i)) > UCase(myArray(j)) Then
                Temp = myArray(j)
                myArray(j) = myArray(i)
                myArray(i) = Temp
            End If
        Next j
    Next i
    
    SortArrayAtoZ = myArray

End Function

Private Sub MakeItStepByStepBtn_Click()

    ' put into new workbook
    Dim mlh As New ModelessLeafHandler
    If repsh.name Like "GREEN_LIGHT_*" Then
        mlh.setMode True, False, Me
    ElseIf repsh.name Like "RECEPTION_*" Then
        mlh.setMode False, True, Me
    End If
    
    mlh.createLeaf
End Sub

Private Sub AddBtn_Click()

    ' put into existing workbook (sheet)
    Dim mlh As New ModelessLeafHandler
    If repsh.name Like "GREEN_LIGHT_*" Then
        mlh.setMode True, False, Me
    ElseIf repsh.name Like "RECEPTION_*" Then
        mlh.setMode False, True, Me
    End If
    
    mlh.addLeafToExisitingOne Me.ComboBoxConnectedWith.Value
    
    ' also need to add data to listing,
    ' which should also be available
    ' ----------------------------------------------------------
    
    ' ----------------------------------------------------------
End Sub

Private Sub ExportBtn_Click()

    Set outWrk = Nothing
    
    
    ' put into new workbook
    Dim mlh As New ModelessLeafHandler
    If repsh.name Like "GREEN_LIGHT_*" Then
        mlh.setMode True, False, Me
    ElseIf repsh.name Like "RECEPTION_*" Then
        mlh.setMode False, True, Me
    End If
    
    mlh.createLeafs
End Sub



Private Sub ListBoxScope_Click()
    

        Dim pnCount As New Dictionary
        
        Dim internalCount As New Dictionary, costInternal As Double
        Dim countNoTango As New Dictionary
        
        Dim countTango As New Dictionary
        Dim countTangoOK As New Dictionary
        Dim countTangoNOK As New Dictionary
        Dim calcOKNOK As Double
        Dim costTango As Double, costTarget As Double, costGap As Double
        
        Dim costNoTango As Double, targetNoTango As Double
        
        
        
        ' for rate without tango
        Dim rateSumWithoutTango As Double
        Dim rateCountWithoutTango As Long
        
        
        Dim key As Variant
        
        
        Dim rRef As Range
        Set rRef = repsh.Cells(2, 1)
        
        Do
        
        
            If rRef.EntireRow.Hidden And Me.CheckBoxVisiblity Then
            
            Else
            

                If repsh.name Like "GREEN_LIGHT_*" Then
                
                
                    
                    If rRef.offset(0, EVO.E_GREEN_LIGHT_ECHANCIER_ONL_semaine - 1).Value = Me.ListBoxScope.Value Then
                    
                        key = rRef.offset(0, EVO.E_GREEN_LIGHT_Reference - 1).Value
                        If Not pnCount.Exists(key) Then pnCount.Add key, 1
                        
                        ' internal
                        If rRef.offset(0, EVO.E_GREEN_LIGHT_IS_INTERNAL - 1).Value = "internal" Then
                            If Not internalCount.Exists(key) Then internalCount.Add key, 1
                            
                            costInternal = costInternal + CDbl(rRef.offset(0, EVO.E_GREEN_LIGHT_Spending_sigapp - 1).Value)
                        Else
                            
                            ' most important scope
                            ' --------------------------------------------------------------
                            If rRef.offset(0, EVO.E_GREEN_LIGHT_TANGO_OKNOK - 1).Value = "NO TANGO PRICE" Then
                            
                                ' scope only for no tango price PNs
                                If Not countNoTango.Exists(key) Then countNoTango.Add key, 1
                                costNoTango = costNoTango + CDbl(rRef.offset(0, EVO.E_GREEN_LIGHT_Spending_sigapp - 1).Value)
                                ' targetNoTango = targetNoTango + CDbl(rRef.Offset(0, EVO.E_GREEN_LIGHT_Spending_Target - 1).Value)
                                
                                If CStr(rRef.offset(EVO.E_GREEN_LIGHT_RATE_PRE_SERIAL_div_INIT_PRICE - 1).Value) <> "" Then
                                    rateSumWithoutTango = rateSumWithoutTango + CDbl(rRef.offset(0, EVO.E_GREEN_LIGHT_RATE_PRE_SERIAL_div_INIT_PRICE - 1).Value)
                                    rateCountWithoutTango = rateCountWithoutTango + 1
                                End If
                            Else
                            
                                
                                
                                If Not countTango.Exists(key) Then countTango.Add key, 1
                                
                                costTango = costTango + CDbl(rRef.offset(0, EVO.E_GREEN_LIGHT_Spending_sigapp - 1).Value)
                                costTarget = costTarget + CDbl(rRef.offset(0, EVO.E_GREEN_LIGHT_Spending_Target - 1).Value)
                                
                            
                                If rRef.offset(0, EVO.E_GREEN_LIGHT_TANGO_OKNOK - 1).Value = "OK" Then
                                    If Not countTangoOK.Exists(key) Then countTangoOK.Add key, 1
                                ElseIf rRef.offset(0, EVO.E_GREEN_LIGHT_TANGO_OKNOK - 1).Value = "NOK" Then
                                    If Not countTangoNOK.Exists(key) Then countTangoNOK.Add key, 1
                                End If
                            End If
                            
                            ' --------------------------------------------------------------
                        End If
                        
                    End If
                    
                
                ElseIf repsh.name Like "RECEPTION_*" Then
                
                
                
                    If Trim(Me.TextBoxMag.Value) = "" Or _
                        rRef.Value Like "*" & Me.TextBoxMag.Value & "*" Then
                
                         
                        If rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sem - 1).Value = Me.ListBoxScope.Value Then
                        
                        
                            key = rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_article - 1).Value
                            If Not pnCount.Exists(key) Then pnCount.Add key, 1
                            
                            ' internal
                            If rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Interne - 1).Value = "internal" Then
                                If Not internalCount.Exists(key) Then internalCount.Add key, 1
                                
                                costInternal = costInternal + CDbl(rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sigapp - 1).Value)
                            Else
                                
                                ' all section which is not internal
                                ' most important scope
                                ' --------------------------------------------------------------
                                
                                
                                If Trim(rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_Tango - 1).Value) = "" Or _
                                   Trim(rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_OKNOK - 1).Value) = "NO TANGO" Or _
                                   Trim(rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_OKNOK - 1).Value) = "TP04 PRICE" Then
                                
                                    ' scope only for no tango price PNs
                                    If Not countNoTango.Exists(key) Then countNoTango.Add key, 1
                                    costNoTango = costNoTango + CDbl(rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sigapp - 1).Value)
                                    If CStr(rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_cible - 1).Value) <> "" Then
                                        targetNoTango = targetNoTango + CDbl(rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_cible - 1).Value)
                                    End If
                                Else
                                
                                    calcOKNOK = CDbl(rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Ecart - 1).Value)
                                
                                    
                                    If Not countTango.Exists(key) Then countTango.Add key, 1
                                    
                                    costTango = costTango + CDbl(rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sigapp - 1).Value)
                                    costTarget = costTarget + CDbl(rRef.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_cible - 1).Value)
                                    
                                
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
                
                End If
            End If
            
            
            Set rRef = rRef.offset(1, 0)
            
        Loop Until Trim(rRef.Value) = ""
        
        Me.TextBoxCPN.Value = ""
        Me.TextBoxCPN.Value = pnCount.count
        Me.TextBox_CountInternal.Value = ""
        Me.TextBox_CountInternal.Value = internalCount.count
        
        Me.TextBox_CostInternal.Value = ""
        Me.TextBox_CostInternal.Value = Round(costInternal, 2)
        
        Me.TextBox_CountNoTango.Value = countNoTango.count
        Me.TextBox_CostNoTango.Value = Round(costNoTango, 2)
        
        Me.TextBox_RateNoTango.Value = ""
        If rateCountWithoutTango > 0 Then Me.TextBox_RateNoTango.Value = Round(rateSumWithoutTango / CDbl(rateCountWithoutTango), 2)
        
        Me.TextBox_TargetNoTango.Value = ""
        If targetNoTango > 0 Then
            Me.TextBox_TargetNoTango.Value = Round(targetNoTango, 2)
            
            
            If CStr(Me.TextBox_RateNoTango.Value) = "" Then
                Me.TextBox_RateNoTango.Value = Round(1# * (costNoTango / targetNoTango), 2)
            End If
        End If
            
        
        
        
        
        Me.TextBox_CostTango.Value = Round(costTango, 2)
        Me.TextBox_CountTango = countTangoOK.count + countTangoNOK.count
        Me.TextBox_CountTangoNOK = countTangoNOK.count
        
        Me.TextBox_CostTarget.Value = Round(costTarget, 2)
        
        If costTarget > 0 Then Me.TextBox_RATE = Round(CDbl(costTango / costTarget), 2)
        
        ' final gap
        costGap = costTango - costTarget
        Me.TextBox_CostGap.Value = Round(costGap, 2)
        
End Sub



Private Sub TextBoxMag_Change()
    ListBoxScope_Click
End Sub
