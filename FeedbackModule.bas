Attribute VB_Name = "FeedbackModule"
Option Explicit


Public Sub makeSomeFeedback()
    innerMakeSomeFeedback
End Sub

Private Sub innerMakeSomeFeedback(Optional sh1 As Worksheet, Optional sh2 As Worksheet)
    
    
    ' sh1 as reception, sh2 as green light
    
    Dim fdbck As New FeedbackHandler
    
    
    ' before do anytning check of sh2 is really green light
    
    If sh2 Is Nothing Then
        Set sh2 = ThisWorkbook.ActiveSheet
    End If
    
    
    If fdbck.checkIfReallyGreenLight(sh2) Then
    
        Dim sh As Variant
        If sh1 Is Nothing Then
            
            With FeedbackForm
                With .ComboBox1
                    .Clear
                    
                    For Each sh In ThisWorkbook.Sheets
                    
                        If sh.name Like "RECEPTION*" Then
                            .addItem sh.name
                        End If
                    Next sh
                End With
                
                .show
                
                Set sh1 = Nothing
                On Error Resume Next
                Set sh1 = ThisWorkbook.Sheets(.shName)
            End With
            
        End If
        
        
        If fdbck.checkIfReallyReception(sh1) Then
            ' ---------------------------------------------------------------
            
            ' double check on reception and green light ready!
            ' we can go with the logic now
            ' ===============================================================
            
            
            fdbck.setupSheets sh1, sh2
            
            Dim glRef As Range
            Dim recRef As Range
            
            
            Set glRef = sh2.Cells(2, 1)
            Set recRef = sh1.Cells(2, 1)
            
            ' simple double loop on checking by each line in gl going through reception
            
            Dim progBar As New StatusHandler
            progBar.init_statusbar 100
            progBar.simplifyStatusBar
            progBar.show
            
            Dim fi As FeedbackItem
            Do
            
                ' every new iteration need to have going from the top in reception
                Set recRef = sh1.Cells(2, 1)
                Set fi = Nothing
                
                
                If Trim(glRef.Offset(0, EVO.E_GREEN_LIGHT_IS_INTERNAL - 1).Value) = "" Then
                
                
                    Set fi = New FeedbackItem
                    Set fi.rangeRefernceInGreen = glRef
                    
                    
                    Do
                    
                        If CStr(glRef.Offset(0, EVO.E_GREEN_LIGHT_Reference - 1).Value) = _
                            CStr(recRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_article - 1).Value) Then
                            
                            
                                fi.dictionaryOfMatchesFromReception.Add CStr(recRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_article - 1).row), _
                                    recRef.Offset(0, EVO.E_FINAL_TOUCH_RECEPTION_article - 1)
                            
                        End If
                        
                        Set recRef = recRef.Offset(1, 0)
                    Loop Until Trim(recRef.Value) = ""
                    
                    ' OK
                    ' Debug.Print "fi.dictionaryOfMatchesFromReception.Count: " & fi.dictionaryOfMatchesFromReception.Count
                    fdbck.tryToAddNewElement fi
                End If
                
                progBar.progress_increase
            
                Set glRef = glRef.Offset(1, 0)
            Loop Until Trim(glRef.Value) = ""
            
            
            ' after filling dictionaries
            ' go again and try to add some comments if possible
            Dim cmntTxt As String
            
            Set glRef = sh2.Cells(2, 1)
            Do
            
                Set fi = fdbck.getElementFromDictionary(glRef.row)
                If Not fi Is Nothing Then
                    
                    cmntTxt = ""
                    If Not glRef.Offset(0, EVO.E_GREEN_LIGHT_PRE_SERIAL_PRICE_YPRS_contract - 1).Comment Is Nothing Then
                        cmntTxt = glRef.Comment.Text
                        glRef.Offset(0, EVO.E_GREEN_LIGHT_PRE_SERIAL_PRICE_YPRS_contract - 1).Comment.Delete
                    Else
                        
                    End If
                    
                    With glRef.Offset(0, EVO.E_GREEN_LIGHT_PRE_SERIAL_PRICE_YPRS_contract - 1)
                        
                        
                        ' parseToText has side effects - for example calc trigger
                        .AddComment cmntTxt & Chr(10) & fi.parseToText()
                        
                        .Font.Bold = True
                        .Comment.Shape.TextFrame.AutoSize = True
                        
                        
                        If fi.trigger > 0.1 Then
                            .Interior.Color = RGB(40, 240, 60)
                        ElseIf fi.lastPriceFromReception = 0 Then
                            .Interior.Color = RGB(200, 200, 200)
                        Else
                            .Interior.Color = RGB(20, 20, 250)
                        End If
                    End With
                    
                    
                End If
                
                progBar.progress_increase
                Set glRef = glRef.Offset(1, 0)
            Loop Until Trim(glRef.Value) = ""
            
            
            
            
            progBar.hide
            Set progBar = Nothing
            
            
            ' ===============================================================
            
            ' ---------------------------------------------------------------
        Else
            MsgBox "Chosen sheet is in wrong std - please verify if there is really reception output!", vbCritical
        End If
    Else
        MsgBox "Chosen sheet is in wrong std - please verify if there is really green light output!", vbCritical
    End If
    
End Sub
