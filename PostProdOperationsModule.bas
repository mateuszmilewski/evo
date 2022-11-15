Attribute VB_Name = "PostProdOperationsModule"
Public Sub changeNameToSynthesis()
Attribute changeNameToSynthesis.VB_ProcData.VB_Invoke_Func = "R\n14"
    
    ActiveSheet.name = "synthesis"
End Sub


Public Sub copyShFromActiveEVOToFirstAvailablePPx1Report()
Attribute copyShFromActiveEVOToFirstAvailablePPx1Report.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' copySh1 Macro
'

'
    ' Sheets("GREEN_LIGHT_").Select
    
    
    Dim wrk1 As Workbook, sh1 As Worksheet
    
    Dim swrk As Workbook
    Set swrk = Nothing
    
    
    For Each wrk1 In Workbooks
        
        For Each sh1 In wrk1.Sheets
            
            If sh1.name = "synthesis" Then
                Set swrk = wrk1
                Exit For
            End If
        Next sh1
        
        If Not swrk Is Nothing Then
            Exit For
        End If
        
    Next wrk1
    
    
    ActiveSheet.Copy Before:=swrk.Sheets(1)
End Sub

