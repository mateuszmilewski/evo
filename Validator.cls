VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pusSh As Worksheet
Private cpl As Worksheet



Public Function getWorksheetForValidation(pusWrkName As String, fileType As E_FILE_TYPE) As Boolean



    Dim shName As String
    Dim rowLabel As Integer
    Dim rRef As Range
    
    If fileType = E_MASTER_PUS Then
        Set refR = ThisWorkbook.Sheets("forValidation").Range("D6")
        
        shName = CStr(ThisWorkbook.Sheets("forValidation").Range("A6").Value)
        rowLabel = Int(ThisWorkbook.Sheets("forValidation").Range("C6").Value)
    ElseIf fileType = E_FEED_CPL Then
        Set refR = ThisWorkbook.Sheets("forValidation").Range("D2")
        
        shName = CStr(ThisWorkbook.Sheets("forValidation").Range("A2").Value)
        rowLabel = Int(ThisWorkbook.Sheets("forValidation").Range("C2").Value)
    Else
        shName = ""
        rowLabel = -1
        Set refR = Nothing
    End If
    

    getWorksheetForValidation = True

    Dim wrk As Workbook
    Set wrk = Workbooks(CStr(pusWrkName))
    Dim mainSh As Worksheet
    Set main = Nothing
    On Error Resume Next
    Set main = wrk.Sheets(shName)
    
    If main Is Nothing Then
        getWorksheetForValidation = False
        
    Else
    
        ' going further
        Dim vr As Range
        Dim srcr As Range
        
        
        If Not refR Is Nothing Then
        
            Set srcr = main.Cells(rowLabel, 1)
            Set vr = refR
            Do
            
                ' 100% must be
                If CStr(vr.Value) = CStr(srcr.Value) Then
                    ' nop required
                Else
                    getWorksheetForValidation = False
                    Exit Do
                End If
                
                Set vr = vr.Offset(0, 1)
                Set srcr = srcr.Offset(0, 1)
            Loop Until Trim(vr) = ""
        
        Else
            getWorksheetForValidation = False
        End If
        
    End If
    
End Function




Private Sub Class_Initialize()
    Set pusSh = Nothing
    Set cpl = Nothing
End Sub

Private Sub Class_Terminate()
    Set pusSh = Nothing
    Set cpl = Nothing
End Sub
