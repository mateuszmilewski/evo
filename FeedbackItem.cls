VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeedbackItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rangeRefernceInGreen As Range
Public dictionaryOfMatchesFromReception As Dictionary

Public trigger As Double
Public lastPriceFromReception As Double
Public numGL_price As Double




Public Function parseToText() As String


    Dim yyyycw As Long, yyyycwCurr As Long
    yyyycw = 0
    yyyycwCurr = 0

    Dim parseSummary As String
    Dim getPriceFromGreenLight As Range
    Set getPriceFromGreenLight = rangeRefernceInGreen.offset(0, EVO.E_GREEN_LIGHT_PRE_SERIAL_PRICE_YPRS_contract - 1)
    
    parseSummary = parseSummary & " receptions: " & dictionaryOfMatchesFromReception.count
    
    
    Dim key As Variant, r As Range, count1 As Integer
    count1 = 0
    

    Dim numRE_price As Double
    
    On Error Resume Next
    numGL_price = CDbl(getPriceFromGreenLight.Value)



    
    Dim tmpStr As String
    tmpStr = ""
    
    
    For Each key In dictionaryOfMatchesFromReception.Keys
    
        Set r = Nothing
        On Error Resume Next
        Set r = dictionaryOfMatchesFromReception(key)
        If Not r Is Nothing Then
        
            ' calcul!
            ' ---------------------------------------------------------------------------------------------------------------------
            On Error Resume Next
            numRE_price = CDbl(r.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_Sigapp - EVO.E_FINAL_TOUCH_RECEPTION_article).Value)
            
            
            If yyyycw = 0 And yyyycwCurr = 0 Then
                ' sem : std: 2020 CW19 YYYY CWXX
                yyyycw = CLng(Replace(r.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sem - EVO.E_FINAL_TOUCH_RECEPTION_article).Value, " CW", ""))
                yyyycwCurr = CLng(Replace(r.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sem - EVO.E_FINAL_TOUCH_RECEPTION_article).Value, " CW", ""))
                lastPriceFromReception = numRE_price
            Else
                
                yyyycwCurr = CLng(Replace(r.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sem - EVO.E_FINAL_TOUCH_RECEPTION_article).Value, " CW", ""))
                
                If yyyycwCurr > yyyycw Then
                    lastPriceFromReception = numRE_price
                    yyyycw = yyyycwCurr
                End If
            End If
            
            
            If numGL_price > 0 And numRE_price > 0 Then
            
                If CDbl(Math.Abs(numRE_price - numGL_price) / numGL_price) <= 0.1 Then
                    count1 = count1 + 1
                End If
            End If
            ' ---------------------------------------------------------------------------------------------------------------------
        
        
            tmpStr = tmpStr & Chr(10) & _
                "reception src-row: " & key & _
                ", cw: " & CStr(r.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_Sem - EVO.E_FINAL_TOUCH_RECEPTION_article)) & _
                ", prix sigapp: " & CStr(r.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_Sigapp - EVO.E_FINAL_TOUCH_RECEPTION_article))
        End If
    Next
    
    
    
    
    If dictionaryOfMatchesFromReception.count > 0 Then
        parseSummary = parseSummary & " with NOK PPx1: " & CStr(count1) & Chr(10)
        parseSummary = parseSummary & "last price from reception: " & CStr(lastPriceFromReception)
        
        ' !!!
        trigger = CDbl((numGL_price - lastPriceFromReception) / numGL_price)
    End If
    
    
    
    parseToText = CStr(parseSummary) & Chr(10) & Chr(10) & CStr(tmpStr)
    
End Function


Private Sub Class_Initialize()
    Set dictionaryOfMatchesFromReception = New Dictionary
    trigger = 0
End Sub

Private Sub Class_Terminate()
    Set dictionaryOfMatchesFromReception = Nothing
End Sub









' obsolete - logic now also in parseToText
Public Function parseSummary() As String
    parseSummary = ""
    
    Dim getPriceFromGreenLight As Range
    Set getPriceFromGreenLight = rangeRefernceInGreen.offset(0, EVO.E_GREEN_LIGHT_PRE_SERIAL_PRICE_YPRS_contract - 1)
    
    parseSummary = parseSummary & " receptions: " & dictionaryOfMatchesFromReception.count
    
    Dim key As Variant, r As Range, count1 As Integer
    count1 = 0
    
    Dim numGL_price As Double
    Dim numRE_price As Double
    
    On Error Resume Next
    numGL_price = CDbl(getPriceFromGreenLight.Value)
    
    
    For Each key In dictionaryOfMatchesFromReception.Keys
        Set r = Nothing
        On Error Resume Next
        Set r = dictionaryOfMatchesFromReception(key)
        If Not r Is Nothing Then
        
            On Error Resume Next
            numRE_price = CDbl(r.offset(0, EVO.E_FINAL_TOUCH_RECEPTION_prix_Sigapp - EVO.E_FINAL_TOUCH_RECEPTION_article).Value)
            
            
            If numGL_price > 0 And numRE_price > 0 Then
            
                If CDbl(Math.Abs(numRE_price - numGL_price) / numGL_price) <= 0.1 Then
                    count1 = count1 + 1
                End If
            End If
        End If
    Next
    
    'parseSummary = parseSummary & " with PPx1: " & CStr(count1) & _
    '    ", coef:" & CStr(CDbl((dictionaryOfMatchesFromReception.Count * 1#) / (count1 * 1#)))
    parseSummary = parseSummary & " with PPx1: " & CStr(count1)
End Function


