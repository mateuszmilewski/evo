VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InterrocomHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'The MIT License (MIT)
'
'Copyright (c) 2020 FORREST
' Mateusz Milewski mateusz.milewski@mpsa.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
'
' THE EVO TOOL


Private interrocomDictionary As Dictionary
Private interrocomCollection As Collection
Private sh1 As Worksheet
Private lean As Worksheet


Public Sub setSh1(sh As Worksheet)
    Set sh1 = sh
End Sub

Public Sub setLeanSheet(ln As Worksheet)
    Set lean = ln
End Sub

Public Function getLean() As Worksheet
    Set getLean = lean
End Function


Public Sub putCollectionIntoEvo(Optional auto1 As Boolean, Optional projectAlias As String)


    Application.Calculation = xlCalculationManual
    
    If interrocomCollection.count > 0 Then
        
        Dim outEvo As Worksheet
        Set outEvo = ThisWorkbook.Sheets.Add
        
        Set lean = outEvo
        
        Dim strProjectName As String
        strProjectName = ""
        
        If auto1 Then
            If projectAlias <> "" Then
                strProjectName = projectAlias
            End If
        Else
            On Error Resume Next
            strProjectName = InputBox("please put for which project it is:", "please put for which project it is:")
        End If
        
        outEvo.name = tryToRenameWorksheet(outEvo, "INTERROCOM_" & strProjectName & "_")
        
        Dim row As Long
        row = 1
        
        With outEvo
            If row = 1 Then
                .Cells(row, EVO.E_OUT_INTERROCOM_A).Value = "ARTICLE"
                .Cells(row, EVO.E_OUT_INTERROCOM_COFOR).Value = "COFOR"
                .Cells(row, EVO.E_OUT_INTERROCOM_DIV).Value = "DIV"
                .Cells(row, EVO.E_OUT_INTERROCOM_FINAL_PRIX).Value = "T_PRICE"
                .Cells(row, EVO.E_OUT_INTERROCOM_REP).Value = "REPARTITION"
                .Cells(row, EVO.E_OUT_INTERROCOM_UN).Value = "UN"
                
                row = row + 1
            End If
            
            Dim i As InterrocomItem
            
            For Each i In interrocomCollection
            
            
                .Cells(row, EVO.E_OUT_INTERROCOM_A).Value = i.article
                .Cells(row, EVO.E_OUT_INTERROCOM_COFOR).Value = i.cofor
                .Cells(row, EVO.E_OUT_INTERROCOM_DIV).Value = i.div
                .Cells(row, EVO.E_OUT_INTERROCOM_FINAL_PRIX).Value = i.finalPrix
                .Cells(row, EVO.E_OUT_INTERROCOM_REP).Value = i.repartition
                .Cells(row, EVO.E_OUT_INTERROCOM_UN).Value = i.un
            
                row = row + 1
            Next i
        End With
    End If
    
    Application.Calculation = xlCalculationAutomatic
End Sub


Public Sub makeLoop()
    
    Dim rng As Range
    
    
    Dim row As Long
    row = 2
    
    
    Dim i As InterrocomItem
    
    Do
        Set rng = sh1.Cells(row, 1)
        
        ' ------------------------------------------------------
        Set i = New InterrocomItem
        
        With i
            .article = CStr(sh1.Cells(row, EVO.INTERROCOM_article).Value)
            .cofor = CStr(sh1.Cells(row, EVO.INTERROCOM_cofor).Value)
            .datePrix = parseToDate(sh1.Cells(row, EVO.INTERROCOM_Date_prix))
            .designation = CStr(sh1.Cells(row, EVO.INTERROCOM_Designation_achat).Value)
            .div = CStr(sh1.Cells(row, EVO.INTERROCOM_Div).Value)
            
            .orgCurrency = CStr(sh1.Cells(row, EVO.INTERROCOM_Devise).Value)
            
            ' Debug.Assert .orgCurrency = "EUR"
            
            On Error Resume Next
            .orgPrix = CDbl(prixToDouble(CStr(sh1.Cells(row, EVO.INTERROCOM_prix).Value)))
            
            On Error Resume Next
            .currencyRate = findCurrencyRate(.orgCurrency)
            
            .repartition = CStr(sh1.Cells(row, EVO.INTERROCOM_Repartition).Value)
            
            If IsNumeric(.orgPrix) And IsNumeric(.currencyRate) Then
                .orgPrixToEur = prixToEur(.orgPrix, .currencyRate)
            End If
            
            
            .un = CStr(sh1.Cells(row, EVO.INTERROCOM_Un).Value)
            .qtyUN = CLng(parseUN(.un))
            
            
            .finalPrix = .orgPrixToEur / .qtyUN
            
            
        End With
        
        
        interrocomCollection.Add i
        
        ' ------------------------------------------------------
        
        row = row + 1
        Set rng = rng.offset(1, 0)
    Loop Until Trim(rng.Value) = ""
    
    
End Sub

Private Function parseUN(unStr As String) As Long
    
    parseUN = 1
    
    
    Dim rng As Range, ir As Range
    Set rng = ThisWorkbook.Sheets("register").Range("J100:J200")
    
    For Each ir In rng
    
        If Trim(ir.Value) = "" Then
            Exit For
        End If
        
        If Trim(ir.Value) = unStr Then
            parseUN = CLng(ir.offset(0, 1).Value)
            Exit Function
        End If
    Next ir
    
    
End Function


Private Function prixToEur(prix As Double, rate As Double) As Double

    If rate > 0 Then
        prixToEur = prix / rate
    Else
        prixToEur = prix
    End If

End Function


Private Function prixToDouble(str1 As String) As Double
    prixToDouble = CDbl(Replace(str1, ".", ","))
End Function


Private Function findCurrencyRate(currStr) As Double


    findCurrencyRate = 1

    Dim formulaString As String
    Dim tmp As Range
    Set tmp = ThisWorkbook.Sheets("register").Range("D98")
    tmp.Value = ""
    
    formulaString = ThisWorkbook.Sheets("register").Range("D99").Formula
    
    
    
    formulaString = Replace(formulaString, "X", """X""")
    formulaString = Replace(formulaString, "X", currStr)
    
    tmp.Formula = formulaString
    
    findCurrencyRate = tmp.Value
    tmp.Value = ""
    
    ' Debug.Print findCurrencyRate
End Function



Public Function parseToDate(rng As Range) As Date
    
    If IsDate(rng.Value) Then
        parseToDate = CDate(rng.Value)
    Else
        parseToDate = CDate(1)
    End If
End Function


' obsolete
Public Sub addItem(i As InterrocomItem)
    Dim c As Collection
    
    If Not interrocomDictionary.Exists(CStr(i.rngRef.row)) Then
        interrocomDictionary.Add CStr(i.rngRef.row), i
    Else
        Debug.Print "duplicate on row! " & i.rngRef.row & " " & i.rngRef.Value
    End If
End Sub




Private Sub Class_Initialize()
    Set interrocomDictionary = New Dictionary
    Set interrocomCollection = New Collection
    
    Set lean = Nothing
End Sub




Private Sub Class_Terminate()
    Set interrocomDictionary = Nothing
    Set interrocomCollection = Nothing
    
    Set lean = Nothing
End Sub
