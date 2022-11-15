VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KpiHandler"
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

' KPI handler

Private m As Worksheet
Private pns As Dictionary
Private pnItem As PartNumberItemForKpi
Private stHndlr As StatusHandler

Public Sub set_master(mm As Worksheet)
    Set m = mm
End Sub

Public Sub setStatusHandler(sh1 As StatusHandler)
    Set stHndlr = sh1
End Sub


Public Sub makeRepFromPns()
    
    If pns.count > 0 Then
    
        stHndlr.init_statusbar (pns.count + 1)
        stHndlr.show
        stHndlr.progress_increase
        
        ' make some rep here
        Dim repsh As Worksheet, refRange As Range
        Set repsh = ThisWorkbook.Sheets.Add
        repsh.name = tryToRenameWorksheet(repsh)
        
        makeSomeLabelsInNewRepSh repsh
        
        
        Set refRange = repsh.Cells(2, 1)
        
        Dim key As Variant, pi As PartNumberItemForKpi
        For Each key In pns.Keys
            
            Set pi = pns(key)
            
            With refRange
                .offset(0, EVO.E_KPI_SRC_REF - 1).Value = pi.pn
                .offset(0, EVO.E_KPI_SRC_E_COFOR - 1).Value = pi.COFOR_EXPEDITEUR
                .offset(0, EVO.E_KPI_SRC_V_COFOR - 1).Value = pi.COFOR_VENDEUR
                
                .offset(0, EVO.E_KPI_SRC_SUPPLIER_NAME - 1).Value = pi.NOM_FOURNISSEUR
                
                
                .offset(0, EVO.E_KPI_SRC_DAP - 1).Value = Math.Abs(Int(pi.isDap))
                
                ' negative logic - bi kerfur!
                .offset(0, EVO.E_KPI_SRC_NO_DATA - 1).Value = Math.Abs(Int(pi.noDataAvail))
                
                ' ferrage or montage?
                .offset(0, EVO.E_KPI_SRC_TYPE_DE_PIECE - 1).Value = pi.TYPE_DE_PIECE
                
                .offset(0, EVO.E_KPI_SRC_BLUE - 1).Value = Math.Abs(Int(pi.isBlue))
                .offset(0, EVO.E_KPI_SRC_GREEN - 1).Value = Math.Abs(Int(pi.isGreen))
                .offset(0, EVO.E_KPI_SRC_YELLOW - 1).Value = Math.Abs(Int(pi.isYellow))
                
                .offset(0, EVO.E_KPI_SRC_PLE - 1).Value = Math.Abs(Int(pi.pleData))
                .offset(0, EVO.E_KPI_SRC_UA - 1).Value = Math.Abs(Int(pi.uaData))
                
                
            End With
            
            Set refRange = refRange.offset(1, 0)
            stHndlr.progress_increase
        Next
        
        stHndlr.hide
    End If
End Sub

Private Sub makeSomeLabelsInNewRepSh(s1 As Worksheet)
    
    With s1
        .Cells(1, EVO.E_KPI_SRC_REF).Value = "REF"
        .Cells(1, EVO.E_KPI_SRC_REF).AddComment "Part Number"
        
        .Cells(1, EVO.E_KPI_SRC_DAP).Value = "DAP?"
        
        .Cells(1, EVO.E_KPI_SRC_V_COFOR).Value = "V-COFOR"
        .Cells(1, EVO.E_KPI_SRC_E_COFOR).Value = "E-COFOR"
        .Cells(1, EVO.E_KPI_SRC_SUPPLIER_NAME).Value = "SUPPLIER"
        .Cells(1, EVO.E_KPI_SRC_TYPE_DE_PIECE).Value = "TYPE DE PIECE"
        .Cells(1, EVO.E_KPI_SRC_UA).Value = "UA?"
        .Cells(1, EVO.E_KPI_SRC_PLE).Value = "PLE?"
        .Cells(1, EVO.E_KPI_SRC_GREEN).Value = "GREEN?"
        .Cells(1, EVO.E_KPI_SRC_BLUE).Value = "BLUE?"
        .Cells(1, EVO.E_KPI_SRC_YELLOW).Value = "YELLOW?"
        
        .Cells(1, EVO.E_KPI_SRC_NO_DATA).Value = "NO DATA!"
        .Cells(1, EVO.E_KPI_SRC_NO_DATA).AddComment "No data available!"
        
    End With
End Sub

Private Function tryToRenameWorksheet(s1 As Worksheet) As String

    tryToRenameWorksheet = s1.name
    
    Dim tmpNewName As String, mm As String, dd As String
    mm = CStr(Month(Date))
    dd = CStr(Day(Date))
    
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
    
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    
    tmpNewName = "KPI_SRC_" & CStr(Year(Date)) & mm & dd & "_"
    
    
    On Error Resume Next
    s1.name = CStr(tmpNewName)
    
    noFukinDryBecauseFukOff s1, tmpNewName
    tryToRenameWorksheet = s1.name

End Function

Private Sub noFukinDryBecauseFukOff(ByRef psh As Worksheet, newName As String)

    On Error Resume Next
    psh.name = CStr(newName)
    
    If psh.name = newName Then
        Exit Sub
    Else
    
        If Len(newName) < 30 Then
            noFukinDryBecauseFukOff psh, newName & "I"
        Else
            Exit Sub
        End If
    End If
End Sub


Public Sub fillPartNumberDictionary()
    
    Dim r As Range
    Set r = m.Cells(3, 1)
    
    
    Do
    
        If Not pns.Exists(CStr(r.Value)) Then
            
            Set pnItem = New PartNumberItemForKpi
            pnItem.pn = CStr(r.Value)
            pnItem.noDataAvail = False
            pnItem.COFOR_VENDEUR = m.Cells(r.row, EVO.E_MASTER_2_COFOR_VENDEUR).Value
            pnItem.COFOR_EXPEDITEUR = m.Cells(r.row, EVO.E_MASTER_2_COFOR_EXPEDITEUR).Value
            pnItem.NOM_FOURNISSEUR = m.Cells(r.row, EVO.E_MASTER_2_NOM_FOURNISSEUR).Value
            pnItem.TYPE_DE_PIECE = m.Cells(r.row, EVO.E_MASTER_2_TYPE_DE_PIECE).Value
                        

            pnItem.isDap = checkIfDap(r)
            pnItem.uaData = checkIfUADataAvail(r, pnItem)
            pnItem.pleData = checkIfPleDataAvail(r, pnItem)
            
            
            ' color by column with cofor: COFOR VENDEUR
            pnItem.isBlue = checkIfBlue(m.Cells(r.row, EVO.E_MASTER_2_COFOR_VENDEUR))
            pnItem.isGreen = checkIfGreen(m.Cells(r.row, EVO.E_MASTER_2_COFOR_VENDEUR))
            ' a little diff logic, becuase yellow can be random
            ' so checking if yellow appears at least once!
            pnItem.isYellow = checkIfYellow(r)
            
            
            ' NO DATA AT ALL!
            If r.Font.Bold = True And r.Font.Color = RGB(255, 0, 0) Then
                pnItem.noDataAvail = True
            End If
            
            pns.Add CStr(r.Value), pnItem
            
            Set pnItem = Nothing

        End If
    
        Set r = r.offset(1, 0)
    Loop Until r = ""
    

    
End Sub


Private Function checkIfBlue(ByRef r17 As Range) As Boolean
    
    If r17.Interior.Color = RGB(200, 200, 255) Then
        checkIfBlue = True
    Else
        checkIfBlue = False
    End If
End Function

Private Function checkIfGreen(ByRef r17 As Range) As Boolean
    
    If r17.Interior.Color = RGB(200, 255, 200) Then
        checkIfGreen = True
    Else
        checkIfGreen = False
    End If
End Function

Private Function checkIfYellow(ByRef r As Range) As Boolean
    
    checkIfYellow = False
    
    ' 255, 255, 150
    
    Dim x As Integer
    For x = EVO.E_MASTER_2_COFOR_VENDEUR To EVO.E_MASTER_2_ML
        
        If m.Cells(r.row, x).Interior.Color = RGB(255, 255, 150) Then
            checkIfYellow = True
            Exit For
        End If
    Next x
    
End Function


Private Function checkIfPleDataAvail(ByRef r As Range, pi As PartNumberItemForKpi) As Boolean
    checkIfPleDataAvail = False
    
    
    ' checking 3 things:
    ' is gefco cofot avail
    ' is mail in proper format
    ' is dh dates are really dates
    
    
    Dim rGefcoCofor As Range
    Dim rMail As Range
    Dim rDHEF As Range, rDHAS As Range
    
    Set rGefcoCofor = m.Cells(r.row, EVO.E_MASTER_2_COFOR_GEFCO)
    Set rMail = m.Cells(r.row, EVO.E_MASTER_2_mail_contact_transport)
    Set rDHEF = m.Cells(r.row, EVO.E_MASTER_2_DHEF)
    Set rDHAS = m.Cells(r.row, EVO.E_MASTER_2_DHAS)
    
    
    If rGefcoCofor.Value Like "??????  ??" Then
        
        If rMail.Value Like "*@*" Then
            
            If IsDate(rDHEF.Value) And IsDate(rDHAS.Value) Then
            
                checkIfPleDataAvail = True
            End If
        End If
    End If
End Function


Private Function checkIfUADataAvail(ByRef r As Range, pi As PartNumberItemForKpi) As Boolean
    checkIfUADataAvail = False
    
    Dim rr As Range, rr2 As Range
    ' checking tylko po condi i qty - should be enough!
    Set rr = m.Cells(r.row, EVO.E_MASTER_2_CONDI)
    Set rr2 = m.Cells(r.row, EVO.E_MASTER_2_UC)
    
    If pi.isDap = False Then
    
        If CStr(rr.Value) <> "" And CStr(rr2.Value) <> "" Then
        
            ' 1st scenario - the simple one
            If IsNumeric(rr.Value) And IsNumeric(rr2.Value) Then
                checkIfUADataAvail = True
            
            ElseIf CStr(rr.Value) <> "" And IsNumeric(rr2.Value) Then
                checkIfUADataAvail = True
                
            ElseIf CStr(rr.Value) = "" Then
                checkIfUADataAvail = False
                
            ElseIf Not IsNumeric(rr2.Value) Then
                checkIfUADataAvail = False
                
            End If
            
        Else
            checkIfUADataAvail = False
            
        End If
    Else
        checkIfUADataAvail = False
    End If
End Function

Private Function checkIfDap(ByRef r As Range) As Boolean
    checkIfDap = False
    
    Dim refRange As Range
    Set refRange = m.Cells(r.row, EVO.E_MASTER_2_COFOR_VENDEUR)
    
    If refRange.Value = EVO.G_TXT_IN_CELL Then
        checkIfDap = True
    
    ElseIf refRange.Value = EVO.G_TXT_IN_CELL_II Then
        checkIfDap = True
        
    Else
        checkIfDap = False
        
    End If
End Function

Private Sub Class_Initialize()
    Set pns = New Dictionary
End Sub


Private Sub Class_Terminate()
    Set pns = Nothing
End Sub
