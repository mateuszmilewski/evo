VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotHandler2"
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

' PIVOT HANDLER 2


Private pc As PivotCache
Private pT As PivotTable


Private srcSheet As Worksheet
Private refRangeInSrcSheet As Range
Private pivotSheet As Worksheet
Private srcSheetName As String
Private pivotSheetName As String
Private tcamSheetName As String

' really special one!
Public proxy2 As Worksheet
Private rngFromProxy2 As Range

' tcam final
Public tcamSheet As Worksheet
Private tcamRng As Range




Public Sub runMainLogicForCreationTheoreticPivotTable()


    If Not rngFromProxy2 Is Nothing Then
    
        Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, rngFromProxy2)
        Set pT = pc.CreatePivotTable(pivotSheet.Cells(1, 1), "THEORETIC_PIVOT2_" & CStr(pivotSheet.name))
        
        ' test loop through fields
        Dim pfs As PivotFields, pf As PivotField, pi As PivotItem
        Set pfs = pT.PivotFields
        
        

        For Each pf In pfs
            Debug.Print pf.name
            ' Debug.Print pf.Orientation
            
            If CStr(pf.name) = "S_ORDER_DATE" Then
                pf.Orientation = xlRowField
                pf.Position = 1
                ' pf.AutoGroup
                
                pf.name = "ORDER"
                
            End If
            
            If CStr(pf.name) = "S_DELIVERY_DATE" Then
                pf.Orientation = xlColumnField
                pf.Position = 1
                ' pf.AutoGroup
                
                
                pf.name = "DELIVERY"
                
                'For Each pi In pf.PivotItems
                '
                '    If IsDate(pi.Name) Then
                '        ' Debug.Print pi.Name
                '        pi.Name = Format(pi.Name, "dd/mm/yyyy")
                '    End If
                'Next pi
            End If
        Next pf
        
        
        For Each pf In pfs
            
            If CStr(pf.name) = "ROUTE NAME AND PILOT" Then
                pf.Orientation = xlRowField
                pf.Position = 2
            End If
            
            
            
            If CStr(pf.name) = "TYPE_DE_PIECE" Then
                pf.Orientation = xlPageField
                pf.Position = 1
            End If
            
            
            If CStr(pf.name) = "DELIVERY YEAR" Then
                pf.Orientation = xlPageField
                pf.Position = 1
                On Error Resume Next
                pf.CurrentPage = Year(Date)
            End If
            
            If CStr(pf.name) = "DELIVERY WEEK" Then
                pf.Orientation = xlPageField
                pf.Position = 1
                On Error Resume Next
                pf.CurrentPage = CStr(Application.WorksheetFunction.IsoWeekNum(CDate(Date)))
                
            End If
            
            
            
            If CStr(pf.name) = "OQ 2" Then
                pf.Orientation = xlDataField
                pf.Position = 1
                pf.Function = xlSum
                pf.name = "OQ ( OQ 2 )"
            End If
            

        Next pf
        
        ' ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
        ' pT.RowAxisLayout xlTabularRow
        'pT.RepeatAllLabels xlRepeatLabels
        
        
        For Each pf In pfs
            If CStr(pf.name) = "COFOR_EXPEDITEUR" Then
                pf.Orientation = xlRowField
                pf.Position = 3
            End If
            
        Next pf
        
        For Each pf In pfs
            If CStr(pf.name) = "NOM_FOURNISSEUR" Then
                pf.Orientation = xlRowField
                pf.Position = 4
            End If
        Next pf
        
        
        For Each pf In pfs
            
            If CStr(pf.name) = "CONDI" Then
                pf.Orientation = xlRowField
                pf.Position = 5
            End If
        Next pf
        
        
        
        pT.RowAxisLayout xlTabularRow
        pT.RepeatAllLabels xlRepeatLabels
        pivotSheet.columns("A:A").ColumnWidth = 17
        pivotSheet.columns("B:B").ColumnWidth = 37
        pivotSheet.columns("C:C").ColumnWidth = 17
        pivotSheet.columns("D:D").ColumnWidth = 14
        pivotSheet.columns("E:E").ColumnWidth = 12
        
        
        For Each pf In pfs
            pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        Next pf
        
        
        
        
    Else
    
        MsgBox "no data provided for runMainLogicForCreationTheoritcPivotTable", vbCritical
    End If
    
End Sub



Public Sub runMainLogicForCreationPivotTable()


    If Not rngFromProxy2 Is Nothing Then
    
        Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, rngFromProxy2)
        Set pT = pc.CreatePivotTable(pivotSheet.Cells(1, 1), "PIVOT2_" & CStr(pivotSheet.name))
        
        
        ' test loop through fields
        Dim pfs As PivotFields, pf As PivotField, pi As PivotItem
        Set pfs = pT.PivotFields
        
        

        For Each pf In pfs
            Debug.Print pf.name
            ' Debug.Print pf.Orientation
            
            If CStr(pf.name) = "S_ORDER_DATE" Then
                pf.Orientation = xlRowField
                pf.Position = 1
                ' pf.AutoGroup
                
                pf.name = "ORDER"
                
            End If
            
            If CStr(pf.name) = "S_DELIVERY_DATE" Then
                pf.Orientation = xlColumnField
                pf.Position = 1
                ' pf.AutoGroup
                
                
                pf.name = "DELIVERY"
                
                'For Each pi In pf.PivotItems
                '
                '    If IsDate(pi.Name) Then
                '        ' Debug.Print pi.Name
                '        pi.Name = Format(pi.Name, "dd/mm/yyyy")
                '    End If
                'Next pi
            End If
        Next pf
        
        
        For Each pf In pfs
            
            If CStr(pf.name) = "ROUTE NAME AND PILOT" Then
                pf.Orientation = xlRowField
                pf.Position = 2
            End If
            
            
            
            If CStr(pf.name) = "TYPE_DE_PIECE" Then
                pf.Orientation = xlPageField
                pf.Position = 1
            End If
            
            
            If CStr(pf.name) = "DELIVERY YEAR" Then
                pf.Orientation = xlPageField
                pf.Position = 1
                On Error Resume Next
                pf.CurrentPage = Year(Date)
            End If
            
            If CStr(pf.name) = "DELIVERY WEEK" Then
                pf.Orientation = xlPageField
                pf.Position = 1
                On Error Resume Next
                pf.CurrentPage = CStr(Application.WorksheetFunction.IsoWeekNum(CDate(Date)))
                
            End If
            
            
            
            If CStr(pf.name) = "COQ 2" Then
                pf.Orientation = xlDataField
                pf.Position = 1
                pf.Function = xlSum
                pf.name = "Confirmed OQ ( COQ 2 )"
            End If
            

        Next pf
        
        ' ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
        ' pT.RowAxisLayout xlTabularRow
        'pT.RepeatAllLabels xlRepeatLabels
        
        
        For Each pf In pfs
            If CStr(pf.name) = "COFOR_EXPEDITEUR" Then
                pf.Orientation = xlRowField
                pf.Position = 3
            End If
            
        Next pf
        
        For Each pf In pfs
            If CStr(pf.name) = "NOM_FOURNISSEUR" Then
                pf.Orientation = xlRowField
                pf.Position = 4
            End If
        Next pf
        
        
        For Each pf In pfs
            
            If CStr(pf.name) = "CONDI" Then
                pf.Orientation = xlRowField
                pf.Position = 5
            End If
        Next pf
        
        
        
        pT.RowAxisLayout xlTabularRow
        pT.RepeatAllLabels xlRepeatLabels
        pivotSheet.columns("A:A").ColumnWidth = 17
        pivotSheet.columns("B:B").ColumnWidth = 37
        pivotSheet.columns("C:C").ColumnWidth = 17
        pivotSheet.columns("D:D").ColumnWidth = 14
        pivotSheet.columns("E:E").ColumnWidth = 12
        
        For Each pf In pfs
            pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        Next pf
        
        
        
    Else
    
        MsgBox "no data provided for runMainLogicForCreationPivotTable", vbCritical
    End If
    
End Sub







Public Sub initPivotSource()
    Set srcSheet = ThisWorkbook.Sheets.Add
    srcSheet.name = CStr(EVO.TryToRenameModule.tryToRenameWorksheet(srcSheet, "SrcPivot_"))
    srcSheetName = CStr(srcSheet.name)
End Sub

Public Sub initPivotSheet()
    Set pivotSheet = ThisWorkbook.Sheets.Add
    pivotSheet.name = CStr(EVO.TryToRenameModule.tryToRenameWorksheet(pivotSheet, "Pivot_"))
    pivotSheetName = CStr(pivotSheet.name)
End Sub

Public Sub initTcamRepSheet()
    
    Set tcamSheet = ThisWorkbook.Sheets.Add
    tcamSheet.name = CStr(EVO.TryToRenameModule.tryToRenameWorksheet(tcamSheet, "TCAM_"))
    tcamSheetName = CStr(tcamSheet.name)
    
End Sub





Public Sub setRangeDataFromProxy2ForPivot(mr As Range)
    Set rngFromProxy2 = mr
End Sub

Public Function getRangeDataFromProxy2() As Range
    
    Set getRangeDataFromProxy2 = rngFromProxy2
End Function


Public Function getRefRangeInSrcSheet() As Range
    Set getRefRangeInSrcSheet = refRangeInSrcSheet
End Function

Public Sub moveRefRangeInSrcSheet()
    Set refRangeInSrcSheet = refRangeInSrcSheet.Offset(1, 0)
End Sub



Public Function getPivotSource() As Worksheet
    Set getPivotSource = srcSheet
End Function

Public Function getPivot() As Worksheet
    Set getPivot = pivotSheet
End Function

Public Sub setTcamPivot(mp As PivotTable)
    Set pT = Nothing
    On Error Resume Next
    Set pT = mp
End Sub

Public Sub setPivotSheet(psh1 As Worksheet)
    Set pivotSheet = Nothing
    On Error Resume Next
    Set pivotSheet = psh1
End Sub



Private Sub Class_Initialize()
    Set srcSheet = Nothing
    Set pivotSheet = Nothing
    Set rngFromProxy2 = Nothing
End Sub

Private Sub Class_Terminate()
    Set srcSheet = Nothing
    Set pivotSheet = Nothing
    Set rngFromProxy2 = Nothing
End Sub

