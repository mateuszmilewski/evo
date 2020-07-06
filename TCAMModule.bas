Attribute VB_Name = "TCAMModule"
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



Public Sub showDetails()

End Sub


Public Sub createSourceForPivot(ictrl As IRibbonControl)

    
    Dim w As Workbook
    
    With FileChooser
        .scenarioType = E_FORM_SCENARIO_CREATE_PIVOT_SCENARIO
        .BtnCopy.Enabled = False
        .BtnValid.Enabled = True
        
        .ComboBoxFeed.Clear
        .ComboBoxMaster.Clear
        
        
        For Each w In Application.Workbooks
            .ComboBoxFeed.AddItem w.name
            .ComboBoxMaster.AddItem w.name
        Next w
        
        .ComboBoxFeed.Value = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("m2").Value
        .ComboBoxMaster.Value = ThisWorkbook.Sheets(EVO.REG_SH_NM).Range("m1").Value
        
        .show
    End With
    
    MsgBox "GOTOWE!", vbInformation
    
End Sub

Public Sub innerCreateSourceForPivot(masterFileName, feedFileName, Optional sh As StatusHandler)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    
    If IsMissing(sh) Or (sh Is Nothing) Then
        Set sh = New StatusHandler
    End If
    
    ' master and feed worksheets
    Dim m As Worksheet, f As Worksheet
    ' starting from most impotant sheets!
    On Error Resume Next
    Set m = Workbooks(masterFileName).Sheets(MAIN_SH_BASE)
    On Error Resume Next
    Set f = Workbooks(feedFileName).Sheets(G_FEED_SH_MAIN)

    Dim ch As CopyHandler, ph As PivotHandler
    Set ch = New CopyHandler
    Set ph = New PivotHandler
    
    ch.init m, f, E_COPY_HANDLER_FOR_PIVOT_CREATION
    
    ' MsgBox "implementation under way!", vbInformation
    ch.copyForSourcePivot ph, sh
    
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Public Sub makePivot(ictrl As IRibbonControl)
    innerMakePivot
    MsgBox "READY!"
End Sub

Public Sub makeTPivot(ictrl As IRibbonControl)
    innerMakeTPivot
    MsgBox "READY!"
End Sub

Public Sub makeTCAM(ictrl As IRibbonControl)
    innerMakeTcamByCopingFromPivotOnActiveSheet
    MsgBox "READY!", vbInformation
End Sub

Public Sub goToProxy2(ictrl As IRibbonControl)
    
    innerGotoProxy2
    
End Sub




Private Sub innerMakePivot()


    If checkActiveSheetIfItIsProxy2() Then
    
        Dim ph As PivotHandler, oh As OperationsHandler
        Set ph = New PivotHandler
        Set oh = New OperationsHandler
        
        Set ph.proxy2 = ActiveSheet
        ph.initPivotSheet
        oh.makePivot ph
        
    Else
        MsgBox "Proxy2 sheet need to be active to perform action!", vbInformation
    End If
    
End Sub

Private Sub innerMakeTPivot()


    If checkActiveSheetIfItIsProxy2() Then
    
        Dim ph As PivotHandler, oh As OperationsHandler
        Set ph = New PivotHandler
        Set oh = New OperationsHandler
        
        Set ph.proxy2 = ActiveSheet
        ph.initPivotSheet
        oh.makeTPivot ph
        
    Else
        MsgBox "Proxy2 sheet need to be active to perform action!", vbInformation
    End If
    
End Sub

Private Function checkActiveSheetIfItIsProxy2() As Boolean
    checkActiveSheetIfItIsProxy2 = False
    
    If ActiveSheet.name Like "Proxy2_*" Then
        If ActiveSheet.Cells(1, 1).Value = "ID" Then
        
            If ActiveSheet.Cells(1, 2).Value = "WIERSZ" Then
            
                If ActiveSheet.Cells(1, 3).Value = "REF" Then
                
                
                    checkActiveSheetIfItIsProxy2 = True
                    
                    Dim ans As Variant
                    ans = MsgBox("Do you want to create PIVOT for: " & CStr(ActiveSheet.name) & " ? ", vbYesNo + vbQuestion)
                        
                    If ans = vbYes Then
                        checkActiveSheetIfItIsProxy2 = True
                    Else
                        checkActiveSheetIfItIsProxy2 = False
                        MsgBox "Nothing to do!"
                    End If
                End If
            End If
        End If
    End If

End Function



Public Sub innerMakeTcamByCopingFromPivotOnActiveSheet()



    Dim ph As New PivotHandler
    Dim oh As New OperationsHandler


    Dim thisPivot As PivotTable, srcAdrString As String
    Set thisPivot = checkIfPivotIsAvailInSheetWhichIsActive()
    
    If Not thisPivot Is Nothing Then
    
        ' logic
        Debug.Print thisPivot.name ' OK
        
        ph.setPivotSheet ThisWorkbook.Sheets(ActiveSheet.name)
        ph.setTcamPivot thisPivot
        
        With thisPivot
            ' Debug.Print .GrandTotalName ' OK but for nothing
            ' Debug.Print .GetData("ROUTE NAME AND PILOT") ' NOK no[]?
            
            Debug.Print .SourceData ' this is really important!
            ' example: Proxy2_20200518_!R1C1:R865C43
            srcAdrString = CStr(.SourceData)
            
            ph.initTcamRepSheet
            Dim prxy2ShNm As String
            prxy2ShNm = Split(srcAdrString, "!")(0)
            
            Set ph.proxy2 = Nothing
            On Error Resume Next
            Set ph.proxy2 = ThisWorkbook.Sheets(prxy2ShNm)
            oh.copyThisPivotToTcamReport ph
        End With
    End If
End Sub

Private Function checkIfPivotIsAvailInSheetWhichIsActive() As PivotTable
    Set checkIfPivotIsAvailInSheetWhichIsActive = Nothing
    
    
    If ActiveSheet.name Like "Pivot_*" Then
        ' almost
        ' checkIfPivotSheetIsActive = True
        
        Dim tmpSh As Worksheet
        Set tmpSh = ActiveSheet
        
        Dim pTmp As PivotTables
        On Error Resume Next
        Set pTmp = tmpSh.PivotTables
        
        Dim p1 As PivotTable
        Set p1 = pTmp.item(1)
        
        ' Debug.Print p1.Name
        Set checkIfPivotIsAvailInSheetWhichIsActive = p1
    Else
        MsgBox "Pivot worksheet need to be active!", vbCritical
    End If
End Function








Private Sub innerGotoProxy2()
    Debug.Print "innerGotoProxy2!"
    
    
    Dim filterForOrderDate As Date
    Dim filterForRouteName As String
    Dim filterForDeliveryDate As Date
    Dim tcamSheet As Worksheet
    Dim proxy2Sheet As Worksheet
    Dim startingPoint As Range
    
    If ActiveSheet.name Like "TCAM_*" Then
        If ActiveSheet.Cells(1, 1).Value = "TCAM REPORT" Then
        
            ' we are ok to go with searching data
        
            Set tcamSheet = ActiveSheet
            With tcamSheet
                Set startingPoint = .Range("h1")
            
            
                If IsDate(.Cells(startingPoint.Value, ActiveCell.Column).Value) Then
                    filterForDeliveryDate = CDate(.Cells(startingPoint.Value, ActiveCell.Column).Value)
                End If
                filterForRouteName = .Cells(ActiveCell.row, 1).Value
                filterForOrderDate = offsetUpForFirstOrderDate(startingPoint, .Cells(ActiveCell.row, 1))
            End With
            
            
            Debug.Print "filters: " & Chr(10) & _
                " filterForOrderDate: " & filterForOrderDate & Chr(10) & _
                " filterForRouteName: " & filterForRouteName & Chr(10) & _
                " filterForDeliveryDate: " & filterForDeliveryDate & Chr(10)
                
            ' all important filters are now defined - now go to proper proxy2 and put those filters!
            
            ThisWorkbook.Sheets(tcamSheet.Range("E1").Value).Activate
            Set proxy2Sheet = ThisWorkbook.Sheets(tcamSheet.Range("E1").Value)
            On Error Resume Next
            proxy2Sheet.ShowAllData
            
            If CLng(filterForOrderDate) <> 0 Then
                If CLng(filterForDeliveryDate) <> 0 Then
                    With proxy2Sheet.UsedRange
                        .AutoFilter Field:=E_PIVOT_PROXY2_ORDER_DATE, Operator:= _
                            xlFilterValues, Criteria2:=Array(2, filterForOrderDate)
                        .AutoFilter Field:=E_PIVOT_PROXY2_ROUTE_AND_PILOT, Criteria1:=filterForRouteName
                        .AutoFilter Field:=E_PIVOT_PROXY2_DELIVERY_DATE, Operator:= _
                            xlFilterValues, Criteria2:=Array(2, filterForDeliveryDate)
                    End With
                End If
            End If
        End If
    End If
End Sub


Private Function offsetUpForFirstOrderDate(sp As Range, cl As Range) As Date

    
    Do
        If IsDate(cl.Value) Then
            offsetUpForFirstOrderDate = CDate(cl.Value)
            Exit Do
        End If
        Set cl = cl.Offset(-1, 0)
    Loop Until IsDate(cl.Offset(1, 0).Value) Or Int(cl.row) = Int(sp.Value)
End Function




Public Function checkIfTcamSheet(sh1 As Worksheet) As Boolean


    checkIfTcamSheet = False
    If sh1.name Like "TCAM*" Then
        If sh1.Range("A1").Value = "TCAM REPORT" Then
            checkIfTcamSheet = True
        End If
    End If
End Function


Public Function calculateCostForCloe(c1 As Range) As Double

    calculateCostForCloe = 0
    
    Dim sh1 As Worksheet
    Set sh1 = c1.Parent
    
    Dim rts As Range
    Set rts = sh1.Range("A" & c1.row)
    
    
    Dim prxy2 As Worksheet, wynikSzukania As Range
    Set prxy2 = ThisWorkbook.Sheets(sh1.Range("E1").Value)
    
    Set wynikSzukania = prxy2.UsedRange.Find(rts.Value)
    
    If wynikSzukania Is Nothing Then
        calculateCostForCloe = 0
    Else
        calculateCostForCloe = CDbl(wynikSzukania.Parent.Cells(wynikSzukania.row, EVO.E_PIVOT_PROXY2____CLOE_COL_E_PRICE).Value)
        
        On Error Resume Next
        c1.Offset(0, 1).AddComment "colE: " & CStr(calculateCostForCloe) & Chr(10)
        
'        If c1.Offset(0, 1).Comment Is Nothing Then
'            c1.Offset(0, 1).AddComment "CLOE_COL_E_PRICE: " & CStr(calculateCostForCloe) & Chr(10)
'            c1.Offset(0, 1).Comment.Shape.TextFrame.AutoSize = True
'        Else
'            c1.Offset(0, 1).Comment.Text c1.Offset(0, 1).Comment.Text() & _
'                "CLOE_COL_E_PRICE: " & CStr(calculateCostForCloe) & Chr(10)
'            c1.Offset(0, 1).Comment.Shape.TextFrame.AutoSize = True
'        End If
    End If
End Function


Public Function metersForCloe(c2 As Range, Optional allRowForNothing As Range) As Double
    metersForCloe = 0#
    
    Dim sum2 As Double
    sum2 = 0#
    Do
        If IsNumeric(c2.Value) Then
            sum2 = sum2 + CDbl(c2.Value)
        End If
        
        
        Set c2 = c2.Offset(0, -1)
    Loop While c2.Column > 1
    
    If sum2 > 0 Then
        metersForCloe = sum2
        
        'On Error Resume Next
        'c2.AddComment "M: " & CStr(metersForCloe) & Chr(10)
        
'        If c2.Offset(0, 1).Comment Is Nothing Then
'            c2.Offset(0, 1).AddComment "M: " & CStr(metersForCloe) & Chr(10)
'            c2.Offset(0, 1).Comment.Shape.TextFrame.AutoSize = True
'        Else
'            c2.Offset(0, 1).Comment.Text c2.Offset(0, 1).Comment.Text() & _
'                "M: " & CStr(metersForCloe) & Chr(10)
'            c2.Offset(0, 1).Comment.Shape.TextFrame.AutoSize = True
'        End If
    End If
End Function
