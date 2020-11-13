VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XtraCopyDataReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private timestamp As Date
Private w As Workbook
Private sh As Worksheet

Private d As Dictionary
Private c As Collection

Private Sub tryToAddTheLineItemToDictionary(li As LineItem)

    ' nie mam pomyslu co moglo by byc keyem dla line nokow
    'zatem narazie ten sub niech zostanie privem
End Sub

Public Sub addLiToCollection(li As LineItem)
    
    c.Add li
    
End Sub


Public Sub createXtraReport()

    Dim rw As Variant
    If Not c Is Nothing Then
        If c.count > 0 Then
    
            Set w = Workbooks.Add
            Set sh = w.Sheets.Add
            sh.name = "xtraNOKrep"
            
            With sh
                .Cells(1, 1).Value = "NOKs"
                rw = 2
                .Cells(rw, 1).Value = "li.sourceRng.Address"
                .Cells(rw, 2).Value = "li.targetRow"
                .Cells(rw, 3).Value = "PN"
                .Cells(rw, 4).Value = "COFOR_COFOR"
                .Cells(rw, 5).Value = " Full Log "
                

                rw = 3
                Dim li As LineItem
                For Each li In c
                    
                    If Not li.sourceRng Is Nothing Then
                        .Cells(rw, 1).Value = li.sourceRng.Address
                    End If
                    .Cells(rw, 2).Value = li.wiersz
                    .Cells(rw, 3).Value = li.pn
                    .Cells(rw, 4).Value = li.COFOR_COFOR
                    .Cells(rw, 5).Value = li.getFullLog()
                    rw = rw + 1
                Next li
            
            
                .columns("A:A").ColumnWidth = 12
                .columns("B:B").ColumnWidth = 12
                .columns("C:C").ColumnWidth = 19
                .columns("D:D").ColumnWidth = 19
                .columns("E:E").ColumnWidth = 22
                
                .UsedRange.WrapText = True
            End With
        End If
    End If
End Sub


Private Sub Class_Initialize()
    Set d = New Dictionary
    Set c = New Collection
    timestamp = Date
End Sub

Private Sub Class_Terminate()
    Set d = Nothing
    Set c = Nothing
End Sub
