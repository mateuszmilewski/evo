Attribute VB_Name = "DeleteSpecialModule"
Option Explicit



Public Sub deleteSelectedAndVisibleBeCarefulItIsNotRestorable()
    
    Dim toDel As Range
    Set toDel = Selection
    
    Set toDel = toDel.SpecialCells(xlCellTypeVisible)
    
    toDel.EntireRow.Delete xlShiftUp
End Sub
