Attribute VB_Name = "TryToRenameModule"
Option Explicit

Public Function tryToRenameWorksheet(psh As Worksheet, prefix As String) As String

    tryToRenameWorksheet = psh.name
    
    Dim tmpNewName As String, mm As String, dd As String
    mm = CStr(Month(Date))
    dd = CStr(Day(Date))
    
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
    
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    
    ' On Error Resume Next
    tmpNewName = "" & CStr(prefix) & CStr(Year(Date)) & mm & dd & "_"
    
    
    On Error Resume Next
    psh.name = CStr(tmpNewName)
    
    innerJokeFoo psh, tmpNewName
    
    
    tryToRenameWorksheet = psh.name

End Function



Private Sub innerJokeFoo(ByRef psh As Worksheet, newName As String)

    On Error Resume Next
    psh.name = CStr(newName)
    
    If psh.name = newName Then
        Exit Sub
    Else
    
        If Len(newName) < 30 Then
            innerJokeFoo psh, newName & "I"
        Else
            Exit Sub
        End If
    End If
End Sub
