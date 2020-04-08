Attribute VB_Name = "ExportThisProjectMod"
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

' working great!
Global Const REPO_PATH = "C:\WORKSPACE\dev\c41_tools\evo\repo\"

Private Sub export_this_project()


    Dim txt As String
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ThisWorkbook.VBProject
    Set VBComps = VBProj.VBComponents
    
    For Each VBComp In VBComps
        
        If VBComp.Type = vbext_ct_StdModule Then
            txt = VBComp.Name & ".bas"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        ElseIf VBComp.Type = vbext_ct_ClassModule Then
            txt = VBComp.Name & ".cls"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        ElseIf VBComp.Type = vbext_ct_MSForm Then
            txt = VBComp.Name & ".frm"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        End If
         
    Next VBComp
    
    MsgBox "ready!"

End Sub


Private Sub import_this_project()
    
    
    remove_current_implementation
    
    
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Set objFSO = New Scripting.FileSystemObject
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ThisWorkbook.VBProject
    Set VBComps = VBProj.VBComponents
    
    For Each objFile In objFSO.GetFolder(REPO_PATH).Files
        ' body
        ' ==============================================================
        
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            VBComps.Import objFile.Path
        End If
        
        ' ==============================================================
    Next objFile
    
    MsgBox "ready!"

End Sub


Private Sub remove_current_implementation()


    Dim txt As String
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ThisWorkbook.VBProject
    Set VBComps = VBProj.VBComponents
    
    For Each VBComp In VBComps
        
        If VBComp.Type = vbext_ct_Document Then
            txt = VBComp.Name
            Debug.Print txt & " zostaje"
            
        ElseIf VBComp.Type = vbext_ct_ActiveXDesigner Then
            txt = VBComp.Name
            Debug.Print txt & " zostaje"

        ElseIf CStr(VBComp.Name) = "ExportThisProjectMod" Then
            txt = VBComp.Name
            Debug.Print txt & " zostaje"
        Else
            
            VBComps.Remove VBComp
        End If
         
    Next VBComp
    
    ' MsgBox "ready!"

End Sub
