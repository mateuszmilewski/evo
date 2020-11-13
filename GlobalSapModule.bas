Attribute VB_Name = "GlobalSapModule"
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


' sap globals


' late binging necessary becuase there is some users without SIGAPP SAP...

Global sapGuiAuto As Object
'Global sapGui As SAPFEWSELib.GuiApplication
'Global cnn As SAPFEWSELib.GuiConnection
'Global sess As SAPFEWSELib.GuiSession
'Global window As SAPFEWSELib.GuiModalWindow
'Global c As SAPFEWSELib.GuiCollection
'Global item As SAPFEWSELib.GuiComponent
'Global Btn As SAPFEWSELib.GuiButton
'Global table As SAPFEWSELib.GuiTableControl ' ISapTableControlTarget
'Global gv As SAPFEWSELib.GuiGridView
'Global myTableRow As SAPFEWSELib.GuiTableRow
'Global coll As SAPFEWSELib.GuiCollection


Global sapGui As Object
Global cnn As Object
Global sess As Object
Global window As Object
Global c As Object
Global item As Object
Global Btn As Object
Global table As Object
Global gv As Object
Global myTableRow As Object
Global coll As Object

