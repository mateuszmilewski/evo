Attribute VB_Name = "TAWModule"
Option Explicit

'The MIT License (MIT)
'
'Copyright (c) 2021 FORREST
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



Private Function findTawWorkbook() As Workbook
    Set findTawWorkbook = Nothing
    
    
    Dim tmp As Workbook
    For Each tmp In Workbooks
        If Trim(tmp.ActiveSheet.Range("G14").Value) = "Table" Then
            If tmp.ActiveSheet.Range("G15").Value Like "*ZMATPLANT__ZARTMAITR*" Then
                Set findTawWorkbook = tmp
                Exit Function
            End If
        End If
    Next tmp
End Function


Public Sub MakeOut1FromTawQuery()


    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    
    Dim domainStr As String
    domainStr = InputBox("Get data from domain: ")
    
    Dim tawWorkbook As Workbook
    Set tawWorkbook = findTawWorkbook()
    
    
    
    
    If Not tawWorkbook Is Nothing Then
    
    
        Debug.Print tawWorkbook.FullName

    
        
        Dim numHandler As NumberHandler
        Set numHandler = New NumberHandler
        Dim ish As Worksheet, osh As Worksheet
        Dim irng As Range, orng As Range
        Set osh = ThisWorkbook.Sheets.Add
    
        
        osh.name = EVO.TryToRenameModule.tryToRenameWorksheet(osh, "OUT1_" & CStr(domainStr) & "_")
        
        
        
        ' LABLES ------------------------------------------------
        
        fillLabels osh.Range("A1")
        
        ' fillLabels inter4Sh.Range("A1")
        
        ' -------------------------------------------------------
        
    
        ' MAIN LOGIC FOR TAW
        Dim g As Range, r1 As Range, tmpForMoneyString As String
        Set r1 = osh.Range("A2")
        Set g = tawWorkbook.ActiveSheet.Range("G16")
        Do
        
            If Trim(g.offset(0, EVO.eTawQuery1.eTawQuery4DOM - 1).Value) = Trim(domainStr) Then
            
            
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_DOMAIN).Value = _
                    Trim(g.offset(0, EVO.eTawQuery1.eTawQuery4DOM - 1).Value)
            
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_ARTICLE).Value = _
                    Trim(g.offset(0, EVO.eTawQuery8REF - 1).Value)
                    
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_RU).Value = _
                    Trim(g.offset(0, EVO.eTawQuery1.eTawQuery14RU - 1).Value)
                    
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_DOC_ACHAT).Value = _
                    Trim(g.offset(0, EVO.eTawQuery1.eTawQuery5DOC - 1).Value)
                    
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_TYPE).Value = _
                    ""
                    
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_DIV).Value = _
                    Trim(g.offset(0, EVO.eTawQuery1.eTawQuery2PLT - 1).Value)
                    
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_FOUR).Value = _
                    Trim(g.offset(0, EVO.eTawQuery1.eTawQuery16FOUR - 1).Value)
                    
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_DATE_DEBUT).Value = _
                    Trim(g.offset(0, EVO.eTawQuery1.eTawQuery10DEBUT - 1).Value)
                    
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_DATE_FIN).Value = _
                    Trim(g.offset(0, EVO.eTawQuery1.eTawQuery11FIN - 1).Value)
                    
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_UNITE).Value = _
                    Trim(g.offset(0, EVO.eTawQuery1.eTawQuery18UN - 1).Value)
                
                tmpForMoneyString = CStr(g.offset(0, EVO.eTawQuery1.eTawQuery19SUM - 1).Text)
                
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_SUM).Value = _
                    CDbl(g.offset(0, EVO.eTawQuery1.eTawQuery19SUM - 1).Value)
                    
                osh.Cells(r1.row, EVO.E_FROM_SQ01_QUASI_TP04.E_FROM_SQ01_QUASI_TP04_CURRENCY).Value = _
                    CStr(Right(Trim(tmpForMoneyString), 3))
                    
                    

                
                Set r1 = r1.offset(1, 0)
            End If
        
            Set g = g.offset(1, 0)
        Loop Until Trim(g.Value) = ""
        
        
        
        
        
        ' COPY AND PASTE AS VALUES ------------------------------
        
        ' ???
        ' copyAndPasteAsValues osh.Range("A1").Offset(0, E_FROM_SQ01_QUASI_TP04_ARTICLE - 1)
        
        ' -------------------------------------------------------
        
        ' data ready - change string price into normal num
        changePricesIntoDouble osh
        
        
        Set numHandler = Nothing
        
        
        '---------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------
    End If
End Sub


