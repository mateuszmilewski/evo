Attribute VB_Name = "ChartsModule"
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

Sub addChart1()
Attribute addChart1.VB_Description = "This is test for adding new chart which is default"
Attribute addChart1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' addChart1 Macro
' This is test for adding new chart which is default
'

'
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("KPI_20200518_!$B$4:$D$6")
    'ActiveSheet.Shapes("Chart 1").IncrementLeft -145.5
    'ActiveSheet.Shapes("Chart 1").IncrementTop -92.25
End Sub


Public Sub addCharts(sh As Worksheet, rng1 As Range)

    
    Dim myTempChart1 As Chart, myTempChart2 As Chart, myTempChart3 As Chart
    
    
    sh.Shapes.AddChart2(201, xlColumnClustered, rng1.offset(0, 4).Left, rng1.Top, 200, 100).Select
    ActiveChart.SetSourceData sh.Range(rng1, rng1.offset(2, 2))
    ActiveChart.HasTitle = False
    ActiveChart.HasLegend = True
    
    
    Set rng1 = rng1.End(xlDown).End(xlDown)
    
    sh.Shapes.AddChart2(201, xlColumnClustered, rng1.offset(0, 4).Left, rng1.offset(2, 3).Top, 200, 100).Select
    ActiveChart.SetSourceData sh.Range(rng1, rng1.offset(2, 2))
    ActiveChart.HasTitle = False
    ActiveChart.HasLegend = True

    Set rng1 = rng1.End(xlDown).End(xlDown)
    
    sh.Shapes.AddChart2(201, xlColumnClustered, rng1.offset(0, 4).Left, rng1.offset(5, 3).Top, 200, 100).Select
    ActiveChart.SetSourceData sh.Range(rng1, rng1.offset(2, 2))
    ActiveChart.HasTitle = False
    ActiveChart.HasLegend = True

End Sub
