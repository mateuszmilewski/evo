Attribute VB_Name = "VersionModule"
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

' ==============================================================
' 0.01 basic scafold for the dev
' ==============================================================


' ==============================================================
' 0.02
' Export Dev Implementation Module for GitHub
' status form with the handling class reusable
' ==============================================================


' ==============================================================
' 0.03
'
' simple test mail handler
' ==============================================================


' ==============================================================
' 0.04
'
' simple verification on docinfo files
' ==============================================================


' ==============================================================
' 0.05
'
' copy data logic - initial stuff
' ==============================================================


' ==============================================================
' 0.06
'
' TMC basic logic
' ==============================================================


' ==============================================================
' 0.07
'
' starting point for tech doc from angel and hamid
' docinfomodule starts to be obsolete at this time
' ==============================================================


' ==============================================================
' 0.08
'
' TMC = 'non'
' part not found on red first column
' extra post formatting for the dates in DHEF and DHAS
' initial prototype for history/archive in comment

' ==============================================================


' ==============================================================
'0.09
'
' - if CPL cell empty do not overwrite
'yellow scenario - do not overwrite! - but leave note in comment
'manual ajdustment in PUS
'
'initial KPI for PUS MASTER
'
'
'
'case with TMC and PN -> (2 diff COFORS rememeber!!) rather then TMC and COFOR
'to correct 86945E  01
'
'
'if (ECHANCIER ONL (semaine) < TOday then
'    move? DHEF and DHAS?
'    1. ?
'
' ==============================================================


' ==============================================================
' 0.10
' dry in hard way
' ==============================================================


' ==============================================================
' 0.11
' optimise optimise
' ==============================================================


' ==============================================================
' 0.12
' FIX
'the new version works well except the pickup date if there is a weekend between the dates.
'
'Examples for body parts (ferrage/peinture) and this COFORS:
'94656L  01 ? its OK
'A007GT  01? its OK
'26805K  01? not OK
'A009YG  01? not OK
' ==============================================================



' ==============================================================
' 0.13
' first class for pivot handling
'
' error on formula - blocking field with route
'
' btn with remove colors and comments at the end
' ==============================================================



' 0.14
' ==============================================================
' small fix on working with visible data
' frist buttons for TCAM
' ==============================================================


' 0.15
' ==============================================================
' tcam initial logic
' information for you:
' there is a 3 diff ref range for each sepearte source worksheet
'Dim findByExpCoforInPle As Range
'Dim findByExpCoforInCloe As Range
'Dim findByExpCoforInMain As Range
' for coping there is new sub called:
' Public Sub copyForSourcePivot(ByRef ph As PivotHandler, sh As StatusHandler)
'
' also, as you can see there is a new param for pivot hadnler - important
'
' new standard for naming worksheets - more english rather angel root naming
'
' new class for:
' TRANSPORT CALCULATION AND MONITORING
'
' LineItemPivotSrouceSupplement - component suppling LineItem
' ==============================================================


' 0.16
' ==============================================================
' moved fields connected with TCAM into LineItemPivotSrouceSupplement
' new private sub inside CopyHandler:
' insideCopyForSourcePivot_setupOrderDate
' ==============================================================


' 0.17
' ==============================================================
' new fields in pivot source
'
' ! in Private Sub makeMySumIfs(sh As Worksheet)
' !! in CopyIterationForPivotSource class i've put static code with
' col A as "matchy"
' ==============================================================


' 0.18
' ==============================================================
' Force explicit variable declaration.
' Option Explicit On !!!
' change name makeMySumIfs na makeMyFormulas - change into static
' ==============================================================



' 0.19
' ==============================================================
' removed manquant on sap gui scripting lib
' ==============================================================



' 0.20
' ==============================================================
' initial logic for pivot
' ==============================================================


' 0.21
' ==============================================================
' runMainLogicForCreationPivotTable in PivotHandler
' runMainLogicForCreationTheoreticPivotTable in PivotHandler
' ==============================================================
