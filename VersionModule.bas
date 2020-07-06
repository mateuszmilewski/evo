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


' 0.22
' ==============================================================
' init kpi feature
' KPI module
' kpi class with (KpiHandler):
' set master
' fillPartNumberDictionary (and some components)
' new class as component PartNumberItemForKpi for KPI logic
' ==============================================================\


' 0.23
' ==============================================================
' slight change on KPI logic from TRUE and FALSE into 1 and 0
' for more lean data for final kpi table
' as source for the charts
' ==============================================================


' 0.24
' ==============================================================
' pivot showing current week by default by if data is future
' there will be error that is why i added on error resume next:
' subs like: runMainLogicForCreationTheoreticPivotTable
' on err... before trying to add current yyyycw
' ==============================================================


' 0.25
' ==============================================================
' pivot coping pivot
' tcam report
' ==============================================================


' 0.26
' ==============================================================
' go from tcam report to proxy2
' ==============================================================


' 0.27
' ==============================================================
' first test with real price data on CLOE from control tower
' database - standard on col E is number 100%, but
' please check if there is some problem!
' ==============================================================


' 0.28
' ==============================================================
'
' isSat and isSun for clac DHxx for tcam logic - extra
' inside class CopyHandler inside LINE: 820! !!!
' ==============================================================



' 0.28
' ==============================================================
'
' AK - nie liczymy weekendow!
' dodatkowa implementacja:
'        For xo = 1 To howManyDaysOfOffset
'
'            tmpdate = tmpdate - 1
'
'            If isSunday(tmpdate) Then
'                tmpdate = tmpdate - 2
'            ElseIf isSaturday(tmpdate) Then
'                tmpdate = tmpdate - 1
'            End If
'
'        Next xo
' ==============================================================




' 0.30
' ==============================================================
'
' jest problem z COFOR EXPEDITEUR
' moze sie okazac, ze pomimo spasowania coforow i tak nie bedzie
' cos dzialac poprawnie zatem nalezalo by dodatkowo sprawdzic jeszcze
' matchi matchi rowniez po COFOR VENDEUR
' ==============================================================


' 0.31
' ==============================================================
'
' next feature for the prices and matching stuff...
' still.. this version have some gaps on tcam logic
' the dates DHxx for copy handler
' and dates in tcam pivot
' can be differently calculated - 1st prio to fix it!
' should for sure only one logic and some double check
' if its possible!


' new module and new class connected with TP04 logic!



' line 734 : copy handler
' readjustDatesAgainIfThereIsDifferentInPusMaster li, linesIter, master
' thing about TCAM logic!

' ==============================================================


' 0.32
' ==============================================================
'
' big refractorisation stuff with implementation from
' sq01 tool - new class SAP_Handler! - be careful - some
' side effects :'(

' ==============================================================


' 0.33
' ==============================================================
'
' predefined run for sq01 logic - to be a little bit faster
' in register: PRE_DEF_RUN_FOR_SQ01
' new button: getSq01DataWithPreDefParams

' ==============================================================

' 0.34
' ==============================================================
'
' small fixes on prev logic
' copy and paste special for removing starge working on
' Range().FIND() and FINDNEXT()

' ==============================================================



' 0.35
' ==============================================================
'
' massive trim into cstr!!! be careful now!

' ==============================================================



' 0.36
' ==============================================================
'
' register sheet removal on redundant column before T - was
' problem with TCAM
'
'
' SAP library EVO not working if somebody do not have access to
' SAP

' ==============================================================



' 0.37
' ==============================================================
'
' wykorzystanie UNITE 2 pod recalcul - osobna formula przeliczajaca
' z perspektywy pckg

' ==============================================================
