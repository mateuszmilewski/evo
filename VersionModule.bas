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



' 0.38
' ==============================================================
'
' welcome
' fixes on register

' ==============================================================





' 0.39
' ==============================================================
'
' changes on TCAM report
' based on information of Caroline!
' a little bit different approach on data
' new values, new columns etc...
' shipper cofor + supplier name
' Change meters to condi and boxes
' brak danych ple, cloe or main -> musi byc drugi raport ktory pokazuje braki!


' CopyHandler line 745: starts from here extra worksheet - but only flag
' for nok report - missing data - nto available for pivot!
' for pivot proxy2 -> and finally from TCAM!
' later doubleLoop... sub for treatment
'
'Private Sub doubleLoopForNOKrepForTCAMLogic(mmk1 As Variant, mmk2 As Variant,
'    ByRef mmexpCofors As Dictionary,
'    ByRef mmlinesIter As Dictionary,
'    mmli As LineItem)


' new NOK report:
' TCAM_NOKs_YYYYMMDD_ checking availablity for CONDI, CLOE nad PLE
' creating after clicking craete source for pivot
' in parallel with source and proxy2


' ==============================================================




' 0.39
' ==============================================================
'
' zmiana na ECHANCIER ONL (semaine) z S??/?? na CW??
' dodatkowe sprawdzenie


' ==============================================================

' 0.40 + 0.41
' ==============================================================
' minor changes in formatting
' new feature for internal fourniseures ...
' new transaction from sigapp
' Y_PI1_80000391
' newSh.PasteSpecial Paste:=xlPasteValues - in copy to proxy2!
' ==============================================================




' 0.42
' ==============================================================
'
' fix on tcam logic- proxy2 stopped in the middle!
' now it should work!
' ==============================================================


' 0.43
' ==============================================================
' preliminary feature for forecase pp1x so called green light
' ==============================================================


' 0.44
' ==============================================================
' tcam update - fix about CLOE NOK and PLE NOK
' ==============================================================



' 0.45
' ==============================================================
' mb51 now inside
' some changes in sq01 and internal parts - seperate please!
' ==============================================================

' 0.46
' ==============================================================
' adjusting layouts for final list
' ==============================================================



' 0.47
' ==============================================================
' final touch for reports green light and reception
'
' new final touch modules
'
' ==============================================================


' 0.48
' =============================================================
' final touch for reception ready
' + green light dev in this version
' + starting point for final report leaf
' input def synthese preparation
'
' ==============================================================


' 0.49
' ==============================================================
' quick side fix for DHxx copy
' 1324 CopyHandler - DHxx + DAP + DDP logic - smth wrong here
' ==============================================================


' 0.50
' ==============================================================
' quick side fix for DHxx copy - going back to org logic
' extra check in BASE CPL - if there is a dates - chose them
' first!
' ==============================================================


' 0.50
' ==============================================================
' combo time - one form rulez
' somea automatic recognition on data
' for example for RU - reception can have data only when
' in EVO there is some data from source from green light
' quite complicated for reception part
' but i will just made simple loop and check trough whole evo

' ==============================================================


' 0.51
' ==============================================================
' combo time - one form rulez
' extra logic for smart lookup custom made

' ==============================================================




' 0.52 + 0.53
' ==============================================================
' small fixes
' +
' managers list
' Y_DI3_8000594 - not done inside ManagersDAModule - tbd

' ==============================================================



' 0.54
' ==============================================================
'
' managers list
' Y_DI3_8000594 = tbc

'
' change logic for calc DHxx - now need to have
'
' ECHANCIER ONL (semaine) - question before copy data


' ==============================================================



' 0.55 + 0.56
' ==============================================================
'
' managers da logic initial + final touch  buttons
' starting logic for combo forecast
' ==============================================================


' 0.57
' ==============================================================
'
' lot of mistakes in version < 0.56 > 0.55
' this version run successfully on PPx1 - lot of fixes
' on reception issue with ENUMS !!

' ==============================================================


' 0.58
' ==============================================================
'
' going back to TCAM volume 2

' ==============================================================


' 0.58 + 0.59
' ==============================================================
'
' UN issue - wrong column taken into cosideration!
' there is 2 columns UN: UQs and UNx - UQs is the right one!

' ==============================================================


' 0.60
' ==============================================================
'
' prepa for auto managers da
' new columns in green light! (domain, ru, div)
' to auto recognize where i should look for
' managers in Y_DI3_80000594

' ==============================================================


' 0.61 + 0.62
' ==============================================================
'
' sq01 adjusted so tp04 prefix - new column at the end
' manager da

' partial solution! - some issues with OLE waiting too long
' maybe i should ignore this issue with RU = 4 PM confirmed
'
' recheck pivot 2 logic making final tcam logic
'

' ==============================================================



' 0.63
' ==============================================================
'
' initiall feeback logic from mb51 for tp04
'
'
' ==============================================================


' 0.64
' ==============================================================
'
' ModelessLeaf
'
'
' ==============================================================


' 0.64 + 0.65 + 0.66 + 0.67
' ==============================================================
'
' change in pus - now version 2.6 - change on ECHANCIER ONL (semaine)
' new layout from col E YY-CWXX
' have issue with TMC - who will be now optimising?
' '
' new standard on
' fortunanetly i have good q on code
' implementation provided is in operations handler...
' there is dry logic:
' .yyyycw = CLng(operacja.calculateYYYYCW(CStr(CStr(r.Offset(0, EVO.E_MASTER_ECHANCIER_ONL_S - 1).Value))))
' every echancier onl data will first of all calc into std yyyycw
'
' Public Function calculateYYYYCW(s As String) in 227 in OperationsHandler class
'
'
' ==============================================================


' 0.68
' ==============================================================
'
' ModelessLeaf - activate the export button for making new
' workbook which already provides one leaf of the report
' ==============================================================


' 0.69 + 0.70
' ==============================================================
'
' 069 backup version
' 070 finally first vresion of combo form for green light data
' ==============================================================


' 0.71
' ==============================================================
'
' 071 backup version
' ' continue with green light combo implementation
' MODULE: ComboGreenLightModule initially ready
'  'FORM: ComboFormGreenLightReport - version 1.0 ready
''' '(compatible with the current setup of reception and green light
' ==============================================================


' 0.72
' ==============================================================
'
' 072 backup version
' + reception new column!
' ==============================================================



' 0.73
' ==============================================================
'
' 073 backup version
' final touch on green light re-check auto
' ==============================================================


' 0.74
' ==============================================================
'
' 073 backup version
' final touch on green light re-check auto - extra fine tunning
' ==============================================================



' 0.75
' ==============================================================
' ==============================================================
' big change on runMainLogicForSQ01__with_preDef ' clearing and putting again
' ==============================================================
' ==============================================================


' 0.76
' ==============================================================
'
' heavy cating on types in ManagersDaModule
' for green light adjusted worksheet
' to be sure that data will match
' ==============================================================



' 0.77 + 0.78
' ==============================================================
'
' split list for sq01 to not block OLE logic
' ==============================================================

' 0.79
' ==============================================================
' make modeless leaf a little but smarter - some logic
' not only step by step - create leaf
' but make all with one click - create leafs
' should have some basic cfg and loop for create leaf x times
' and at the end there will be as well some total scaffold logic
' ==============================================================



' 0.80
' ==============================================================
' backup with ready make new for modeless leaf
' ==============================================================


' 0.81
' ==============================================================
' late binding to the SAP lib to avoid errors for users
' which are without SAP
'
' +++++
' new columns in source PIVOT - LUCIE feature
' ==============================================================



' 0.82
' ==============================================================
' final touch extension for colour frmtting + auto synthesis
' 0.83
' ==============================================================
' hot-fix for modeless leaf not starting from first Cw
' hot-fix interrocom currency issue still
' ==============================================================


' 0.84 + 0.85
' ==============================================================
' initial version of new PUS version with extra columns  from PM
' also minor changes for COPY DATA LOGIC
' ==============================================================



' 0.852
' ==============================================================
' huge fix for internal part of screen mb51 - still layout problem
' and order of the columns inside sap - now fixed hard with
' internal name of the each field - the order is random
' each week so be careful:
'
'0 MATNR -100550588
'1 MAKTX - Agrafes ARaymond
'2 WERKS -5820
'3 LGORT -3770
'4 KOSTL -
'5 BWART -101
'6 MBLNR -5000096857#
'7 CHARG -2506007
'8 MENGE -1
'9 DMBTR -7.74, 0
'10 BUDAT - 22.10.2020
'11 CPUDT - 22.10.2020
'12 BLDAT - 22.10.2020
'13 CPUTM - 11:26:02
'14 ERFME -UN
'15 BPRME -UN
'16 WAERS -EUR
'17 BUKRS -260
'18 BWTAR -2506007
'19 EBELP -30
'20 EXBWR -0, 0
'21 GRUND -
'22 KDAUF -
'23 KDPOS -
'24 KUNNR -
'25 MJAHR -2020
'26 VORNR -
'27 PSPID -
'28 SHKZG -s
'29 XABLN -
'30 NAME1 - ONL MADRID
'31 BTEXT - EM Entrée marchand.
'32 SOBKZ -
'33 ZEILE -1
'34 ERFMG -1
'35 ANLN1 -
'36 APLZL -
'37 AUFPL -
'38 BPMNG -1
'39 BSTME -UN
'40 BSTMG -1
'41 LONGNUM -
'42 EXVKW -0, 0
'43 KDEIN -
'44 KZBEW -b
'45 KZVBR -
'46 KZZUG -
'47 MEINS -UN
'48 NPLNR -
'49 RSNUM -
'50 RSPOS -
'51 USNAM -U313961
'52 VGART -WE
'53 VKWRT -0, 0
'54 XAUTO -
'55 AUFNR -
'56 XBLNR -ADLC5081
'57 EBELN -3939539113#
'58 LIFNR - 98780U  01
'59 ANLN2 -
' ==============================================================


' 0.86
' ==============================================================
' fix on mb51 : decimal separator vs grouping separator
' ==============================================================


' 0.87
' -
' version with readme -ReadMeButtonsModule
' -
' modeless leaf add for existing one



'0.88
' ==============================================================
' li.addRawInfoToLog "this PN is missing in CPL BASE: Control Tower DB!"
' ==============================================================

'0.89
' ==============================================================
' fine tunning the reception ppx1 add to existing one!
' ==============================================================


' 0.90
' ==============================================================
' hot fix for calcul in DHxx between dates
' ' fine tunning the reception ppx1 add to existing one! 2nd
' ==============================================================
