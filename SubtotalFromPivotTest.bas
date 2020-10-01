Attribute VB_Name = "SubtotalFromPivotTest"
Option Explicit

Private Sub removeSubtotalFromPivot()
Attribute removeSubtotalFromPivot.VB_ProcData.VB_Invoke_Func = " \n14"
'
' removeSubtotalFromPivot Macro
'

'
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("ID"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "WIERSZ").Subtotals = Array(False, False, False, False, False, False, False, False, _
        False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("REF"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "COFOR_VENDEUR").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "COFOR_EXPEDITEUR").Subtotals = Array(False, False, False, False, False, False, False _
        , False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "NOM_FOURNISSEUR").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "DELIVERY DATE").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "DELIVERY YEAR").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "DELIVERY MONTH").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "DELIVERY WEEK").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "DELIVERY YYYYCW").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("AK"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "ORDER DATE").Subtotals = Array(False, False, False, False, False, False, False, False _
        , False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "ORDER YEAR").Subtotals = Array(False, False, False, False, False, False, False, False _
        , False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "ORDER MONTH").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "ORDER WEEK").Subtotals = Array(False, False, False, False, False, False, False, False _
        , False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "ORDER YYYYCW").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "ROUTE NAME AND PILOT").Subtotals = Array(False, False, False, False, False, False, _
        False, False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("QTY 2" _
        ).Subtotals = Array(False, False, False, False, False, False, False, False, False, False _
        , False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "CQTY 2").Subtotals = Array(False, False, False, False, False, False, False, False, _
        False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("UC 2") _
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("OQ 2") _
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "Confirmed OQ 2").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("CONDI" _
        ).Subtotals = Array(False, False, False, False, False, False, False, False, False, False _
        , False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("PC_GV" _
        ).Subtotals = Array(False, False, False, False, False, False, False, False, False, False _
        , False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("BPC"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("MC"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("MBU"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "MAX CAPACITY").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "(TN)(mL)").Subtotals = Array(False, False, False, False, False, False, False, False, _
        False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "Confirmed (TN)(mL)").Subtotals = Array(False, False, False, False, False, False, _
        False, False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("LQ"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "Confirmed LQ").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("RP"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "Confirmed RP").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "(TNbox)(mL)").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "Confirmed (TNbox)(mL)").Subtotals = Array(False, False, False, False, False, False, _
        False, False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "(RN)(mL)").Subtotals = Array(False, False, False, False, False, False, False, False, _
        False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "Confirmed (RN)(mL)").Subtotals = Array(False, False, False, False, False, False, _
        False, False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "ROUNDUP (RN)(mL)").Subtotals = Array(False, False, False, False, False, False, False _
        , False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "ROUNDUP Confirmed (RN)(mL)").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields("ORDER" _
        ).Subtotals = Array(False, False, False, False, False, False, False, False, False, False _
        , False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "DELIVERY").Subtotals = Array(False, False, False, False, False, False, False, False, _
        False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "TYPE_DE_PIECE").Subtotals = Array(False, False, False, False, False, False, False, _
        False, False, False, False, False)
    ActiveSheet.PivotTables("THEORETIC_PIVOT2_Pivot_20200917_").PivotFields( _
        "PRICE_FROM_CLOE_COL_E").Subtotals = Array(False, False, False, False, False, False, _
        False, False, False, False, False, False)
End Sub
