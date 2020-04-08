Attribute VB_Name = "EnumModule"
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


Public Enum E_FILE_TYPE
    E_MASTER_PUS
    E_FEED_CPL
End Enum

Public Enum E_MASTER_ORDER
    E_MASTER_Reference = 1
    E_MASTER_DESIGNATION
    E_MASTER_QTY
    E_MASTER_CQTY
    E_MASTER_ECHANCIER_ONL_S
    E_MASTER_cofor
    E_MASTER_Nom_fournisseur
    E_MASTER_prog_livraison
    E_MASTER_num_contrat
    E_MASTER_Nom_appro
    E_MASTER_Tel_appro
    E_MASTER_mail_appro
    E_MASTER_Nom_contact_log
    E_MASTER_Tel_contact_log
    E_MASTER_mail_contact_log
    E_MASTER_TMC
End Enum


Public Enum E_FORM_SCENARIO_TYPE
    E_FORM_SCENARIO_COPY_PASTE = 1
    E_FORM_SCENARIO_CREATE_PIVOT_SCENARIO
End Enum


Public Enum E_COPY_HANDLER_INIT
    E_COPY_HANDLER_COPY_ONE = 1
    E_COPY_HANDLER_BY_TMC_OPT
    E_COPY_HANDLER_FOR_PIVOT_CREATION
End Enum


Public Enum E_PIVOT_SRC
    E_PIVOT_SRC_ID = 1
    E_PIVOT_SRC_WIERSZ
    E_PIVOT_SRC_REF
    E_PIVOT_SRC_COFOR_VENDEUR
    E_PIVOT_SRC_COFOR_EXPEDITEUR
    E_PIVOT_SRC_SUPPLIER_NAME
    E_PIVOT_SRC_DELIVERY_DATE
    E_PIVOT_SRC_DELIVERY_YEAR
    E_PIVOT_SRC_DELIVERY_MONTH
    E_PIVOT_SRC_DELIVERY_WEEK
    E_PIVOT_SRC_DELIVERY_YYYYCW
    E_PIVOT_SRC_AK_COL_FROM_PLE
    E_PIVOT_SRC_ORDER_DATE
    E_PIVOT_SRC_ORDER_YEAR
    E_PIVOT_SRC_ORDER_MONTH
    E_PIVOT_SRC_ORDER_WEEK
    E_PIVOT_SRC_ORDER_YYYYCW
    E_PIVOT_SRC_ROUTE
    E_PIVOT_SRC_PILOT
    E_PIVOT_SRC_ROUTE_AND_PILOT
    E_PIVOT_SRC_QTY
    E_PIVOT_SRC_CQTY
    E_PIVOT_SRC_UC
    E_PIVOT_SRC_OQ
    E_PIVOT_SRC_COQ
    E_PIVOT_SRC_ROUNDUP_OQ1
    E_PIVOT_SRC_ROUNDUP_COQ1
    E_PIVOT_SRC_SUMIF_QTY
    E_PIVOT_SRC_SUMIF_CQTY
    E_PIVOT_SRC_SUMIF_UC
    E_PIVOT_SRC_SUMIF_OQ
    E_PIVOT_SRC_SUMIF_COQ
    E_PIVOT_SRC_CONDI
    E_PIVOT_SRC_UA_PC_GV
    E_PIVOT_SRC_UA_BPC
    E_PIVOT_SRC_UA_MC
    E_PIVOT_SRC_UA_MBU
    E_PIVOT_SRC_UA_MAX_CAPACITY
    E_PIVOT_SRC__TN_ML
    E_PIVOT_SRC__CONFIRMED_TN_ML
End Enum


Public Enum E_PIVOT_PROXY2
    E_PIVOT_PROXY2_ID = 1
    E_PIVOT_PROXY2_WIERSZ
    E_PIVOT_PROXY2_REF
    E_PIVOT_PROXY2_COFOR_VENDEUR
    E_PIVOT_PROXY2_COFOR_EXPEDITEUR
    E_PIVOT_PROXY2_SUPPLIER_NAME
    E_PIVOT_PROXY2_DELIVERY_DATE
    E_PIVOT_PROXY2_DELIVERY_YEAR
    E_PIVOT_PROXY2_DELIVERY_MONTH
    E_PIVOT_PROXY2_DELIVERY_WEEK
    E_PIVOT_PROXY2_DELIVERY_YYYYCW
    E_PIVOT_PROXY2_AK_COL_FROM_PLE
    E_PIVOT_PROXY2_ORDER_DATE
    E_PIVOT_PROXY2_ORDER_YEAR
    E_PIVOT_PROXY2_ORDER_MONTH
    E_PIVOT_PROXY2_ORDER_WEEK
    E_PIVOT_PROXY2_ORDER_YYYYCW
    E_PIVOT_PROXY2_ROUTE_AND_PILOT
    E_PIVOT_PROXY2_SUMIF_QTY
    E_PIVOT_PROXY2_SUMIF_CQTY
    E_PIVOT_PROXY2_SUMIF_UC
    E_PIVOT_PROXY2_SUMIF_OQ
    E_PIVOT_PROXY2_SUMIF_COQ
    E_PIVOT_PROXY2_CONDI
    E_PIVOT_PROXY2_UA_PC_GV
    E_PIVOT_PROXY2_UA_BPC
    E_PIVOT_PROXY2_UA_MC
    E_PIVOT_PROXY2_UA_MBU
    E_PIVOT_PROXY2_UA_MAX_CAPACITY
    E_PIVOT_PROXY2__TN_ML
    E_PIVOT_PROXY2__CONFIRMED_TN_ML
End Enum
