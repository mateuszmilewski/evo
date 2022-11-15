VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAP_Handler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


' OK - we need to change it into late binging becuase EVO users
' might do not have any SAP app

Private sapGuiAuto As Object
'Private sapGui As SAPFEWSELib.GuiApplication
'Private cnn As SAPFEWSELib.GuiConnection
'Private sess As SAPFEWSELib.GuiSession
'Private window As SAPFEWSELib.GuiModalWindow
'Private c As SAPFEWSELib.GuiCollection
' Private item2 As SAPFEWSELib.GuiComponent
' Private item As Variant ' double naming
'Private Btn As SAPFEWSELib.GuiButton
' Private table2 As SAPFEWSELib.GuiTableControl ' ISapTableControlTarget
Private table2 As Variant
'Private grid As SAPFEWSELib.GuiGridView
'Private gv As SAPFEWSELib.GuiGridView
'Private myTableRow As SAPFEWSELib.GuiTableRow
'Private coll As SAPFEWSELib.GuiCollection

'Private chbx As SAPFEWSELib.GuiCheckBox
'Private txt As SAPFEWSELib.GuiTextedit

'Private tf As SAPFEWSELib.GuiTextField
'Private lt As SAPFEWSELib.GuiLabel
Private tf As Variant
Private lt As Variant
    
Private sapGui As Object
Private cnn As Object
Private sess As Object
Private window As Object
Private c As Object
Private item As Object
Private Btn As Object
Private table As Object
Private grid As Object
Private gv As Object
Private myTableRow As Object
Private coll As Object
    
Private chbx As Object
Private txt As Object


' !!
Private cols As Variant

' params for table inside sap
Private y As Variant
Private x As Variant
Private x2 As Variant
Private x3 As Variant







Private Sub Class_Initialize()

    If EVO.GlobalSapModule.sapGuiAuto Is Nothing Then
        Set sapGuiAuto = GetObject("SAPGUI")
        Set EVO.GlobalSapModule.sapGuiAuto = sapGuiAuto
        Set sapGui = sapGuiAuto.GetScriptingEngine
        Set EVO.GlobalSapModule.sapGui = sapGui
    Else
        Set sapGuiAuto = EVO.GlobalSapModule.sapGuiAuto
        Set sapGui = EVO.GlobalSapModule.sapGui
    End If
    
    Set cnn = sapGui.Connections(0)
    
    Debug.Print cnn.ConnectionString
    Debug.Print cnn.Sessions.count

    
    Set sess = cnn.Children(0)
    Set item = sess.Children(0)
    Debug.Print item.name
    
    Debug.Print sess.Children.count
    
    
    ' Set item = sess.Children(0)
    Set item = sess.FindById("wnd[0]/usr")
    Debug.Print item.Children.count

    
    
    ' sess.FindById("wnd[0]").Maximize
End Sub

Private Sub Class_Terminate()
    Set sapGuiAuto = Nothing
    Set sapGui = Nothing
    Set cnn = Nothing
    Set sess = Nothing
    Set item = Nothing
End Sub



Public Sub justTouch()
    
    ' Debug.Print item.name
    On Error Resume Next
    sess.FindById("wnd[0]/tbar[0]/btn[12]").press
End Sub

Public Sub justTouch2()
    
    ' Debug.Print item.name
    Dim btnInner As Variant
    On Error Resume Next
    Set btnInner = sess.FindById("wnd[0]/tbar[0]/btn[3]")
    
    
End Sub

Public Sub runMainLogicFor__ME33K_mass(r As Range, Optional txtForDiv As String)
    
    If r.Parent.name Like "GREEN_LIGHT_*" Then
    
        Dim ash As Worksheet
        Set ash = r.Parent
    
        If ash.Range("A1").Value = "REFERENCE" And ash.Range("AF1").Value = "DOMAIN" Then
        
            If r.Cells(1, 1).Column = 1 Then
                
                 '------------------------------------------------------
                 
                 Dim ir As Range
                 Dim idForMag As String, idForArticle As String, idForDiv As String
                 Dim articleWoIndex As String, domain As String, mag As String
                 
                 Dim x As Variant, y As Variant
                 
                 Dim innerColl As Collection
                 Dim tmpStr As String, prevStr As String
                 
                 Dim directDataRow As Variant, el As Variant
                 
                 For Each ir In r
                 
                 
                 
                    If ir.EntireRow.Hidden = True Then
                        ' no operation for filtered out
                    Else
                    
                                        
                        idForMag = _
                            "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]"
                        idForArticle = _
                            "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[5,24]"
                            
                        idForDiv = _
                            "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[3,24]"
                            
                        
                        
                        articleWoIndex = CStr(Split(CStr(ir.Value), "-")(0))
                        domain = CStr(ash.Range("AF" & CStr(ir.row)))
                        mag = domain & "0"
                        
                        sess.FindById("wnd[0]/tbar[0]/okcd").Text = "ME33K"
                        sess.FindById("wnd[0]").sendVKey 0
                        sess.FindById("wnd[0]").sendVKey 4
                        
                        ' tab 20 == possibility to put articles!
                        sess.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020").Select
                        ' sess.findById("wnd[0]/tbar[1]/btn[8]").press
                        
                        If Trim(mag) = "0" Then
                            sess.FindById(idForMag).Text = "*"
                        Else
                            sess.FindById(idForMag).Text = mag
                        End If
                        ' sess.findById(idForArticle).Text = "98398447XT"
                        sess.FindById(idForArticle).Text = articleWoIndex
                        
                        
                        ' new for 107
                        sess.FindById(idForDiv).Text = CStr(txtForDiv)
                        
                        'sess.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]").SetFocus
                        'sess.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]").caretPosition = 4
                        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
                        
                        
                        
                        
                        x = 1
                        y = 1
                        
                        Set lt = Nothing
                        On Error Resume Next
                        Set lt = sess.FindById("wnd[1]/usr/lbl[" & CStr(y) & "," & CStr(x) & "]")
                        
                        If lt Is Nothing Then ' still
                            sess.FindById(idForMag).Text = "*"
                            sess.FindById("wnd[1]/tbar[0]/btn[0]").press
                        End If
                        
                        
                        ' example of output
                        
                        ' row 8 as first ' v on 21 col
                        ' 1, 1 "division" ' v on 21 col
                        ' 2, 1"mag" ' v from mag starts like everything else on 21 col
                        ' Catég. doc. d'achat x 3 y 1 'K x 3 y 21
                        ' Type document achat x 4 y 1 ' YPRS x 4 y 21
                        
                        ' x = 5 empty
                        
                        ' beg of table
                        ' Doc achat x 6 y 1
                        ' Poste x 6 y 12
                        ' Article x 6 y 18
                        ' Fourn. x 6 y 37
                        
                        ' x = 7 empty
                        
                        ' x = 8 content starting
                        '
                        ' 3939452822 x 8 y 1 ( doc achat )
                        ' 00010 x 8 y 12 ( poste )
                        ' 9830943980 x 8 y 18 ( article )
                        ' 50871U  01 x 8 y 37 ( cofor)
                        
                        
                        
                        Set innerColl = Nothing
                        Set innerColl = New Collection
                        
                        tmpStr = ""
                        prevStr = ""
                        
                        
                        For x = 1 To 20
                            
                            tmpStr = ""
                            prevStr = ""
                        
        
                            
                            For y = 1 To 100
                            
                                If y = 1 Then prevStr = ""
                                
                            
                                Set lt = Nothing
                                On Error Resume Next
                                Set lt = sess.FindById("wnd[1]/usr/lbl[" & CStr(y) & "," & CStr(x) & "]")
                                If Not lt Is Nothing Then
                                    ' Debug.Print lt.Text & " x " & x & " y " & y
                                    
                                    tmpStr = tmpStr & lt.Text & ";"
                                    prevStr = lt.Text
                                    
                                End If
                            Next y
                            
                            innerColl.Add tmpStr
                        Next x
                        
                        
    
                        Dim pozycja As Integer ' , lesposib
                        pozycja = 1
                        
                        For Each el In innerColl
                            
                            If CStr(el) Like "*" & Trim(ir.offset(0, 5).Value) & "*" Then
                                If CStr(el) Like "*YPRS*" Then
                                    Exit For
                                Else
                                    Debug.Print el
                                    Exit For
                                End If
                            End If
                            
                            pozycja = pozycja + 1
                            
                        Next el
                        
                        If pozycja > innerColl.count Then pozycja = 3
                        
                        
                        Debug.Print "pozycja: " & CStr(pozycja)
                        directDataRow = pozycja
                        sess.FindById("wnd[1]/usr/lbl[1," & CStr(directDataRow) & "]").SetFocus
                        sess.FindById("wnd[1]/usr/lbl[1," & CStr(directDataRow) & "]").caretPosition = 1
                        
                        
                        
                        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
                        'sess.findById("wnd[0]").sendVKey 0
                        
                        sess.FindById("wnd[0]/tbar[1]/btn[5]").press
                        
                        
                        ' ISapCTextField
                        ' Debug.Print TypeName(sess.findByid("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-BPRME[9,1]"))
                        
                        ' session.findById("wnd[0]/usr/ctxtRM06E-EVRTN")
                        Debug.Print TypeName(sess.FindById("wnd[0]/usr/ctxtRM06E-EVRTN"))
                        Debug.Print TypeName(sess.FindById("wnd[0]/usr/ctxtEKKO-LIFNR"))
                            
                        Set tf = Nothing
                            
                        On Error Resume Next
                        Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-BPRME[9,1]")
                        
                        Debug.Print tf.Text
                        Debug.Print tf.id
                        
                        
                        
                        ' very first field:  contract
                        Set tf = Nothing
                        Set tf = sess.FindById("wnd[0]/usr/ctxtRM06E-EVRTN")
                        Debug.Print tf.Text
                        Debug.Print tf.id
                        
                        
                        
                        ' fournisseur
                        Set tf = Nothing
                        Set tf = sess.FindById("wnd[0]/usr/ctxtEKKO-LIFNR")
                        Debug.Print tf.Text
                        Debug.Print tf.id
                        
                        
                        
                        ' date contr
                        Set tf = Nothing
                        Set tf = sess.FindById("wnd[0]/usr/ctxtRM06E-VEDAT")
                        Debug.Print tf.Text
                        Debug.Print tf.id
                        
                        
                        
                        ' type de contrat
                        Set tf = Nothing
                        Set tf = sess.FindById("wnd[0]/usr/ctxtRM06E-EVART")
                        Debug.Print tf.Text
                        Debug.Print tf.id
                        
                        
                        ' devise / currency
                        Set tf = Nothing
                        Set tf = sess.FindById("wnd[0]/usr/ctxtEKKO-WAERS")
                        Debug.Print tf.Text
                        Debug.Print tf.id
                        
                        
                        
                        ' 3 or 4 column jako artykul zero bez indexu
                        Set tf = Nothing
                        Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,0]")
                        Debug.Print tf.Text
                        Debug.Print tf.id
                        
                        
                        ' 3 or 4 column jako artykul - jeden nizej z indexem
                        Set tf = Nothing
                        Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,1]")
                        Debug.Print tf.Text
                        Debug.Print tf.id
                        
                        ' jesli tylko dwa, to to powinno byc puste!
                        Set tf = Nothing
                        Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,2]")
                        Debug.Print tf.Text & " tutaj powinno byc pusto!" ' al jest cos takiego: "__________________"
                        Debug.Print tf.id
                        
                        Dim underscores As String
                        underscores = "__________________"
                        
                        
                        Dim ajtem As ME33K_Item
                        Dim collectionOfItems As Collection
                        Set collectionOfItems = Nothing
                        Set collectionOfItems = New Collection
                        
                        For x = 0 To 12
                        
                        
                            Set ajtem = New ME33K_Item
                        
                            Set tf = Nothing
                            y = 3
                            On Error Resume Next
                            Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[" & CStr(y) & "," & CStr(x) & "]")
                                
                            If Not tf Is Nothing Then
                                If tf.Text = "__________________" Then
                                    Exit For
                                Else
                                
                                    ajtem.article = tf.Text
                                    ' session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-NETPR[7,0]")
                                    
                                    
                                    ' EUR
                                    ' ====================================================
                                    Set tf = Nothing
                                    Set tf = sess.FindById("wnd[0]/usr/ctxtEKKO-WAERS")
                                    On Error Resume Next
                                    ajtem.curr = tf.Text
                                    ' ====================================================
                                    
                                    
                                    
                                    ' PRICE
                                    ' ====================================================
                                    ' ====================================================
                                    Set tf = Nothing
                                    y = 7
                                    On Error Resume Next
                                    Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-NETPR[" & CStr(y) & "," & CStr(x) & "]")
                                    
                                    If Not tf Is Nothing Then
                                        ajtem.strPrice = tf.Text
                                        ajtem.parsePrice
                                        
                                        
                                        Set tf = Nothing
                                        y = 9
                                        ' On Error Resume Next
                                        Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-BPRME[" & CStr(y) & "," & CStr(x) & "]")
                                        If Not tf Is Nothing Then ajtem.up = tf.Text
                                        
                                        ajtem.makePriceByUnit
                                        
                                        
                                        collectionOfItems.Add ajtem
                                    End If
                                    ' ====================================================
                                    ' ====================================================
                                    
                                End If
                            End If
                        Next x
                        
                        
                        
                        
                        ' ISapTableControlTarget
                        ' Debug.Print TypeName(sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220"))
                        
                        'Set table2 = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220")
                        '
                        '
                        'x = 3
                        'y = 1
                        'Set item2 = table2.GetCell(x, y)
                        '
                        'Debug.Print item2.id & " " & TypeName(item2)
                        '
                        'Set tf = item2
                        'Debug.Print item2.Text
                        
                        
                        'For Each ajtem In collectionOfItems
                        '    Debug.Print "ajtem: " & ajtem.article & " " & ajtem.price
                        'Next ajtem
                        
                        
                        
                        ' ir jako column A
                        ' final green light re-check price
                        
                        Dim e1 As E_GREEN_LIGHT, e2 As E_GREEN_LIGHT
                        e1 = E_GREEN_LIGHT_PRE_SERIAL_PRICE_YPRS_contract
                        e2 = E_GREEN_LIGHT_INITIAL_SERIAL_PRICE
                        
                        Dim preSerialPriceRange As Range, initialSerialPriceRange As Range
                        
                        Set preSerialPriceRange = ir.offset(0, e1 - 1)
                        
                        preSerialPriceRange.Font.Bold = True
                        preSerialPriceRange.Font.Color = RGB(0, 0, 0)
                        preSerialPriceRange.Interior.Color = RGB(255, 255, 255)
                        
                        Set initialSerialPriceRange = ir.offset(0, e2 - 1)
                        
                        initialSerialPriceRange.Font.Bold = True
                        initialSerialPriceRange.Font.Color = RGB(0, 0, 0)
                        initialSerialPriceRange.Interior.Color = RGB(255, 255, 255)
                        
                        Debug.Print "pre: " & preSerialPriceRange.Address & " " & preSerialPriceRange.Value
                        Debug.Print "ini: " & initialSerialPriceRange.Address & " " & initialSerialPriceRange.Value
                        
                        
                        If collectionOfItems.count > 0 Then
                        
                        
                            
                            Dim txtForPreSerialFromME33K As String, xtraComment As String
                            
                            txtForPreSerialFromME33K = ""
                            xtraComment = ""
                            
                            txtForPreSerialFromME33K = txtForPreSerialFromME33K & CStr(Date) & Chr(10) & _
                                "data from ME33K: " & Chr(10)
                        
                            For Each ajtem In collectionOfItems
                                txtForPreSerialFromME33K = txtForPreSerialFromME33K & Chr(10) & _
                                    ": " & ajtem.article & " " & ajtem.price
                                    
                                    
                                If CStr(ajtem.article) = CStr(ir.Value) Then
                                    ' w index
                                    If (1# * (preSerialPriceRange.Value / ajtem.price)) > 1.1 Then
                                        ' out of margin
                                        ' MsgBox "price on me33k is diff for " & CStr(ir.Value) & "!", vbCritical
                                        
                                        xtraComment = Chr(10) & "pre-serial price issue: " & Chr(10) & _
                                            "me33k pre-serial price: " & ajtem.price & Chr(10) & _
                                            "tp04: " & CStr(preSerialPriceRange.Value)
                                            
                                        preSerialPriceRange.Font.Color = RGB(255, 0, 0)
                                        preSerialPriceRange.Value = "" & ajtem.price
                                    Else
                                        preSerialPriceRange.Font.Color = RGB(0, 255, 0)
                                    End If
                                    
                                    If CLng(preSerialPriceRange.Value) = 0 Then
                                        preSerialPriceRange.Value = ajtem.price
                                    End If
                                End If
                                
                                If CStr(ajtem.article) = CStr(Split(ir.Value, "-")(0)) Then
                                    ' wo index
                                    If (1# * (initialSerialPriceRange.Value / ajtem.price)) > 1.1 Then
                                        ' out of margin
                                        ' MsgBox "price on me33k is diff for " & CStr(ir.Value) & "!", vbCritical
                                        
                                        xtraComment = Chr(10) & "initial price issue: " & Chr(10) & _
                                            "me33k initial price: " & ajtem.price & Chr(10) & _
                                            "tp04 initial price: " & CStr(initialSerialPriceRange.Value)
                                            
                                        initialSerialPriceRange.Font.Color = RGB(255, 0, 0)
                                    Else
                                        initialSerialPriceRange.Font.Color = RGB(0, 255, 0)
                                    End If
                                    
                                    If CLng(initialSerialPriceRange.Value) = 0 Then
                                        initialSerialPriceRange.Value = ajtem.price
                                    End If
                                End If
                            Next ajtem
                        
                            If preSerialPriceRange.Comment Is Nothing Then
                                
                            Else
                                preSerialPriceRange.ClearComments
                            End If
                            
                            preSerialPriceRange.AddComment xtraComment & Chr(10) & txtForPreSerialFromME33K
                            preSerialPriceRange.Comment.Shape.TextFrame.AutoSize = True
                        End If
                    
                    
                    
                    
                        sess.FindById("wnd[0]/tbar[0]/btn[15]").press
                    
                    End If
                    
                Next ir
                 
                '------------------------------------------------------
            Else
                MsgBox "Selection only if first column!"
            End If
        Else
            MsgBox "Works only with Green Light!"
        End If
    Else
        MsgBox "Works only with Green Light!"
    End If
End Sub






Public Sub runMainLogicFor__ME33K(sh1 As Worksheet, st_h As StatusHandler, ByRef xStHelper As Integer)


    Application.Calculation = xlCalculationManual
    
    ' we just taking all the data from this transaction - nothing special!"
    ' im not proud of this "hack"
    ' ----------------------------------------------------
    'Dim x17 As Variant
    'For x17 = 0 To 10
    '    On Error Resume Next
    '    sess.findById("wnd[0]/tbar[0]/btn[12]").press
    'Next x17
    ' ----------------------------------------------------
    
    
    ' this logic provides semi-auto solution when the green light list is done
    '
    ' 1st step: verify if we are in the final touch green light list
    ' 2nd stecheck if selected fields make sense for the logic
    ' 3rd make loop accordingly
    ' 4th format data in reception if necessary
    
    
    
    '1
    ' make verification simple - just check if there is some articles available!
    
    ' w gore - ale nie jest to zawsze najlepsze rozwiazanie - zatem w rem dopoki czego madrego nie wymysle!
    'On Error Resume Next
    'sess.FindById("wnd[0]/tbar[0]/btn[15]").press
    
    Dim ash As Worksheet
    Set ash = ThisWorkbook.ActiveSheet
    
    ' if ash.Range("H1").Value = "Article"
    If ash.Range("A1").Value = "REFERENCE" And ash.Range("AF1").Value = "DOMAIN" Then
        
        Dim selectedRange As Range
        Set selectedRange = Selection
        
        ' 8 stands for H - simplification only for me - if go pro - we need to change this
        If selectedRange.Column = 1 Then
        
            Dim idForMag As String, idForArticle As String
            idForMag = _
                "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]"
            idForArticle = _
                "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[5,24]"

            
            
            
            Dim ir As Range, articleWoIndex As String, domain As String, mag As String
            For Each ir In selectedRange
                
                ' it is a semi-auto loop
                '
                articleWoIndex = CStr(Split(CStr(ir.Value), "-")(0))
                domain = CStr(ash.Range("AF" & CStr(ir.row)))
                mag = domain & "0"
                
                sess.FindById("wnd[0]/tbar[0]/okcd").Text = "ME33K"
                sess.FindById("wnd[0]").sendVKey 0
                sess.FindById("wnd[0]").sendVKey 4
                
                ' tab 20 == possibility to put articles!
                sess.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020").Select
                ' sess.findById("wnd[0]/tbar[1]/btn[8]").press
                
                sess.FindById(idForMag).Text = mag
                ' sess.findById(idForArticle).Text = "98398447XT"
                sess.FindById(idForArticle).Text = articleWoIndex
                
                'sess.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]").SetFocus
                'sess.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]").caretPosition = 4
                sess.FindById("wnd[1]/tbar[0]/btn[0]").press
                
                
                ' choose contract?
                ' Debug.Print TypeName(sess.FindById("wnd[1]/usr/lbl[1,8]"))
                ' .findById("wnd[1]/usr/lbl[37,8]")
                
                Dim x As Variant, y As Variant
                x = 1
                y = 1
                
                
                ' example of output
                
                ' row 8 as first ' v on 21 col
                ' 1, 1 "division" ' v on 21 col
                ' 2, 1"mag" ' v from mag starts like everything else on 21 col
                ' Catég. doc. d'achat x 3 y 1 'K x 3 y 21
                ' Type document achat x 4 y 1 ' YPRS x 4 y 21
                
                ' x = 5 empty
                
                ' beg of table
                ' Doc achat x 6 y 1
                ' Poste x 6 y 12
                ' Article x 6 y 18
                ' Fourn. x 6 y 37
                
                ' x = 7 empty
                
                ' x = 8 content starting
                '
                ' 3939452822 x 8 y 1 ( doc achat )
                ' 00010 x 8 y 12 ( poste )
                ' 9830943980 x 8 y 18 ( article )
                ' 50871U  01 x 8 y 37 ( cofor)
                
                
                Dim innerColl As Collection
                Set innerColl = Nothing
                Set innerColl = New Collection
                Dim tmpStr As String, prevStr As String
                tmpStr = ""
                prevStr = ""
                
                
                For x = 1 To 20
                    
                    tmpStr = ""
                    prevStr = ""
                

                    
                    For y = 1 To 100
                    
                        If y = 1 Then prevStr = ""
                        
                    
                        Set lt = Nothing
                        On Error Resume Next
                        Set lt = sess.FindById("wnd[1]/usr/lbl[" & CStr(y) & "," & CStr(x) & "]")
                        If Not lt Is Nothing Then
                            ' Debug.Print lt.Text & " x " & x & " y " & y
                            
                            tmpStr = tmpStr & lt.Text & ";"
                            prevStr = lt.Text
                            
                            
                        End If
                    Next y
                    
                    innerColl.Add tmpStr
                    

                Next x
                
                
                
                ' form show
                ' then choose line
                
                Dim directDataRow As Variant, el As Variant
                
                PoorForm.ListBox1.Clear
                
                For Each el In innerColl
                    PoorForm.ListBox1.addItem el
                Next el
                
                PoorForm.show
                Debug.Print "PoorForm.TextBoxCaretPos.Value: " & CStr(PoorForm.TextBoxPos.Value)
                directDataRow = PoorForm.TextBoxPos.Value
                sess.FindById("wnd[1]/usr/lbl[1," & CStr(directDataRow) & "]").SetFocus
                sess.FindById("wnd[1]/usr/lbl[1," & CStr(directDataRow) & "]").caretPosition = 1
                

                
                sess.FindById("wnd[1]/tbar[0]/btn[0]").press
                'sess.findById("wnd[0]").sendVKey 0

                sess.FindById("wnd[0]/tbar[1]/btn[5]").press
                
                
                ' ISapCTextField
                ' Debug.Print TypeName(sess.findByid("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-BPRME[9,1]"))
                
                ' session.findById("wnd[0]/usr/ctxtRM06E-EVRTN")
                Debug.Print TypeName(sess.FindById("wnd[0]/usr/ctxtRM06E-EVRTN"))
                Debug.Print TypeName(sess.FindById("wnd[0]/usr/ctxtEKKO-LIFNR"))
                    
                Set tf = Nothing
                    
                On Error Resume Next
                Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-BPRME[9,1]")
                
                Debug.Print tf.Text
                Debug.Print tf.id
                
                
                
                ' very first field:  contract
                Set tf = Nothing
                Set tf = sess.FindById("wnd[0]/usr/ctxtRM06E-EVRTN")
                Debug.Print tf.Text
                Debug.Print tf.id
                
                
                
                ' fournisseur
                Set tf = Nothing
                Set tf = sess.FindById("wnd[0]/usr/ctxtEKKO-LIFNR")
                Debug.Print tf.Text
                Debug.Print tf.id
                
                
                
                ' date contr
                Set tf = Nothing
                Set tf = sess.FindById("wnd[0]/usr/ctxtRM06E-VEDAT")
                Debug.Print tf.Text
                Debug.Print tf.id
                
                
                
                ' type de contrat
                Set tf = Nothing
                Set tf = sess.FindById("wnd[0]/usr/ctxtRM06E-EVART")
                Debug.Print tf.Text
                Debug.Print tf.id
                
                
                ' devise / currency
                Set tf = Nothing
                Set tf = sess.FindById("wnd[0]/usr/ctxtEKKO-WAERS")
                Debug.Print tf.Text
                Debug.Print tf.id
                
                
                
                ' 3 or 4 column jako artykul zero bez indexu
                Set tf = Nothing
                Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,0]")
                Debug.Print tf.Text
                Debug.Print tf.id
                
                
                ' 3 or 4 column jako artykul - jeden nizej z indexem
                Set tf = Nothing
                Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,1]")
                Debug.Print tf.Text
                Debug.Print tf.id
                
                ' jesli tylko dwa, to to powinno byc puste!
                Set tf = Nothing
                Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,2]")
                Debug.Print tf.Text & " tutaj powinno byc pusto!" ' al jest cos takiego: "__________________"
                Debug.Print tf.id
                
                Dim underscores As String
                underscores = "__________________"
                
                
                Dim ajtem As ME33K_Item
                Dim collectionOfItems As Collection
                Set collectionOfItems = Nothing
                Set collectionOfItems = New Collection
                
                For x = 0 To 12
                
                
                    Set ajtem = New ME33K_Item
            
                    Set tf = Nothing
                    y = 3
                    On Error Resume Next
                    Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[" & CStr(y) & "," & CStr(x) & "]")
                        
                    If Not tf Is Nothing Then
                        If tf.Text = "__________________" Then
                            Exit For
                        Else
                        
                            ajtem.article = tf.Text
                            ' session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-NETPR[7,0]")
                            
                            
                            ' EUR
                            ' ====================================================
                            Set tf = Nothing
                            Set tf = sess.FindById("wnd[0]/usr/ctxtEKKO-WAERS")
                            On Error Resume Next
                            ajtem.curr = tf.Text
                            ' ====================================================
                            
                            
                            
                            ' PRICE
                            ' ====================================================
                            ' ====================================================
                            Set tf = Nothing
                            y = 7
                            On Error Resume Next
                            Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-NETPR[" & CStr(y) & "," & CStr(x) & "]")
                            
                            If Not tf Is Nothing Then
                                ajtem.strPrice = tf.Text
                                ajtem.parsePrice
                                
                                
                                Set tf = Nothing
                                y = 9
                                ' On Error Resume Next
                                Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-BPRME[" & CStr(y) & "," & CStr(x) & "]")
                                If Not tf Is Nothing Then ajtem.up = tf.Text
                                
                                ajtem.makePriceByUnit
                                
                                
                                collectionOfItems.Add ajtem
                            End If
                            ' ====================================================
                            ' ====================================================
                            
                        End If
                    End If
                Next x
                
                
                
                
                ' ISapTableControlTarget
                ' Debug.Print TypeName(sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220"))
                
                'Set table2 = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220")
                '
                '
                'x = 3
                'y = 1
                'Set item2 = table2.GetCell(x, y)
                '
                'Debug.Print item2.id & " " & TypeName(item2)
                '
                'Set tf = item2
                'Debug.Print item2.Text
                
                
                'For Each ajtem In collectionOfItems
                '    Debug.Print "ajtem: " & ajtem.article & " " & ajtem.price
                'Next ajtem
                
                
                
                ' ir jako column A
                ' final green light re-check price
                
                Dim e1 As E_GREEN_LIGHT, e2 As E_GREEN_LIGHT
                e1 = E_GREEN_LIGHT_PRE_SERIAL_PRICE_YPRS_contract
                e2 = E_GREEN_LIGHT_INITIAL_SERIAL_PRICE
                
                Dim preSerialPriceRange As Range, initialSerialPriceRange As Range

                Set preSerialPriceRange = ir.offset(0, e1 - 1)
                
                preSerialPriceRange.Font.Bold = True
                preSerialPriceRange.Font.Color = RGB(0, 0, 0)
                preSerialPriceRange.Interior.Color = RGB(255, 255, 255)
                
                Set initialSerialPriceRange = ir.offset(0, e2 - 1)
                
                initialSerialPriceRange.Font.Bold = True
                initialSerialPriceRange.Font.Color = RGB(0, 0, 0)
                initialSerialPriceRange.Interior.Color = RGB(255, 255, 255)
                
                Debug.Print "pre: " & preSerialPriceRange.Address & " " & preSerialPriceRange.Value
                Debug.Print "ini: " & initialSerialPriceRange.Address & " " & initialSerialPriceRange.Value
                
                
                If collectionOfItems.count > 0 Then
                

                    
                    Dim txtForPreSerialFromME33K As String, xtraComment As String
                    
                    txtForPreSerialFromME33K = ""
                    xtraComment = ""
                    
                    txtForPreSerialFromME33K = txtForPreSerialFromME33K & CStr(Date) & Chr(10) & _
                        "data from ME33K: " & Chr(10)

                    For Each ajtem In collectionOfItems
                        txtForPreSerialFromME33K = txtForPreSerialFromME33K & Chr(10) & _
                            ": " & ajtem.article & " " & ajtem.price
                            
                            
                        If CStr(ajtem.article) = CStr(ir.Value) Then
                            ' w index
                            If (1# * (preSerialPriceRange.Value / ajtem.price)) > 1.1 Then
                                ' out of margin
                                ' MsgBox "price on me33k is diff for " & CStr(ir.Value) & "!", vbCritical
                                
                                xtraComment = Chr(10) & "pre-serial price issue: " & Chr(10) & _
                                    "me33k pre-serial price: " & ajtem.price & Chr(10) & _
                                    "tp04: " & CStr(preSerialPriceRange.Value)
                                    
                                preSerialPriceRange.Font.Color = RGB(255, 0, 0)
                            Else
                                preSerialPriceRange.Font.Color = RGB(0, 255, 0)
                            End If
                        End If
                        
                        If CStr(ajtem.article) = CStr(Split(ir.Value, "-")(0)) Then
                            ' wo index
                            If (1# * (initialSerialPriceRange.Value / ajtem.price)) > 1.1 Then
                                ' out of margin
                                ' MsgBox "price on me33k is diff for " & CStr(ir.Value) & "!", vbCritical
                                
                                xtraComment = Chr(10) & "initial price issue: " & Chr(10) & _
                                    "me33k initial price: " & ajtem.price & Chr(10) & _
                                    "tp04 initial price: " & CStr(initialSerialPriceRange.Value)
                                    
                                initialSerialPriceRange.Font.Color = RGB(255, 0, 0)
                            Else
                                initialSerialPriceRange.Font.Color = RGB(0, 255, 0)
                            End If
                        End If
                    Next ajtem
                
                    If preSerialPriceRange.Comment Is Nothing Then
                        
                    Else
                        preSerialPriceRange.ClearComments
                    End If
                    
                    preSerialPriceRange.AddComment xtraComment & Chr(10) & txtForPreSerialFromME33K
                    preSerialPriceRange.Comment.Shape.TextFrame.AutoSize = True
                End If
                
                
                
                
                sess.FindById("wnd[0]/tbar[0]/btn[15]").press
                
                
            Next ir
        End If
    Else
        MsgBox "You need to be in active green light report with selected articles for potential conversion!", vbInformation
    End If
    
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub runMainLogicFor__Y_PI1_80000391(sh1 As Worksheet, ByRef st_h As StatusHandler, ByRef xStHelper As Integer)


    Application.Calculation = xlCalculationManual
    
    ' we just taking all the data from this transaction - nothing special!"
    ' im not proud of this "hack"
    ' ----------------------------------------------------
    Dim x17 As Variant
    For x17 = 0 To 10
        On Error Resume Next
        sess.FindById("wnd[0]/tbar[0]/btn[12]").press
    Next x17
    ' ----------------------------------------------------
    
    
    sess.FindById("wnd[0]/tbar[0]/okcd").Text = "Y_PI1_80000391"
    sess.FindById("wnd[0]").sendVKey 0
    sess.FindById("wnd[0]/tbar[1]/btn[8]").press
    
    'sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetCurrentCell 53, "LFA1-NAME1"
    'sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").FirstVisibleRow = 31
    'sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectedRows = "53"
    
    Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    Set cols = grid.ColumnOrder
    
    
    Dim refOutput2 As Range  ' , lm As Long
    Set refOutput2 = sh1.Cells(2, 1)
    
    ' lm = CLng((ish.Range("A1").End(xlDown).row) / 50)
    
    For x = 0 To grid.RowCount - 1
        For y = 0 To grid.ColumnCount - 1
        
        
            refOutput2.offset(0, y).Value = CStr(grid.getCellValue(x, cols(y)))
        Next y
        
        If x Mod 50 = 0 Then
        
            grid.FirstVisibleRow = x
            grid.CurrentCellRow = x
            
            ' delegacjaDlaProgresu st_h, xStHelper, 20 ' obsolete
            ' st_h.progress_increase
            
        End If
        
        
        Set refOutput2 = refOutput2.offset(1, 0)
    Next x
    
    Application.Calculation = xlCalculationAutomatic
End Sub


Public Sub runMainLogicForSQ01__with_preDef(mcfg1 As E_SQ01_CONFIG, ByRef ish As Worksheet, ByRef osh As Worksheet, ByRef st_h As StatusHandler, ByRef xStHelper As Integer, whichLineForInput As Integer)


    ' from now on 2 possibilities
    
    Application.Calculation = xlCalculationManual


    If mcfg1 = E_SQ01_CONFIG_START_FROM_PASTING_ALREADY_PREPARED_INPUT_LIST Then
        
        ' avoiding initial stuff
    Else
    
        ' no need to clear anything becuase we just created new worksheet
        ' clearProduitFromInput ish = obsolete
        
        ' im not proud of this "hack"
        ' ----------------------------------------------------
        Dim x17 As Variant
        For x17 = 0 To 10
            On Error Resume Next
            sess.FindById("wnd[0]/tbar[0]/btn[12]").press
        Next x17
        ' ----------------------------------------------------
    
    
        sess.FindById("wnd[0]/tbar[0]/okcd").Text = "SQ01"
        sess.FindById("wnd[0]").sendVKey 0
        sess.FindById("wnd[0]/mbar/menu[1]/menu[7]").Select
        
        'sess.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").SelectedRows = "0"
        'sess.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").DoubleClickCurrentCell
        
        
        runLoopsWithPreDef "wnd[1]/usr/cntlGRID1/shellcont/shell", E_SQ01_PREDEF_SYSTEM, whichLineForInput
        
        runLoopsWithPreDef "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", E_SQ01_PREDEF_LIST1, whichLineForInput
        
        
        sess.FindById("wnd[0]/tbar[1]/btn[17]").press
        sess.FindById("wnd[1]").sendVKey 4
        
        runLoopsWithPreDef "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell", E_SQ01_PREDEF_STD1, whichLineForInput
        
        
        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        ' wynik !
        ' z tego pobieramy liste czesci!
        sess.FindById("wnd[0]/tbar[1]/btn[8]").press
        
        
        ' teraz glowny wrzut na input
        ' =========================================================================
        
        Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        Set cols = grid.ColumnOrder
        
        
        Dim refInput As Range
        Set refInput = ish.Range("A1")
        
        refInput.offset(0, 0).Value = "PRODUIT"
        refInput.offset(0, 1).Value = "DESIGNATION"
        refInput.offset(0, 2).Value = "RU"
        
        Set refInput = refInput.offset(1, 0)
        
        
        
        ' MAIN LOOP FOR FIRST LIST OF PN WITHOUT INDICE
        For x = 0 To grid.RowCount
            For y = 0 To grid.ColumnCount
            
            
                If CStr(cols(y)) Like "*_PRODUIT*" Then
                
                    ' Debug.Print Cstr(grid.GetCellValue(x2, cols(y2))) ' OK
                    refInput.Value = CStr(grid.getCellValue(x, cols(y)))
                    refInput.offset(0, 1).Value = CStr(grid.getCellValue(x, cols(y + 1)))
                    refInput.offset(0, 2).Value = CStr(grid.getCellValue(x, cols(y + 2)))
                    Exit For
                End If
                
            Next y
            
            
            

            If x Mod 50 = 0 Then
            
                grid.FirstVisibleRow = x
                grid.CurrentCellRow = x
                
                
                ' powazny side effect ...
                delegacjaDlaProgresu st_h, xStHelper, 20
                
            End If
            
            
            Set refInput = refInput.offset(1, 0)
        Next x
        
        makeDoubleListWithDashAndStar ish.Range("A2")
        
        st_h.hide
        Set st_h = Nothing
        Set st_h = New StatusHandler
        st_h.init_statusbar CLng((ish.Range("A1").End(xlDown).row) / 50)
        st_h.show
        st_h.progress_increase
        
        ' go back! x2
        sess.FindById("wnd[0]/tbar[0]/btn[15]").press
        sess.FindById("wnd[0]/tbar[0]/btn[15]").press
        
        
        ' teraz etap 2
        ' mamy powiekszona liste czesci - musimy przejsc jeszcze raz przez wybor
        
        ' tym razem wybieramy konkretnie
        
        runLoopsWithPreDef "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", E_SQ01_PREDEF_LIST2, whichLineForInput
        
        
        sess.FindById("wnd[0]/tbar[1]/btn[17]").press
        ' otwieramy liste z dostepnymi standardami do etapu drugiego!
        sess.FindById("wnd[1]").sendVKey 4
        
        sess.FindById("wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "0"
        sess.FindById("wnd[2]/tbar[0]/btn[2]").press
        
        ' runLoopsWithPreDef "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell", E_SQ01_PREDEF_STD2, whichLineForInput
        session.FindById("wnd[1]/usr/ctxtRS38R-VARIANT").Text = getProperRefFromPreDef(E_SQ01_PREDEF_STD2, whichLineForInput)
        
        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        
        ' teraz powinnismy byc na oknie ponownego forma wejsciowego dla listy quasi tp04
        
        
    End If
    
    Dim rngToCopy As Range
    Dim refOutput2 As Range  ' , lm As Long
    Set refOutput2 = osh.Cells(2, 1)
    
    ' A1048576 - always last one
    Set rngToCopy = tryToAssignFirstBatchOfData(ish)
    
    Do
    
    
        With sess
        
        
            ' Debug.Print TypeName(.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH"))
            
            
            Dim btn3 As Variant ' SAPFEWSELib.GuiButton
            Set btn3 = findMyButton() ' .FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH")
            ' simple solution - you need to take 3rd button from top
            'Debug.Print btn.Top
            
            'Set btn = .FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH")
            'Debug.Print btn.Top
            '
            'Set btn = .FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH")
            'Debug.Print btn.Top
            '
            'Set btn = .FindById("wnd[0]/usr/btn%_SP$00004_%_APP_%-VALU_PUSH")
            'Debug.Print btn.Top
            
    
            
            ' Code de suppression dans le
            '.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").Press
            ' please do not delete data from this field!
            ' .FindById("wnd[1]/tbar[0]/btn[16]").Press
            '.FindById("wnd[1]/tbar[0]/btn[0]").Press
            '.FindById("wnd[1]/tbar[0]/btn[8]").Press
        
            '.FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").Press
            '.FindById("wnd[1]/tbar[0]/btn[16]").Press
            '.FindById("wnd[1]/tbar[0]/btn[0]").Press
            '.FindById("wnd[1]/tbar[0]/btn[8]").Press
            
            ' Numéro d'article
            '.FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").Press
            '.FindById("wnd[1]/tbar[0]/btn[16]").Press
            
            btn3.press
            .FindById("wnd[1]/tbar[0]/btn[16]").press
               
            ' now we making it in the loop - to avoid OLE data missing issue
            ' Set rngToCopy = ish.Range(ish.Range("A2"), ish.Range("A2").End(xlDown))
            
            ' Debug.Print rngToCopy.Address
            rngToCopy.Copy
            
            .FindById("wnd[1]/tbar[0]/btn[24]").press
            .FindById("wnd[1]/tbar[0]/btn[8]").press
            .FindById("wnd[0]/tbar[1]/btn[8]").press
        End With
    
        
        Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        Set cols = grid.ColumnOrder
        
        
        
        ' refOutput2
        
        If refOutput2.row > 2 Then
            ' so at least next iteration
            
            ' Debug.Print refOutput2.row
            
            
            ' very simple logic
            ' be careful - almost without any additional checking
            If Trim(refOutput2.Value) = "" Then
                Set refOutput2 = refOutput2.offset(-1, 0)
            End If
            
            
        End If

        
        ' lm = CLng((ish.Range("A1").End(xlDown).row) / 50)
        
        For x = 0 To grid.RowCount - 1
            For y = 0 To grid.ColumnCount - 1
            
            
                refOutput2.offset(0, y).Value = CStr(grid.getCellValue(x, cols(y)))
            Next y
            
            If x Mod 50 = 0 Then
            
                grid.FirstVisibleRow = x
                grid.CurrentCellRow = x
                
                ' delegacjaDlaProgresu st_h, xStHelper, 20 ' obsolete
                st_h.progress_increase
                
            End If
            
            
            Set refOutput2 = refOutput2.offset(1, 0)
        Next x
        
        
        'step back in SIGAPP
        sess.FindById("wnd[0]/tbar[0]/btn[3]").press
        ' new range to copy
        Set rngToCopy = tryToAssignNewRange(ish, rngToCopy)
        
        
        
        
        
    Loop Until youDryOutRangeOfDataForSq01(rngToCopy)
    
    st_h.hide
    
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Private Function tryToAssignFirstBatchOfData(sh As Worksheet) As Range
    Set tryToAssignFirstBatchOfData = sh.Range("A2:A500")
End Function



Private Function tryToAssignNewRange(sh As Worksheet, data As Range) As Range
    Set tryToAssignNewRange = Nothing
    
    
    Dim tmp As Range
    
    Set tmp = data.offset(499, 0)
    
    ' Debug.Print tmp.Address
    
    Set tryToAssignNewRange = tmp
End Function

Private Function youDryOutRangeOfDataForSq01(data As Range) As Boolean

    youDryOutRangeOfDataForSq01 = False
    
    Dim ir As Range, i As Integer, tmp As Range
    
    i = 0
    For Each ir In data
        
        If Trim(ir.Value) <> "" Then
            i = i + 1
        End If
    Next ir
    
    If i > 0 Then
        youDryOutRangeOfDataForSq01 = False
    Else
        youDryOutRangeOfDataForSq01 = True
    End If
End Function

Private Function findMyButton() As Variant ' SAPFEWSELib.GuiButton
    
    Set findMyButton = Nothing
    
    ' top
    '51
    '73
    '95
    '29
    
    Dim tmpButton As Variant ' SAPFEWSELib.GuiButton
    
    Dim x As Variant
    For x = 1 To 4
        ' Set btn = .FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH")
        ' "wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH"
        
        Set tmpButton = Nothing
        On Error Resume Next
        Set tmpButton = sess.FindById("wnd[0]/usr/btn%_SP$0000" & CStr(x) & "_%_APP_%-VALU_PUSH")
        'On Error Resume Next
        'Set tmpButton = sess.FindById("wnd[1]/usr/btn%_SP$0000" & CStr(x) & "_%_APP_%-VALU_PUSH")
        
        If Not tmpButton Is Nothing Then
        
            Debug.Print "lookin for reference button : tmpButton.Top: " & tmpButton.Top
            
            If (tmpButton.Top > 70 And tmpButton.Top < 80) Or (tmpButton.Top = 73) Then
                ' this is it
                Set findMyButton = tmpButton
                Exit For
            End If
        End If
    Next x
End Function



Public Sub runMainLogicForSQ01__forMassItem(ByRef r As Range, ByRef ish As Worksheet, osh As Worksheet, st_h As StatusHandler, ByRef xStHelper As Integer, whichLineForInput As Integer)


    ' im not proud of this "hack"
    ' ----------------------------------------------------
    Dim x17 As Variant
    For x17 = 0 To 10
        On Error Resume Next
        sess.FindById("wnd[0]/tbar[0]/btn[12]").press
    Next x17
    ' ----------------------------------------------------


    sess.FindById("wnd[0]/tbar[0]/okcd").Text = "SQ01"
    sess.FindById("wnd[0]").sendVKey 0
    sess.FindById("wnd[0]/mbar/menu[1]/menu[7]").Select
    
    'sess.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").SelectedRows = "0"
    'sess.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").DoubleClickCurrentCell
    
    
    runLoopsForceString "wnd[1]/usr/cntlGRID1/shellcont/shell", _
        CStr(ThisWorkbook.Sheets("register").Range("A50").Value)
    runLoopsForceString "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", _
        CStr(ThisWorkbook.Sheets("register").Range("B50").Value)
    
    
    sess.FindById("wnd[0]/tbar[1]/btn[17]").press
    sess.FindById("wnd[1]").sendVKey 4
    
    
    Dim stringPattern As String
    stringPattern = CStr(ThisWorkbook.Sheets("register").Range("C50").Value)
    stringPattern = Replace(stringPattern, "XXX", CStr(r.Value))
    
    runLoopsForceString "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell", _
        stringPattern

    
    
    sess.FindById("wnd[1]/tbar[0]/btn[0]").press
    
    ' wynik !
    ' z tego pobieramy liste czesci!
    sess.FindById("wnd[0]/tbar[1]/btn[8]").press
    
    
    ' teraz glowny wrzut na input
    ' =========================================================================
    
    Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    Set cols = grid.ColumnOrder
    
    
    Dim refInput As Range
    Set refInput = ish.Range("A1")
    
    refInput.offset(0, 0).Value = "PRODUIT"
    refInput.offset(0, 1).Value = "DESIGNATION"
    refInput.offset(0, 2).Value = "RU"
    
    Set refInput = refInput.offset(1, 0)
    
    
    ' MAIN LOOP FOR FIRST LIST OF PN WITHOUT INDICE
    For x = 0 To grid.RowCount
        For y = 0 To grid.ColumnCount
        
        
            If CStr(cols(y)) Like "*_PRODUIT*" Then
            
                ' Debug.Print Cstr(grid.GetCellValue(x2, cols(y2))) ' OK
                refInput.Value = CStr(grid.getCellValue(x, cols(y)))
                refInput.offset(0, 1).Value = CStr(grid.getCellValue(x, cols(y + 1)))
                refInput.offset(0, 2).Value = CStr(grid.getCellValue(x, cols(y + 2)))
                Exit For
            End If
            
        Next y
        
        
        

        If x Mod 50 = 0 Then
        
            grid.FirstVisibleRow = x
            grid.CurrentCellRow = x
            
            
            ' powazny side effect ...
            delegacjaDlaProgresu st_h, xStHelper, 20
            
        End If
        
        
        Set refInput = refInput.offset(1, 0)
    Next x
    
    makeDoubleListWithDashAndStar ish.Range("A2")
    
    st_h.hide
    Set st_h = Nothing
    Set st_h = New StatusHandler
    st_h.init_statusbar CLng((ish.Range("A1").End(xlDown).row) / 50)
    st_h.show
    st_h.progress_increase
    
    
    ' go back! x2
    sess.FindById("wnd[0]/tbar[0]/btn[15]").press
    sess.FindById("wnd[0]/tbar[0]/btn[15]").press
    
    
    ' teraz etap 2
    ' mamy powiekszona liste czesci - musimy przejsc jeszcze raz przez wybor
    
    ' tym razem wybieramy konkretnie
    stringPattern = CStr(ThisWorkbook.Sheets("register").Range("D50").Value)
    stringPattern = Replace(stringPattern, "XXX", CStr(r.Value))
    
    runLoopsForceString "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", _
        stringPattern
    
    
    sess.FindById("wnd[0]/tbar[1]/btn[17]").press
    ' otwieramy liste z dostepnymi standardami do etapu drugiego!
    sess.FindById("wnd[1]").sendVKey 4
    
    
    
    stringPattern = CStr(ThisWorkbook.Sheets("register").Range("E50").Value)
    stringPattern = Replace(stringPattern, "XXX", CStr(r.Value))
    
    runLoopsForceString "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell", stringPattern
    
    sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        
        
        
    Dim rngToCopy As Range
    Dim refOutput2 As Range  ' , lm As Long
    Set refOutput2 = osh.Cells(2, 1)
    
    ' A1048576 - always last one
    Set rngToCopy = tryToAssignFirstBatchOfData(ish)
    
    Do
    
    
        With sess
        
        
            ' Debug.Print TypeName(.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH"))
            
            
            Dim btn3 As Variant ' SAPFEWSELib.GuiButton
            Set btn3 = findMyButton() ' .FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH")
            ' simple solution - you need to take 3rd button from top
            'Debug.Print btn.Top
            
            'Set btn = .FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH")
            'Debug.Print btn.Top
            '
            'Set btn = .FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH")
            'Debug.Print btn.Top
            '
            'Set btn = .FindById("wnd[0]/usr/btn%_SP$00004_%_APP_%-VALU_PUSH")
            'Debug.Print btn.Top
            
    
            
            ' Code de suppression dans le
            '.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").Press
            ' please do not delete data from this field!
            ' .FindById("wnd[1]/tbar[0]/btn[16]").Press
            '.FindById("wnd[1]/tbar[0]/btn[0]").Press
            '.FindById("wnd[1]/tbar[0]/btn[8]").Press
        
            '.FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").Press
            '.FindById("wnd[1]/tbar[0]/btn[16]").Press
            '.FindById("wnd[1]/tbar[0]/btn[0]").Press
            '.FindById("wnd[1]/tbar[0]/btn[8]").Press
            
            ' Numéro d'article
            '.FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").Press
            '.FindById("wnd[1]/tbar[0]/btn[16]").Press
            
            btn3.press
            .FindById("wnd[1]/tbar[0]/btn[16]").press
               
            ' now we making it in the loop - to avoid OLE data missing issue
            ' Set rngToCopy = ish.Range(ish.Range("A2"), ish.Range("A2").End(xlDown))
            
            ' Debug.Print rngToCopy.Address
            rngToCopy.Copy
            
            .FindById("wnd[1]/tbar[0]/btn[24]").press
            .FindById("wnd[1]/tbar[0]/btn[8]").press
            .FindById("wnd[0]/tbar[1]/btn[8]").press
        End With
    
        
        Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        Set cols = grid.ColumnOrder
        
        
        
        ' refOutput2
        
        If refOutput2.row > 2 Then
            ' so at least next iteration
            
            ' Debug.Print refOutput2.row
            
            
            ' very simple logic
            ' be careful - almost without any additional checking
            If Trim(refOutput2.Value) = "" Then
                Set refOutput2 = refOutput2.offset(-1, 0)
            End If
            
            
        End If

        
        ' lm = CLng((ish.Range("A1").End(xlDown).row) / 50)
        
        For x = 0 To grid.RowCount - 1
            For y = 0 To grid.ColumnCount - 1
            
            
                refOutput2.offset(0, y).Value = CStr(grid.getCellValue(x, cols(y)))
            Next y
            
            If x Mod 50 = 0 Then
            
                grid.FirstVisibleRow = x
                grid.CurrentCellRow = x
                
                ' delegacjaDlaProgresu st_h, xStHelper, 20 ' obsolete
                st_h.progress_increase
                
            End If
            
            
            Set refOutput2 = refOutput2.offset(1, 0)
        Next x
        
        
        'step back in SIGAPP
        sess.FindById("wnd[0]/tbar[0]/btn[3]").press
        ' new range to copy
        Set rngToCopy = tryToAssignNewRange(ish, rngToCopy)
        
        
        
        
        
    Loop Until youDryOutRangeOfDataForSq01(rngToCopy)
    
    st_h.hide
    

        

End Sub




Public Sub runMainLogicForSQ01__forMassItem_WO_IN_LIST(ByRef r As Range, osh As Worksheet, st_h As StatusHandler, ByRef xStHelper As Integer, whichLineForInput As Integer)


    EVO.KillMessageFilter
    
    
    
    ' im not proud of this "hack"
    ' ----------------------------------------------------
    Dim x17 As Variant
    For x17 = 0 To 10
        On Error Resume Next
        sess.FindById("wnd[0]/tbar[0]/btn[12]").press
    Next x17
    ' ----------------------------------------------------


    sess.FindById("wnd[0]/tbar[0]/okcd").Text = "SQ01"
    sess.FindById("wnd[0]").sendVKey 0
    
    
    Dim stringPattern As String

    
    ' teraz etap 2
    ' mamy powiekszona liste czesci - musimy przejsc jeszcze raz przez wybor
    
    ' tym razem wybieramy konkretnie
    stringPattern = CStr(ThisWorkbook.Sheets("register").Range("D50").Value)
    stringPattern = Replace(stringPattern, "XXX", CStr(r.Value))
    
    runLoopsForceString "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", _
        stringPattern
    
    
    sess.FindById("wnd[0]/tbar[1]/btn[17]").press
    ' otwieramy liste z dostepnymi standardami do etapu drugiego!
    sess.FindById("wnd[1]").sendVKey 4
    
    
    
    stringPattern = CStr(ThisWorkbook.Sheets("register").Range("E50").Value)
    stringPattern = Replace(stringPattern, "XXX", CStr(r.Value))
    
    runLoopsForceString "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell", stringPattern
    
    sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        
        
        
    Dim rngToCopy As Range
    Dim refOutput2 As Range  ' , lm As Long
    Set refOutput2 = osh.Cells(2, 1)
    
    ' A1048576 - always last one
    'Set rngToCopy = tryToAssignFirstBatchOfData(ish)
    

    
    With sess
        
        
        ' Debug.Print TypeName(.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH"))
        
        
        Dim btn3 As Variant ' SAPFEWSELib.GuiButton
        Set btn3 = findMyButton() ' .FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH")
        ' simple solution - you need to take 3rd button from top
        'Debug.Print btn.Top
        
        'Set btn = .FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH")
        'Debug.Print btn.Top
        '
        'Set btn = .FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH")
        'Debug.Print btn.Top
        '
        'Set btn = .FindById("wnd[0]/usr/btn%_SP$00004_%_APP_%-VALU_PUSH")
        'Debug.Print btn.Top
        

        
        ' Code de suppression dans le
        '.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").Press
        ' please do not delete data from this field!
        ' .FindById("wnd[1]/tbar[0]/btn[16]").Press
        '.FindById("wnd[1]/tbar[0]/btn[0]").Press
        '.FindById("wnd[1]/tbar[0]/btn[8]").Press
    
        '.FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").Press
        '.FindById("wnd[1]/tbar[0]/btn[16]").Press
        '.FindById("wnd[1]/tbar[0]/btn[0]").Press
        '.FindById("wnd[1]/tbar[0]/btn[8]").Press
        
        ' Numéro d'article
        '.FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").Press
        '.FindById("wnd[1]/tbar[0]/btn[16]").Press
        
        btn3.press
        .FindById("wnd[1]/tbar[0]/btn[16]").press
           
        ' now we making it in the loop - to avoid OLE data missing issue
        ' Set rngToCopy = ish.Range(ish.Range("A2"), ish.Range("A2").End(xlDown))
        
        ' Debug.Print rngToCopy.Address
        'rngToCopy.Copy
        
        '.FindById("wnd[1]/tbar[0]/btn[24]").press
        .FindById("wnd[1]/tbar[0]/btn[8]").press
        .FindById("wnd[0]/tbar[1]/btn[8]").press
    End With
    
    
    
    
    DoEvents
        
    Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    Set cols = grid.ColumnOrder
        
        
        


    
    ' lm = CLng((ish.Range("A1").End(xlDown).row) / 50)
    
    Debug.Print grid.RowCount
    
    Dim x As Long, y As Long
        
    For x = 0 To grid.RowCount - 1
    
        ' Debug.Print grid.RowCount
         
        ' Debug.Assert (x - 2000) <> grid.RowCount ' NOK
        
        'If x > (grid.RowCount - 1) Then
        '    Exit For
        'End If
    
        
        If Trim(grid.getCellValue(x, "DOMAINE")) = "" Then
            
            ' nop
            ' Exit For
        Else
        
            ' y = 0 for domain - only if domain is not empty to put in excel!
    
            For y = 0 To grid.ColumnCount - 1
            
            
                
            
                refOutput2.offset(0, y).Value = CStr(grid.getCellValue(x, cols(y)))
            Next y
            
            Set refOutput2 = refOutput2.offset(1, 0)
        End If
        
        
        If x Mod 50 = 0 Then
        
            grid.FirstVisibleRow = x
            grid.CurrentCellRow = x
            
            ' delegacjaDlaProgresu st_h, xStHelper, 20 ' obsolete
            ' st_h.progress_increase
            
            
            
        End If
        
        
        If x Mod 100 = 0 Then
            Me.justTouch2 ' not even clicking
        End If
        
        If x Mod 1000 = 0 Then
        
            ' quik re-define every 1k rows
            
            Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            ' Set cols = grid.ColumnOrder
            
            'On Error GoTo loopOut1
            'grid.FirstVisibleRow = x
            
            'On Error GoTo loopOut1
            'grid.CurrentCellRow = x
        End If
        
        
    Next x
    

loopOut1:
    
        
    sess.FindById("wnd[0]/tbar[0]/btn[3]").press
    DoEvents
    sess.FindById("wnd[0]/tbar[0]/btn[3]").press
    
    st_h.hide
    
    EVO.SuppressingMessageModule.RestoreMessageFilter
        

End Sub




Public Sub runMainLogicForSQ01__forMassItem_WO_IN_LIST_2(ByRef r As Range, osh As Worksheet, st_h As StatusHandler, ByRef xStHelper As Integer, whichLineForInput As Integer)


    EVO.KillMessageFilter
    
    
    
    ' im not proud of this "hack"
    ' ----------------------------------------------------
    Dim x17 As Variant
    For x17 = 0 To 10
        On Error Resume Next
        sess.FindById("wnd[0]/tbar[0]/btn[12]").press
    Next x17
    ' ----------------------------------------------------


    sess.FindById("wnd[0]/tbar[0]/okcd").Text = "SQ01"
    sess.FindById("wnd[0]").sendVKey 0
    
    
    Dim stringPattern As String

    
    ' teraz etap 2
    ' mamy powiekszona liste czesci - musimy przejsc jeszcze raz przez wybor
    
    ' tym razem wybieramy konkretnie
    stringPattern = CStr(ThisWorkbook.Sheets("register").Range("D50").Value)
    stringPattern = Replace(stringPattern, "XXX", CStr(r.Value))
    
    runLoopsForceString "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", _
        stringPattern
    
    
    sess.FindById("wnd[0]/tbar[1]/btn[17]").press
    ' otwieramy liste z dostepnymi standardami do etapu drugiego!
    sess.FindById("wnd[1]").sendVKey 4
    
    
    
    stringPattern = CStr(ThisWorkbook.Sheets("register").Range("E50").Value)
    stringPattern = Replace(stringPattern, "XXX", CStr(r.Value))
    
    runLoopsForceString "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell", stringPattern
    
    sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        
        
        
    Dim rngToCopy As Range
    Dim refOutput2 As Range  ' , lm As Long
    Set refOutput2 = osh.Cells(2, 1)
    
    ' A1048576 - always last one
    'Set rngToCopy = tryToAssignFirstBatchOfData(ish)
    

    
    With sess
        
        
        ' Debug.Print TypeName(.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH"))
        
        
        Dim btn3 As Variant ' SAPFEWSELib.GuiButton
        Set btn3 = findMyButton() ' .FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH")
        ' simple solution - you need to take 3rd button from top
        'Debug.Print btn.Top
        
        'Set btn = .FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH")
        'Debug.Print btn.Top
        '
        'Set btn = .FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH")
        'Debug.Print btn.Top
        '
        'Set btn = .FindById("wnd[0]/usr/btn%_SP$00004_%_APP_%-VALU_PUSH")
        'Debug.Print btn.Top
        

        
        ' Code de suppression dans le
        '.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").Press
        ' please do not delete data from this field!
        ' .FindById("wnd[1]/tbar[0]/btn[16]").Press
        '.FindById("wnd[1]/tbar[0]/btn[0]").Press
        '.FindById("wnd[1]/tbar[0]/btn[8]").Press
    
        '.FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").Press
        '.FindById("wnd[1]/tbar[0]/btn[16]").Press
        '.FindById("wnd[1]/tbar[0]/btn[0]").Press
        '.FindById("wnd[1]/tbar[0]/btn[8]").Press
        
        ' Numéro d'article
        '.FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").Press
        '.FindById("wnd[1]/tbar[0]/btn[16]").Press
        
        btn3.press
        .FindById("wnd[1]/tbar[0]/btn[16]").press
           
        ' now we making it in the loop - to avoid OLE data missing issue
        ' Set rngToCopy = ish.Range(ish.Range("A2"), ish.Range("A2").End(xlDown))
        
        ' Debug.Print rngToCopy.Address
        'rngToCopy.Copy
        
        '.FindById("wnd[1]/tbar[0]/btn[24]").press
        .FindById("wnd[1]/tbar[0]/btn[8]").press
        .FindById("wnd[0]/tbar[1]/btn[8]").press
    End With
    
    
    
    
    DoEvents
        
    Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    Set cols = grid.ColumnOrder
        
        
        


    
    ' lm = CLng((ish.Range("A1").End(xlDown).row) / 50)
    
    Debug.Print grid.RowCount
    
    Dim x As Long, y As Long
        
    For x = 0 To grid.RowCount - 1
    
        ' Debug.Print grid.RowCount
         
        ' Debug.Assert (x - 100) <> grid.RowCount ' NOK
        
        'If x > (grid.RowCount - 1) Then
        '    Exit For
        'End If
    
        

        
        ' y = 0 for domain - only if domain is not empty to put in excel! - not in this case

        For y = 0 To grid.ColumnCount - 1
        
        
            
        
            refOutput2.offset(0, y).Value = CStr(grid.getCellValue(x, cols(y)))
        Next y
        
        Set refOutput2 = refOutput2.offset(1, 0)

        
        
        If x Mod 50 = 0 Then
        
            On Error Resume Next
            grid.FirstVisibleRow = x
            
            On Error Resume Next
            grid.CurrentCellRow = x
            
            ' delegacjaDlaProgresu st_h, xStHelper, 20 ' obsolete
            ' st_h.progress_increase
            
            
            
        End If
        
        
        If x Mod 100 = 0 Then
            Me.justTouch2 ' not even clicking
        End If
        
        If x Mod 1000 = 0 Then
        
            ' quik re-define every 1k rows
            
            Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            Set cols = grid.ColumnOrder
            
            On Error Resume Next
            grid.FirstVisibleRow = x
            
            On Error Resume Next
            grid.CurrentCellRow = x
        End If
        
        
    Next x
    
        
    sess.FindById("wnd[0]/tbar[0]/btn[3]").press
    DoEvents
    sess.FindById("wnd[0]/tbar[0]/btn[3]").press
    
    st_h.hide
    
    EVO.SuppressingMessageModule.RestoreMessageFilter
        

End Sub





Public Sub runMainLogicForSQ01(mcfg1 As E_SQ01_CONFIG, ByRef ish As Worksheet, ByRef osh As Worksheet, ByRef st_h As StatusHandler, ByRef xStHelper As Integer, whichLineForInput As Integer)


    If mcfg1 = E_SQ01_CONFIG_START_FROM_PASTING_ALREADY_PREPARED_INPUT_LIST Then
        
        ' avoiding initial stuff
    Else
    
        ' no need to clear anything becuase we just created new worksheet
        ' clearProduitFromInput ish = obsolete
        
        ' im not proud of this "hack"
        ' ----------------------------------------------------
        Dim x17 As Variant
        For x17 = 0 To 10
            On Error Resume Next
            sess.FindById("wnd[0]/tbar[0]/btn[12]").press
        Next x17
        ' ----------------------------------------------------
    
    
        sess.FindById("wnd[0]/tbar[0]/okcd").Text = "SQ01"
        sess.FindById("wnd[0]").sendVKey 0
        sess.FindById("wnd[0]/mbar/menu[1]/menu[7]").Select
        
        'sess.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").SelectedRows = "0"
        'sess.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").DoubleClickCurrentCell
        
        
        runLoops "wnd[1]/usr/cntlGRID1/shellcont/shell"
        runLoops "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell"
        
        
        sess.FindById("wnd[0]/tbar[1]/btn[17]").press
        sess.FindById("wnd[1]").sendVKey 4
        
        runLoops "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell"
        
        
        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        ' wynik !
        ' z tego pobieramy liste czesci!
        sess.FindById("wnd[0]/tbar[1]/btn[8]").press
        
        
        ' teraz glowny wrzut na input
        ' =========================================================================
        
        Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        Set cols = grid.ColumnOrder
        
        
        Dim refInput As Range
        Set refInput = ish.Range("A1")
        
        refInput.offset(0, 0).Value = "PRODUIT"
        refInput.offset(0, 1).Value = "DESIGNATION"
        refInput.offset(0, 2).Value = "RU"
        
        Set refInput = refInput.offset(1, 0)
        
        
        
        ' MAIN LOOP FOR FIRST LIST OF PN WITHOUT INDICE
        For x = 0 To grid.RowCount
            For y = 0 To grid.ColumnCount
            
            
                If CStr(cols(y)) Like "*_PRODUIT*" Then
                
                    ' Debug.Print Cstr(grid.GetCellValue(x2, cols(y2))) ' OK
                    refInput.Value = CStr(grid.getCellValue(x, cols(y)))
                    refInput.offset(0, 1).Value = CStr(grid.getCellValue(x, cols(y + 1)))
                    refInput.offset(0, 2).Value = CStr(grid.getCellValue(x, cols(y + 2)))
                    Exit For
                End If
                
            Next y
            
            
            

            If x Mod 50 = 0 Then
            
                grid.FirstVisibleRow = x
                grid.CurrentCellRow = x
                
                
                ' powazny side effect ...
                delegacjaDlaProgresu st_h, xStHelper, 20
                
            End If
            
            
            Set refInput = refInput.offset(1, 0)
        Next x
        
        makeDoubleListWithDashAndStar ish.Range("A2")
        
        st_h.hide
        Set st_h = Nothing
        Set st_h = New StatusHandler
        st_h.init_statusbar CLng((ish.Range("A1").End(xlDown).row) / 50)
        st_h.show
        st_h.progress_increase
        
        
        ' go back! x2
        sess.FindById("wnd[0]/tbar[0]/btn[15]").press
        sess.FindById("wnd[0]/tbar[0]/btn[15]").press
        
        
        ' teraz etap 2
        ' mamy powiekszona liste czesci - musimy przejsc jeszcze raz przez wybor
        
        ' tym razem wybieramy konkretnie
        
        runLoops "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell"
        
        
        sess.FindById("wnd[0]/tbar[1]/btn[17]").press
        ' otwieramy liste z dostepnymi standardami do etapu drugiego!
        sess.FindById("wnd[1]").sendVKey 4
        
        runLoops "wnd[2]/usr/cntlALV_CONTAINER_1/shellcont/shell"
        
        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        
        ' teraz powinnismy byc na oknie ponownego forma wejsciowego dla listy quasi tp04
        
        
    End If
    
    Dim rngToCopy As Range
    
    With sess
    
        Dim btn3 As Variant ' SAPFEWSELib.GuiButton
        Set btn3 = findMyButton() ' .FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH")
        ' not static
        ' .FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").press
        btn3.press
        .FindById("wnd[1]/tbar[0]/btn[16]").press
        Set rngToCopy = ish.Range(ish.Range("A2"), ish.Range("A2").End(xlDown))
        rngToCopy.Copy
        
        .FindById("wnd[1]/tbar[0]/btn[24]").press
        .FindById("wnd[1]/tbar[0]/btn[8]").press
        .FindById("wnd[0]/tbar[1]/btn[8]").press
    End With

    
    Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    Set cols = grid.ColumnOrder
    
    
    Dim refOutput2 As Range ', 'lm As Long
    Set refOutput2 = osh.Cells(2, 1)
    
    
    'lm = CLng((ish.Range("A1").End(xlDown).row) / 50)
    
    
    
    For x = 0 To grid.RowCount - 1
        For y = 0 To grid.ColumnCount - 1
        
        
            refOutput2.offset(0, y).Value = CStr(grid.getCellValue(x, cols(y)))
        Next y
        
        If x Mod 50 = 0 Then
        
            grid.FirstVisibleRow = x
            grid.CurrentCellRow = x
            
            ' delegacjaDlaProgresu st_h, xStHelper, 20
            st_h.progress_increase
            
        End If
        
        
        Set refOutput2 = refOutput2.offset(1, 0)
    Next x
    
End Sub


Private Function fillFormAndWaitForDecision(cfg1 As Integer) As String


    fillFormAndWaitForDecision = ""
    
    SQ01ConfigForm.ListBox1.Clear
    SQ01ConfigForm.ListBox1.MultiSelect = fmMultiSelectSingle
    
    If cfg1 = 1 Then
        ' pierwsze wybieranie!
        ' -----------------------------------------------------
        
        For x = 0 To grid.RowCount - 1
            SQ01ConfigForm.ListBox1.addItem CStr(grid.getCellValue(x, cols(0)))
        Next x
        
        ' -----------------------------------------------------
    End If
    
    SQ01ConfigForm.show
    
    On Error Resume Next
    fillFormAndWaitForDecision = SQ01ConfigForm.whatYouChoose
End Function




Private Sub makeDoubleListWithDashAndStar(refR As Range)
    
    Dim doubleR As Range, lastOrgRng As Range
    Set doubleR = refR.End(xlDown).End(xlDown).End(xlUp).offset(1, 0)
    
    If refR.offset(1, 0) <> "" Then
        
        Set lastOrgRng = refR.End(xlDown)
    Else
        Set lastOrgRng = refR
    End If
    
    Dim ir As Range
    For Each ir In refR.Parent.Range(refR, lastOrgRng)
        doubleR.Value = ir.Value & "-*"
        doubleR.offset(0, 1).Value = "*-*"
        Set doubleR = doubleR.offset(1, 0)
    Next ir
End Sub




Private Sub runLoops(stringReference As String)


    Dim wynik As String

    Set grid = sess.FindById(CStr(stringReference))
    Set cols = grid.ColumnOrder
    
    wynik = fillFormAndWaitForDecision(1)
    
    For x = 0 To grid.RowCount - 1
        If CStr(wynik) = CStr(grid.getCellValue(x, cols(0))) Then
        
            With grid
                .CurrentCellRow = x
                .SelectedRows = "" & CStr(x)
                .DoubleClickCurrentCell
            End With
            
            Exit For
            
        End If
    Next x

End Sub

Private Sub runLoopsForceString(stringReference As String, forceString As String)


    Dim wynik As String

    Set grid = sess.FindById(CStr(stringReference))
    Set cols = grid.ColumnOrder
    
    wynik = CStr(forceString)
    
    For x = 0 To grid.RowCount - 1
        If CStr(wynik) = CStr(grid.getCellValue(x, cols(0))) Then
        
            With grid
                .CurrentCellRow = x
                .SelectedRows = "" & CStr(x)
                .DoubleClickCurrentCell
            End With
            
            Exit For
            
        End If
    Next x

End Sub


Private Sub runLoopsWithPreDef(stringReference As String, preDefEnum As E_SQ01_PREDEF, preDefLine As Integer)


    Dim wynik As String

    Set grid = sess.FindById(CStr(stringReference))
    Set cols = grid.ColumnOrder
    
    wynik = getPreDef(preDefEnum, preDefLine)
    
    For x = 0 To grid.RowCount - 1
        If CStr(wynik) = CStr(grid.getCellValue(x, cols(0))) Then
        
            With grid
                .CurrentCellRow = x
                .SelectedRows = "" & CStr(x)
                .DoubleClickCurrentCell
            End With
            
            Exit For
            
        End If
    Next x

End Sub

Private Function getProperRefFromPreDef(preDefEnum As E_SQ01_PREDEF, preDefLine As Integer) As String
    getProperRefFromPreDef = getPreDef(preDefEnum, preDefLine)
End Function


Private Function getPreDef(preDefEnum As E_SQ01_PREDEF, m_getPreDef As Integer) As String

    If preDefEnum = E_SQ01_PREDEF_SYSTEM Then
        
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").offset(m_getPreDef - 1, 0).Value
    ElseIf preDefEnum = E_SQ01_PREDEF_LIST1 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").offset(m_getPreDef - 1, 1).Value
    ElseIf preDefEnum = E_SQ01_PREDEF_STD1 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").offset(m_getPreDef - 1, 2).Value
    ElseIf preDefEnum = E_SQ01_PREDEF_LIST2 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").offset(m_getPreDef - 1, 3).Value
    ElseIf preDefEnum = E_SQ01_PREDEF_STD2 Then
        getPreDef = "" & ThisWorkbook.Sheets("register").Range("PRE_DEF_RUN_FOR_SQ01").offset(getPreDef - 1, 4).Value
    Else
        MsgBox "Fatal error in pre-def logic inside foo getPreDef!", vbCritical
        End
    End If
End Function
