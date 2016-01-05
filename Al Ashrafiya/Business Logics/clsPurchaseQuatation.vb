Public Class clsPurchaseQuatation
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Events"
    Private Sub LoadForm(ByVal aform As SAPbouiCOM.Form, ByVal aRow As Integer)
        Dim oRec, oRec1, oRec2 As SAPbobsCOM.Recordset
        Dim oItem As SAPbobsCOM.Items
        Dim strItemcode, strCode As String
        oMatrix = aform.Items.Item("38").Specific
        strCardCode = oApplication.Utilities.getEdittextvalue(aform, "4")
        strItemcode = oApplication.Utilities.getMatrixValues(oMatrix, "1", aRow)
        strCode = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Cost_Ref", aRow)
        If strItemcode <> "" Then
            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            If oItem.GetByKey(strItemcode) Then
                Dim oob As New clsCosting
                frmSourceForm = aform
                frmSourceFormUD = aform.UniqueID
                intSelectedMatrixrow = aRow
                oob.LoadForm(oApplication.Utilities.getEdittextvalue(oForm, "8"), strItemcode, strCode, aRow, "Edit")
                '  Populate(strCode, oForm, aRow, strItemcode) 'oApplication.Utilities.getEdittextvalue(oForm, "8"), strItemcode, strCode, aRow, "Edit")
            End If
        End If
    End Sub

    Private Sub ViewCosting(ByVal aform As SAPbouiCOM.Form, ByVal aRow As Integer)
        Dim oRec, oRec1, oRec2 As SAPbobsCOM.Recordset
        Dim oItem As SAPbobsCOM.Items
        Dim strItemcode, strCode As String
        oMatrix = aform.Items.Item("38").Specific
        strCardCode = oApplication.Utilities.getEdittextvalue(aform, "4")
        strItemcode = oApplication.Utilities.getMatrixValues(oMatrix, "1", aRow)
        strCode = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Cost_Ref", aRow)
        If strItemcode <> "" Then
            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            If oItem.GetByKey(strItemcode) Then
                'If oItem.TreeType <> SAPbobsCOM.BoItemTreeTypes.iNotATree Then
                '    'oCombobox = oMatrix.Columns.Item("U_Z_Costreq").Cells.Item(aRow).Specific
                '    'If oCombobox.Selected.Value = "N" Then
                '    '    oApplication.Utilities.GetB1Price(strItemcode, strCardCode, oMatrix, aRow)
                '    'Else
                '    '    Dim oob As New clsCosting
                '    '    frmSourceForm = aform
                '    '    frmSourceFormUD = aform.UniqueID
                '    '    intSelectedMatrixrow = aRow
                '    '    oob.LoadForm(oApplication.Utilities.getEdittextvalue(oForm, "8"), strItemcode, strCode, aRow, "View")
                '    'End If
                'End If
                Dim oob As New clsCosting
                frmSourceForm = aform
                frmSourceFormUD = aform.UniqueID
                intSelectedMatrixrow = aRow
                oob.LoadForm(oApplication.Utilities.getEdittextvalue(oForm, "8"), strItemcode, strCode, aRow, "View")
            End If
        End If
    End Sub

    Private Sub Resetprice(ByVal aform As SAPbouiCOM.Form, ByVal aRow As Integer)
        Dim oRec, oRec1, oRec2 As SAPbobsCOM.Recordset
        Dim oItem As SAPbobsCOM.Items
        Dim strItemcode, strCode As String
        oMatrix = aform.Items.Item("38").Specific
        strCardCode = oApplication.Utilities.getEdittextvalue(aform, "4")
        strItemcode = oApplication.Utilities.getMatrixValues(oMatrix, "1", aRow)
        strCode = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Cost_Ref", aRow)
        If strItemcode <> "" Then
            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            If oItem.GetByKey(strItemcode) Then
                If oItem.TreeType <> SAPbobsCOM.BoItemTreeTypes.iNotATree Then
                    'oCombobox = oMatrix.Columns.Item("U_Z_Costreq").Cells.Item(aRow).Specific
                    'If oCombobox.Selected.Value = "N" Then
                    '    oApplication.Utilities.GetB1Price(strItemcode, strCardCode, oMatrix, aRow)
                    '    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Ref", aRow, "")
                    'Else
                    '    Dim oob As New clsCosting
                    '    frmSourceForm = aform
                    '    frmSourceFormUD = aform.UniqueID
                    '    intSelectedMatrixrow = aRow
                    '    oob.LoadForm(oApplication.Utilities.getEdittextvalue(oForm, "8"), strItemcode, strCode, aRow, "Edit")
                    'End If
                End If
            End If
        End If
    End Sub
#End Region

    Public Sub Populate(ByVal aform As SAPbouiCOM.Form)
        Dim strsql, strsql1 As String
        Dim oTest, oTest1, otest2, oPO As SAPbobsCOM.Recordset
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode, aDocNum, LineRef As String
        Try
            aform.Freeze(True)
            Dim oRec, oRec1, oRec2 As SAPbobsCOM.Recordset
            Dim oItem As SAPbobsCOM.Items
            Dim strItemcode, strRefCode, strVendor, strCardNameName, strNumAtCard, strQDate As String
            oMatrix = aform.Items.Item("38").Specific
            strCardCode = oApplication.Utilities.getEdittextvalue(aform, "4")
            strCardNameName = oApplication.Utilities.getEdittextvalue(aform, "54")
            strNumAtCard = oApplication.Utilities.getEdittextvalue(aform, "16")
            aDocNum = oApplication.Utilities.getEdittextvalue(aform, "8")
            strQDate = oApplication.Utilities.getEdittextvalue(aform, "10")
            Dim dtDate, dtDate1, dtDocDate, dtDate2, dtDate3 As Date
            ' dtDate = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aform, "10"))
            ' dtDate1 = DateAdd(DateInterval.Month, -3, dtDate)
            oPO = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            dtDate = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aform, "10"))
            dtDate1 = DateAdd(DateInterval.Month, -3, dtDate)
            dtDate2 = dtDate
            dtDate3 = DateAdd(DateInterval.Day, -30, dtDate2)

         

            Dim dblInStock, dblOrderQty, dblCurrentPrice, dblTotalPrice, dbllandedcost, dblLastpurchasedPrice, dblAvgSalesQty, dblAvgSalesAmount, dblAveSalesQuantity, dblAverageSalesAmount As Double
            LineRef = ""
            Dim strPQCurrency As String
            Dim blnApprovedform As Boolean = False
            If aform.Title.ToUpper.Contains("APPROVED") Or aform.Title.ToUpper.Contains("PENDING") Or aform.Title.ToUpper.Contains("APPR") Then
                blnApprovedform = True
            Else
                blnApprovedform = False
            End If
            oUserTable = oApplication.Company.UserTables.Item("Z_AL_COST")
            For introw As Integer = 1 To oMatrix.RowCount
                oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                strItemcode = oApplication.Utilities.getMatrixValues(oMatrix, "1", introw)
                strRefCode = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Cost_Ref", introw)
                If strItemcode <> "" Then
                    strsql = "Select * from [@Z_AL_COST] where U_Z_DocNum='" & aDocNum & "' and  U_Z_ItemCode='" & strItemcode & "' and  code='" & strRefCode & "'"
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    otest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest1.DoQuery("Select * from OITM where ItemCode='" & strItemcode & "'")
                    oTest.DoQuery(strsql)

                    strsql = "Select sum(a.quantity)/3 from (select sum(invqty) as quantity from inv1 a INNER join OINV b on a.docentry = b.docentry"
                    strsql = strsql & " where  b.docdate >='" & dtDate1.ToString("yyyy-MM-dd") & " ' and b.docdate<='" & dtDate.ToString("yyyy-MM-dd") & "' and a.itemcode ='" & strItemcode & "'"
                    strsql = strsql & "  union all select -sum(invqty) from rin1 a INNER join orin b on a.docentry = b.docentry"
                    strsql = strsql & " where    b.docdate  >='" & dtDate1.ToString("yyyy-MM-dd") & " ' and b.docdate<='" & dtDate.ToString("yyyy-MM-dd") & "' and a.itemcode ='" & strItemcode & "' and a.NoInvtryMv='N') a"
                    otest2.DoQuery(strsql)

                    dblAvgSalesQty = otest2.Fields.Item(0).Value

                    strsql = "Select sum(a.quantity ) from (select sum(LineTotal) as quantity  from inv1 a INNER join OINV b on a.docentry = b.docentry"
                    strsql = strsql & " where  b.docdate >='" & dtDate1.ToString("yyyy-MM-dd") & " ' and b.docdate<='" & dtDate.ToString("yyyy-MM-dd") & "' and a.itemcode ='" & strItemcode & "'"
                    strsql = strsql & "  union all select -sum(LineTotal) from rin1 a INNER join orin b on a.docentry = b.docentry"
                    strsql = strsql & " where    b.docdate  >='" & dtDate1.ToString("yyyy-MM-dd") & " ' and b.docdate<='" & dtDate.ToString("yyyy-MM-dd") & "' and a.itemcode ='" & strItemcode & "') a"
                    otest2.DoQuery(strsql)
                    Dim strDocCurrency, strDocType As String
                    Dim dblExchangeRate, dblOpenPOQty As Double
                    Dim oExRS As SAPbobsCOM.Recordset
                    oExRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    oCombobox = aform.Items.Item("70").Specific
                    strDocType = oCombobox.Selected.Value
                    If strDocType = "L" Then
                        strDocCurrency = oApplication.Company.GetCompanyService.GetAdminInfo.LocalCurrency
                    ElseIf strDocType = "S" Then
                        strDocCurrency = oApplication.Company.GetCompanyService.GetAdminInfo.SystemCurrency
                    Else
                        oCombobox = aform.Items.Item("63").Specific
                        strDocCurrency = oCombobox.Selected.Value
                    End If
                    Try
                        oExRS.DoQuery("Select * from ORTT where RateDate='" & Now.Date.ToString("yyyy-MM-dd") & "' and Currency='" & strDocCurrency & "'")
                        If oExRS.RecordCount > 0 Then
                            dblExchangeRate = oExRS.Fields.Item("Rate").Value
                        Else
                            dblExchangeRate = 1
                        End If
                    Catch ex As Exception
                        dblExchangeRate = 1
                    End Try

                    dblAvgSalesAmount = otest2.Fields.Item(0).Value
                    dblAvgSalesAmount = dblAvgSalesAmount / (dblAvgSalesQty * 1) ' 3)


                    strsql = "Select sum(a.quantity) from (select sum(invqty) as quantity from inv1 a INNER join OINV b on a.docentry = b.docentry"
                    strsql = strsql & " where  b.docdate >='" & dtDate3.ToString("yyyy-MM-dd") & " ' and b.docdate<='" & dtDate2.ToString("yyyy-MM-dd") & "' and a.itemcode ='" & strItemcode & "'"
                    strsql = strsql & "  union all select -sum(invqty) from rin1 a INNER join orin b on a.docentry = b.docentry"
                    strsql = strsql & " where    b.docdate  >='" & dtDate3.ToString("yyyy-MM-dd") & " ' and b.docdate<='" & dtDate2.ToString("yyyy-MM-dd") & "' and a.itemcode ='" & strItemcode & "' and a.NoInvtryMv='N') a"
                    otest2.DoQuery(strsql)

                    dblAveSalesQuantity = otest2.Fields.Item(0).Value

                    strsql = "Select sum(a.quantity ) from (select sum(LineTotal) as quantity  from inv1 a INNER join OINV b on a.docentry = b.docentry"
                    strsql = strsql & " where  b.docdate >='" & dtDate3.ToString("yyyy-MM-dd") & " ' and b.docdate<='" & dtDate2.ToString("yyyy-MM-dd") & "' and a.itemcode ='" & strItemcode & "'"
                    strsql = strsql & "  union all select -sum(LineTotal) from rin1 a INNER join orin b on a.docentry = b.docentry" '
                    strsql = strsql & " where    b.docdate  >='" & dtDate3.ToString("yyyy-MM-dd") & " ' and b.docdate<='" & dtDate2.ToString("yyyy-MM-dd") & "' and a.itemcode ='" & strItemcode & "' and a.NoInvtryMv<>'Y' union all select -sum(LineTotal) from rin1 a INNER join orin b on a.docentry = b.docentry" '
                    strsql = strsql & " where    b.docdate  >='" & dtDate3.ToString("yyyy-MM-dd") & " ' and b.docdate<='" & dtDate2.ToString("yyyy-MM-dd") & "' and a.itemcode ='" & strItemcode & "' and a.NoInvtryMv<>'N') a"
                    otest2.DoQuery(strsql)

                    dblAverageSalesAmount = otest2.Fields.Item(0).Value
                    Dim dblAverageSales As Double
                    dblAverageSales = dblAverageSalesAmount / dblAveSalesQuantity

                    dblOrderQty = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "11", introw))
                    dblCurrentPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "14", introw))
                    strPQCurrency = oApplication.Utilities.getMatrixValues(oMatrix, "14", introw)
                    dblTotalPrice = dblOrderQty * dblCurrentPrice

                    Dim dblPercentage As Double
                    Dim bForm As SAPbouiCOM.Form
                    bForm = oApplication.SBO_Application.Forms.GetForm(frm_PurchaseQuatationUDF, aform.TypeCount)
                    dblPercentage = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(bForm, "U_Z_LandedCost"))
                    'dbllandedcost = dblCurrentPrice * 7.5 * 0.22 / 100'
                 
                    'dbllandedcost = dblCurrentPrice * dblPercentage * 0.22 / 100 '
                    dbllandedcost = (dblCurrentPrice * (dblPercentage / 100) * dblExchangeRate) + (dblCurrentPrice * dblExchangeRate) ')
                    dblLastpurchasedPrice = oTest1.Fields.Item("LastPurPrc").Value
                    Dim dblMonth As Double
                    dblMonth = (dblAverageSales - dbllandedcost) / dblAverageSales
                    dblMonth = dblMonth * 100
                    Dim dblCurrentPrice1 As Double = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "14", introw))
                    'oPO.DoQuery("SELECT sum(T1.Quantity)  FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry WHERE T1.[LineStatus] <>'C' and T0.Cardcode='" & strCardCode & "' and T1.ItemCode='" & strItemcode & "'")
                    oPO.DoQuery("SELECT sum(T1.OpenQty)  FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry WHERE T1.[LineStatus] <>'C' and T1.ItemCode='" & strItemcode & "'")
                    If oPO.RecordCount > 0 Then
                        dblOpenPOQty = oPO.Fields.Item(0).Value
                    Else
                        dblOpenPOQty = 0
                    End If

                    If oTest.RecordCount > 0 Then
                        strCode = oTest.Fields.Item("Code").Value
                        If oUserTable.GetByKey(strCode) Then
                            oUserTable.Name = strCode & "N"
                            oUserTable.UserFields.Fields.Item("U_Z_DocNum").Value = aDocNum
                            oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = strItemcode
                            oUserTable.UserFields.Fields.Item("U_Z_ItemName").Value = oTest1.Fields.Item("ItemName").Value
                            oUserTable.UserFields.Fields.Item("U_Z_Origion").Value = oTest1.Fields.Item("U_COO").Value
                            oUserTable.UserFields.Fields.Item("U_Z_Unit").Value = oTest1.Fields.Item("BuyUnitMsr").Value
                            oUserTable.UserFields.Fields.Item("U_Z_LineID").Value = introw
                            oUserTable.UserFields.Fields.Item("U_Z_Packaging").Value = oTest1.Fields.Item("U_Pack").Value 'oTest1.Fields.Item("PurPackMsr").Value '(oApplication.Utilities.getEdittextvalue(aform, "24"))
                            oUserTable.UserFields.Fields.Item("U_Z_OrderQty").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "11", introw))
                            oUserTable.UserFields.Fields.Item("U_Z_InStock").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "33", introw))
                            oUserTable.UserFields.Fields.Item("U_Z_PPrice").Value = oTest1.Fields.Item("LastPurPrc").Value 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "31"))
                            oUserTable.UserFields.Fields.Item("U_Z_Price").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "14", introw))
                            oUserTable.UserFields.Fields.Item("U_Z_Totalvalue").Value = dblTotalPrice ' oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "36"))
                            oUserTable.UserFields.Fields.Item("U_Z_LandedCost").Value = dbllandedcost ' oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "38"))
                            oUserTable.UserFields.Fields.Item("U_Z_AvgMonSal").Value = dblAvgSalesQty  'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "40"))
                            oUserTable.UserFields.Fields.Item("U_Z_AvgSelling").Value = dblAverageSales ' dblAvgSalesAmount  'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "40"))
                            oUserTable.UserFields.Fields.Item("U_Z_Month").Value = dblMonth 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "42"))
                            oUserTable.UserFields.Fields.Item("U_Z_PurCur").Value = oTest1.Fields.Item("LastPurCur").Value
                            oUserTable.UserFields.Fields.Item("U_Z_PQCur").Value = strDocCurrency
                            oUserTable.UserFields.Fields.Item("U_Z_POQty").Value = dblOpenPOQty
                            oUserTable.UserFields.Fields.Item("U_Z_CardCode").Value = strCardCode
                            oUserTable.UserFields.Fields.Item("U_Z_CardName").Value = strCardNameName
                            oUserTable.UserFields.Fields.Item("U_Z_NumAtCard").Value = strNumAtCard
                            oUserTable.UserFields.Fields.Item("U_Z_SE").Value = oTest1.Fields.Item("U_SE").Value
                            oUserTable.UserFields.Fields.Item("U_Z_CE").Value = oTest1.Fields.Item("U_CE").Value
                            If strQDate <> "" Then
                                dtDocDate = oApplication.Utilities.GetDateTimeValue(strQDate)
                                oUserTable.UserFields.Fields.Item("U_Z_QDate").Value = dtDocDate
                            Else
                                oUserTable.UserFields.Fields.Item("U_Z_QDate").Value = Now.Date

                            End If

                            ' oUserTable.UserFields.Fields.Item("U_Z_PropMarging").Value = 0 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "45"))
                            ' oUserTable.UserFields.Fields.Item("U_Z_Margin").Value = 0 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "46"))
                            If oUserTable.Update <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else
                            End If
                        End If
                    Else

                        strCode = oApplication.Utilities.getMaxCode("@Z_AL_COST", "Code")
                        Dim stName As String
                        stName = strCode & "N"
                        oUserTable.Code = strCode
                        oUserTable.Name = strCode & "N"
                        oUserTable.UserFields.Fields.Item("U_Z_DocNum").Value = aDocNum
                        oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = strItemcode
                        oUserTable.UserFields.Fields.Item("U_Z_ItemName").Value = oTest1.Fields.Item("ItemName").Value
                        oUserTable.UserFields.Fields.Item("U_Z_LineID").Value = introw
                        oUserTable.UserFields.Fields.Item("U_Z_Origion").Value = oTest1.Fields.Item("U_COO").Value
                        oUserTable.UserFields.Fields.Item("U_Z_Unit").Value = oTest1.Fields.Item("BuyUnitMsr").Value
                        oUserTable.UserFields.Fields.Item("U_Z_Packaging").Value = oTest1.Fields.Item("U_Pack").Value 'oTest1.Fields.Item("PurPackMsr").Value '('oApplication.Utilities.getEdittextvalue(aform, "24"))
                        oUserTable.UserFields.Fields.Item("U_Z_OrderQty").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "11", introw))
                        oUserTable.UserFields.Fields.Item("U_Z_InStock").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "33", introw))
                        oUserTable.UserFields.Fields.Item("U_Z_PPrice").Value = dblLastpurchasedPrice 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "31"))
                        oUserTable.UserFields.Fields.Item("U_Z_Price").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "14", introw))
                        oUserTable.UserFields.Fields.Item("U_Z_Totalvalue").Value = dblTotalPrice ' oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "36"))
                        oUserTable.UserFields.Fields.Item("U_Z_LandedCost").Value = dbllandedcost ' oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "38"))
                        oUserTable.UserFields.Fields.Item("U_Z_AvgMonSal").Value = dblAvgSalesQty  'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "40"))
                        oUserTable.UserFields.Fields.Item("U_Z_AvgSelling").Value = dblAverageSales ' dblAvgSalesAmount  'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "40"))
                        oUserTable.UserFields.Fields.Item("U_Z_Month").Value = dblMonth 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "42"))
                        oUserTable.UserFields.Fields.Item("U_Z_PropMarging").Value = 0 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "45"))
                        oUserTable.UserFields.Fields.Item("U_Z_Margin").Value = 0 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "46"))
                        oUserTable.UserFields.Fields.Item("U_Z_PurCur").Value = oTest1.Fields.Item("LastPurCur").Value
                        oUserTable.UserFields.Fields.Item("U_Z_PQCur").Value = strDocCurrency
                        oUserTable.UserFields.Fields.Item("U_Z_POQty").Value = dblOpenPOQty
                        oUserTable.UserFields.Fields.Item("U_Z_CardCode").Value = strCardCode
                        oUserTable.UserFields.Fields.Item("U_Z_CardName").Value = strCardNameName
                        oUserTable.UserFields.Fields.Item("U_Z_NumAtCard").Value = strNumAtCard
                        oUserTable.UserFields.Fields.Item("U_Z_SE").Value = oTest1.Fields.Item("U_SE").Value
                        oUserTable.UserFields.Fields.Item("U_Z_CE").Value = oTest1.Fields.Item("U_CE").Value
                        If strQDate <> "" Then
                            dtDocDate = oApplication.Utilities.GetDateTimeValue(strQDate)
                            oUserTable.UserFields.Fields.Item("U_Z_QDate").Value = dtDocDate
                        Else
                            oUserTable.UserFields.Fields.Item("U_Z_QDate").Value = Now.Date

                        End If
                        If blnApprovedform = False Then
                            If oUserTable.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else

                            End If
                        End If
                    End If
                    oTest.DoQuery("Update [@Z_AL_COST] set U_Z_Margin= ((U_Z_PropMarging-U_Z_LandedCost)/isnull(U_Z_PropMarging,1))*100 where U_Z_PropMarging >0 ")
                    If blnApprovedform = False Then
                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Cost_Ref", introw, strCode)
                    End If

                    If LineRef = "" Then
                        LineRef = "'" & strCode & "'"
                    Else
                        LineRef = LineRef & ",'" & strCode & "'"
                    End If
                End If
            Next
            oApplication.Utilities.LoadForm(xml_CostView, frm_CostView)
            Dim objForm As SAPbouiCOM.Form
            objForm = oApplication.SBO_Application.Forms.ActiveForm()
            If objForm.TypeEx = frm_CostView Then
                Try
                    objForm.Freeze(True)
                    Dim ogrid As SAPbouiCOM.Grid
                    ogrid = objForm.Items.Item("1").Specific
                    objForm.Items.Item("1").Enabled = True
                    Dim strstring As String
                    strstring = "SELECT T0.[Code], T0.[Name], T0.[U_Z_DocEntry] , T0.[U_Z_DocNum] 'Document Number', T0.[U_Z_ItemCode] 'Item Code', T0.[U_Z_ItemName] ' Item Name', T0.[U_Z_Origion] 'Origion', T0.[U_Z_LineID] 'Line Number', T0.[U_Z_Unit] 'Unit' , T0.[U_Z_Packaging] 'Packaging', T0.[U_Z_OrderQty] 'OrderQty', T0.[U_Z_InStock] 'In Stock',T0.U_Z_POQTY 'Qty in Open Purchase Order ', T0.[U_Z_AvgMonSal] 'Average Monthly Sales', T0.[U_Z_PurCur] 'Last Purchased Currency' ,T0.[U_Z_PPrice] 'Last Purchased Price', T0.U_Z_PQCur 'Purchase Quotation Currency',T0.[U_Z_Price] 'Purchage Quation Price', T0.[U_Z_Totalvalue] 'Total Value ', T0.[U_Z_LandedCost] 'Landed Cost in KWD', T0.[U_Z_AvgSelling] 'Average Selling Price', T0.[U_Z_Month] 'Monthly %', T0.[U_Z_PropMarging] 'Proposed', T0.[U_Z_Margin] 'Margin%' FROM [dbo].[@Z_AL_COST]  T0"
                    strstring = strstring & " where Code in (" & LineRef & ")"
                    ogrid.DataTable.ExecuteQuery(strstring)
                    'ogrid.DataTable.ExecuteQuery("SElect * from [@Z_AL_COST] where U_Z_DocNum='" & aDocNum & "' and Code in (" & LineRef & ")")
                    ogrid.Columns.Item(0).Visible = False
                    ogrid.Columns.Item(1).Visible = False
                    ogrid.Columns.Item(2).Visible = False
                    For introw As Integer = 3 To 21
                        ogrid.Columns.Item(introw).Editable = False
                    Next
                    Dim oEditTextColumn As SAPbouiCOM.EditTextColumn = ogrid.Columns.Item("Item Code")
                    oEditTextColumn.LinkedObjectType = "4"
                    If blnApprovedform = False Then
                        ogrid.Columns.Item("Proposed").Editable = True
                        objform.Items.Item("3").Enabled=True 
                    Else
                        ogrid.Columns.Item("Proposed").Editable = False
                        objForm.Items.Item("3").Enabled = False
                    End If
                    ogrid.Columns.Item("Proposed").TitleObject.Caption = "Proposed Selling Price"
                    ogrid.AutoResizeColumns()
                    ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
                    objForm.Freeze(False)
                Catch ex As Exception
                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objForm.Freeze(False)
                End Try
            End If
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_PurchaseQuatation Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    'For intRow As Integer = 1 To oMatrix.RowCount
                                    '    Dim stitem As String
                                    '    stitem = oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow)
                                    '    If stitem <> "" Then
                                    '        oCombobox = oMatrix.Columns.Item("U_Z_Costreq").Cells.Item(intRow).Specific
                                    '        If oCombobox.Selected.Value = "Y" Then
                                    '            If oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Ref", intRow) = "" Then
                                    '                oApplication.Utilities.Message("Costing sheet not calcualted for the item : " & stitem, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    '                oMatrix.Columns.Item("1").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    '                BubbleEvent = False
                                    '                Exit Sub
                                    '            End If
                                    '        End If
                                    '    End If
                                    'Next
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_Cost_Ref" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oApplication.Utilities.AddControls(oForm, "btnView", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", , , "2", "Cost Analysis", 120)
                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                ' oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_Cost_Ref" Then
                                    ViewCosting(oForm, pVal.Row)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "btnView" Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    Populate(oForm)
                                    'For intRow As Integer = 1 To oMatrix.RowCount
                                    '    If oMatrix.IsRowSelected(intRow) Then
                                    '        ViewCosting(oForm, intRow)
                                    '        Exit Sub
                                    '    End If
                                    'Next
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                oApplication.Utilities.AddControls(oForm, "btnView", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", , , "2", "Costing Sheet", 120)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "1" And pVal.CharPressed = 9 Then
                                    '  LoadForm(oForm, pVal.Row)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_Costreq" Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    'oCombobox = oMatrix.Columns.Item("U_Z_Cost_req").Cells.Item(pVal.Row).Specific
                                    'If oCombobox.Selected.Value = "N" Then
                                    '    Resetprice(oForm, pVal.Row)
                                    'Else
                                    '    LoadForm(oForm, pVal.Row)
                                    'End If

                                    ' LoadForm(oForm, pVal.Row)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim sCHFL_ID As String
                                Dim intChoice As Integer
                                Dim codebar, val1 As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "38" And pVal.ColUID = "1" Then
                                            oMatrix = oForm.Items.Item("38").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Cost_Ref", pVal.Row, "")
                                            oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            Dim st As String = oDataTable.GetValue("ItemCode", 0)
                                            ' Populate(oForm, pVal.Row)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                End Try
                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_DuplicateRow
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        If oForm.TypeEx = frm_PurchaseQuatation Then
                            oMatrix = oForm.Items.Item("38").Specific
                            If intCurrentRow <> 10000 Then
                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Cost_Ref", intCurrentRow + 1, "")
                            End If
                        End If
                    Else
                        If oForm.TypeEx = frm_PurchaseQuatation Then
                            'oMatrix = oForm.Items.Item("38").Specific
                            'If intCurrentRow <> 10000 Then
                            '    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Ref", intCurrentRow + 1, "")
                            'End If
                        End If
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)


        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        If eventInfo.BeforeAction = True Then
            If oForm.TypeEx = frm_PurchaseQuatation And eventInfo.ItemUID = "38" Then
                oMatrix = oForm.Items.Item("38").Specific
                intCurrentRow = eventInfo.Row
            End If
        End If

    End Sub
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If BusinessObjectInfo.FormTypeEx = frm_PurchaseQuatation Then
                    Dim oDoc As SAPbobsCOM.Documents
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations)
                    If oDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        Dim orec As SAPbobsCOM.Recordset
                        Dim strCode As String
                        orec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If BusinessObjectInfo.Type <> "112" Then
                            For intRow As Integer = 0 To oDoc.Lines.Count - 1
                                oDoc.Lines.SetCurrentLine(intRow)
                                strCode = oDoc.Lines.UserFields.Fields.Item("U_Z_Cost_Ref").Value
                                If strCode <> "" Then
                                    orec.DoQuery("update [@Z_AL_COST] set  U_Z_LineId=" & oDoc.Lines.LineNum & ", U_Z_DocEntry=" & oDoc.DocEntry & ", U_Z_DocNum=" & oDoc.DocNum & " where Code='" & strCode & "'")
                                End If
                            Next
                            'orec.DoQuery("Delete from [@Z_COST] where U_Z_DocNum=" & oDoc.DocNum & " and isnull(U_Z_DocEntry,9999)=9999")
                        Else

                        End If
                    End If



                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
