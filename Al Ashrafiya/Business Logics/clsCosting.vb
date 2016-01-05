Public Class clsCosting
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
    Public Sub LoadForm(ByVal aDocNum As String, ByVal aItemCode As String, ByVal aRefCode As String, ByVal aRow As Integer, ByVal aChoice As String)
        Dim oRec, oRec1, oRec2 As SAPbobsCOM.Recordset
        oApplication.Utilities.LoadForm(xml_CostingSheet, frm_CostElement)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oApplication.Utilities.setEdittextvalue(oForm, "4", aDocNum)
        oApplication.Utilities.setEdittextvalue(oForm, "10", aRefCode)
        oApplication.Utilities.setEdittextvalue(oForm, "12", aItemCode)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select Itemname from OITM where ItemCode='" & aItemCode & "'")
        oApplication.Utilities.setEdittextvalue(oForm, "14", oRec.Fields.Item(0).Value)


        Populate(aRefCode, oForm, aRow, aItemCode)
        ' CalculateTotal(oForm)
        If aChoice = "Edit" Then
            oForm.Items.Item("3").Enabled = True
            oForm.Items.Item("62").Enabled = True
        Else
            oForm.Items.Item("3").Enabled = False
            oForm.Items.Item("62").Enabled = False
        End If
        oForm.PaneLevel = 1
    End Sub
#End Region

    Public Sub Populate(ByVal acode As String, ByVal aform As SAPbouiCOM.Form, ByVal aRow As Integer, ByVal aItemCode As String)
        Dim strsql, strsql1 As String
        Dim oTest, oTest1 As SAPbobsCOM.Recordset
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode, aDocNum As String
        Try
            aform.Freeze(True)
            aDocNum = oApplication.Utilities.getEdittextvalue(aform, "4")
            strsql = "Select * from [@Z_AL_COST] where U_Z_DocNum='" & aDocNum & "' and  U_Z_ItemCode='" & aItemCode & "' and  code='" & acode & "'"
            ' strsql = "Select * from [@Z_COST] where  code='" & acode & "'"
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTest.DoQuery(strsql)
            If oTest.RecordCount > 0 Then
                oApplication.Utilities.setEdittextvalue(aform, "4", oTest.Fields.Item("U_Z_DocNum").Value)
                oApplication.Utilities.setEdittextvalue(aform, "12", oTest.Fields.Item("U_Z_ItemCode").Value)
                oApplication.Utilities.setEdittextvalue(aform, "14", oTest.Fields.Item("U_Z_ItemName").Value)
                oApplication.Utilities.setEdittextvalue(aform, "20", oTest.Fields.Item("U_Z_Origion").Value)
                oApplication.Utilities.setEdittextvalue(aform, "22", oTest.Fields.Item("U_Z_Unit").Value)
                oApplication.Utilities.setEdittextvalue(aform, "24", oTest.Fields.Item("U_Z_Packaging").Value)
                oApplication.Utilities.setEdittextvalue(aform, "26", oTest.Fields.Item("U_Z_OrderQty").Value)
                oApplication.Utilities.setEdittextvalue(aform, "30", oTest.Fields.Item("U_Z_InStock").Value)
                oApplication.Utilities.setEdittextvalue(aform, "31", oTest.Fields.Item("U_Z_AvgMonSal").Value)

                oApplication.Utilities.setEdittextvalue(aform, "32", oTest.Fields.Item("U_Z_PPrice").Value)
                oApplication.Utilities.setEdittextvalue(aform, "34", oTest.Fields.Item("U_Z_Price").Value)
                oApplication.Utilities.setEdittextvalue(aform, "36", oTest.Fields.Item("U_Z_Totalvalue").Value)
                oApplication.Utilities.setEdittextvalue(aform, "38", oTest.Fields.Item("U_Z_LandedCost").Value)
                oApplication.Utilities.setEdittextvalue(aform, "40", oTest.Fields.Item("U_Z_AvgSelling").Value)

                oApplication.Utilities.setEdittextvalue(aform, "42", oTest.Fields.Item("U_Z_Month").Value)
                oApplication.Utilities.setEdittextvalue(aform, "45", oTest.Fields.Item("U_Z_PropMarging").Value)
                oApplication.Utilities.setEdittextvalue(aform, "46", oTest.Fields.Item("U_Z_Margin").Value)
            Else
                oUserTable = oApplication.Company.UserTables.Item("Z_AL_COST")
                strCode = oApplication.Utilities.getMaxCode("@Z_AL_COST", "Code")
                Dim stName As String
                stName = strCode & "N"
                oUserTable.Code = strCode
                oUserTable.Name = strCode & "N"
                oUserTable.UserFields.Fields.Item("U_Z_DocNum").Value = oApplication.Utilities.getEdittextvalue(aform, "4")
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = oApplication.Utilities.getEdittextvalue(aform, "12")
                If oUserTable.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    oApplication.Utilities.setEdittextvalue(aform, "10", stName)
                End If
            End If
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub


    Public Function AddtoUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strsql, aCode, strBaseRef As String
        Dim oTest, oTest1 As SAPbobsCOM.Recordset
        Dim oRec As SAPbobsCOM.Recordset
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode As String
        aCode = oApplication.Utilities.getEdittextvalue(aForm, "10")
        aCode = aCode.Replace("N", "")
        ' CalculateTotal(aForm)
        '   Exit Function
        strsql = "Select * from [@Z_AL_COST] where code='" & aCode & "'"
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery(strsql)
        If oTest.RecordCount > 0 Then
            strCode = aCode
            oUserTable = oApplication.Company.UserTables.Item("Z_AL_COST")
            oUserTable.GetByKey(strCode)
            oUserTable.Code = strCode
            oUserTable.Name = strCode

            oUserTable.UserFields.Fields.Item("U_Z_DocNum").Value = oApplication.Utilities.getEdittextvalue(aForm, "4")
            oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = oApplication.Utilities.getEdittextvalue(aForm, "12")
            oUserTable.UserFields.Fields.Item("U_Z_ItemName").Value = oApplication.Utilities.getEdittextvalue(aForm, "14")
        
            oUserTable.UserFields.Fields.Item("U_Z_Origion").Value = (oApplication.Utilities.getEdittextvalue(aForm, "20"))
            oUserTable.UserFields.Fields.Item("U_Z_Unit").Value = (oApplication.Utilities.getEdittextvalue(aForm, "22"))
            oUserTable.UserFields.Fields.Item("U_Z_Packaging").Value = (oApplication.Utilities.getEdittextvalue(aForm, "24"))
            oUserTable.UserFields.Fields.Item("U_Z_OrderQty").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "26"))
            oUserTable.UserFields.Fields.Item("U_Z_InStock").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "30"))

            oUserTable.UserFields.Fields.Item("U_Z_PPrice").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "31"))
            oUserTable.UserFields.Fields.Item("U_Z_Price").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "32"))
            oUserTable.UserFields.Fields.Item("U_Z_Totalvalue").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "36"))
            oUserTable.UserFields.Fields.Item("U_Z_LandedCost").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "38"))
            oUserTable.UserFields.Fields.Item("U_Z_AvgSelling").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "40"))
            oUserTable.UserFields.Fields.Item("U_Z_Month").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "42"))
            oUserTable.UserFields.Fields.Item("U_Z_PropMarging").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "45"))
            oUserTable.UserFields.Fields.Item("U_Z_Margin").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "46"))


           
            If oUserTable.Update <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else

            End If
        Else
            oUserTable = oApplication.Company.UserTables.Item("Z_COST")
            strCode = oApplication.Utilities.getMaxCode("@Z_COST", "Code")
            oUserTable.Code = strCode
            oUserTable.Name = strCode & "N"
            oUserTable.UserFields.Fields.Item("U_Z_DocNum").Value = oApplication.Utilities.getEdittextvalue(aForm, "4")
            oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = oApplication.Utilities.getEdittextvalue(aForm, "12")
            oUserTable.UserFields.Fields.Item("U_Z_ItemName").Value = oApplication.Utilities.getEdittextvalue(aForm, "14")
            oUserTable.UserFields.Fields.Item("U_Z_Origion").Value = (oApplication.Utilities.getEdittextvalue(aForm, "20"))
            oUserTable.UserFields.Fields.Item("U_Z_Unit").Value = (oApplication.Utilities.getEdittextvalue(aForm, "22"))
            oUserTable.UserFields.Fields.Item("U_Z_Packaging").Value = (oApplication.Utilities.getEdittextvalue(aForm, "24"))
            oUserTable.UserFields.Fields.Item("U_Z_OrderQty").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "26"))
            oUserTable.UserFields.Fields.Item("U_Z_InStock").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "30"))

            oUserTable.UserFields.Fields.Item("U_Z_PPrice").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "31"))
            oUserTable.UserFields.Fields.Item("U_Z_Price").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "32"))
            oUserTable.UserFields.Fields.Item("U_Z_Totalvalue").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "36"))
            oUserTable.UserFields.Fields.Item("U_Z_LandedCost").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "38"))
            oUserTable.UserFields.Fields.Item("U_Z_AvgSelling").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "40"))
            oUserTable.UserFields.Fields.Item("U_Z_Month").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "42"))
            oUserTable.UserFields.Fields.Item("U_Z_PropMarging").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "45"))
            oUserTable.UserFields.Fields.Item("U_Z_Margin").Value = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "46"))

            If oUserTable.Add <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                oApplication.Utilities.setEdittextvalue(aForm, "10", strCode)
            End If
        End If
        strBaseRef = strcode
        Return True
    End Function
#Region "Calculate Total"
    Private Sub PopulateInsurance(ByVal aform As SAPbouiCOM.Form)
        Dim dblTotal1, dblPercentage1, dblNet1, dblCurrency1 As Double
        Dim dblTotal11, dblNet11, dblPer11 As Double
        Dim dblFinal1 As Decimal
        Dim oRec1 As SAPbobsCOM.Recordset
        Dim dblInsurance1 As Double
        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        dblTotal11 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "55"))
        dblNet1 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "57"))
        oRec1.DoQuery("select isnull(U_Z_CreditPer,0),isnull(U_Z_Insurance,0) from OADM")
        dblPer11 = oRec1.Fields.Item(0).Value
        dblInsurance1 = oRec1.Fields.Item(1).Value
        oCombobox = aform.Items.Item("cmbInsuran").Specific
        dblNet11 = ((dblTotal11 * dblPer11 / 100) / 360) * dblNet1
        If oCombobox.Selected.Value = "Y" Then
            dblInsurance1 = dblTotal11 * dblInsurance1 / 100
        Else
            dblInsurance1 = 0
        End If
        oApplication.Utilities.setEdittextvalue(oForm, "40", dblNet11)
        oApplication.Utilities.setEdittextvalue(oForm, "73", dblInsurance1)
    End Sub
    Private Sub CalculateTotal(ByVal aForm As SAPbouiCOM.Form)
        Try
            Dim dblTotal, dblPercentage, dblNet, dblCurrency As Double
            Dim dblTotal1, dblNet1, dblPer1 As Double
            Dim dblFinal As Decimal
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            aForm.Freeze(True)

            dblTotal = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "32"))
            oGrid = oForm.Items.Item("33").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                dblPercentage = oGrid.DataTable.GetValue("Total", intRow)
                If dblPer1 = 0 And dblTotal = 0 Then
                    dblNet = 0
                Else
                    dblNet = dblPercentage / dblTotal ' / dblPercentage ' dblTotal * dblPercentage / 100
                End If


                oGrid.DataTable.SetValue("Price", intRow, dblNet)
            Next


            dblTotal = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "43"))
            dblPercentage = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "44"))
            dblNet = dblTotal * dblPercentage / 100
            oApplication.Utilities.setEdittextvalue(oForm, "17", dblNet)
            dblTotal = dblTotal + dblNet
            oApplication.Utilities.setEdittextvalue(oForm, "19", dblTotal)
            dblTotal = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "21"))

            dblPercentage = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "23"))
            dblNet = dblTotal * dblPercentage / 100
            oApplication.Utilities.setEdittextvalue(oForm, "24", dblNet)


            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTest.DoQuery("Select * from OITM where ItemCode='" & oApplication.Utilities.getEdittextvalue(oForm, "12") & "'")
            oApplication.Utilities.setEdittextvalue(oForm, "28", oTest.Fields.Item("U_Z_MT").Value)

            oTest.DoQuery("Select isnull(Rate ,0) from ORTT where Currency='USD' and RateDate ='" & Now.Date.ToString("yyyy-MM-dd") & "'")
            dblCurrency = oTest.Fields.Item(0).Value
            dblPercentage = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "28"))

            dblTotal = dblPercentage * dblCurrency * (dblTotal + dblNet + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "26")) + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "edFOH")))

            'dblTotal = dblPercentage * (dblTotal + dblNet + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "26")))
            oApplication.Utilities.setEdittextvalue(oForm, "30", dblTotal)
            oGrid = aForm.Items.Item("33").Specific
            dblTotal = 0
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                dblTotal = dblTotal + oGrid.DataTable.GetValue("Price", intRow)
            Next
            dblTotal = dblTotal + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "48"))
            dblTotal = dblTotal + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "50"))
            dblTotal = dblTotal + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "53"))
            oApplication.Utilities.setEdittextvalue(oForm, "46", dblTotal)
            dblTotal = dblTotal + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "19"))
            dblTotal = dblTotal + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "30"))
            'dblTotal = dblTotal + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "48"))
            oApplication.Utilities.setEdittextvalue(aForm, "55", dblTotal)

            dblTotal1 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "55"))
            dblNet = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "57"))

            oRec.DoQuery("select isnull(U_Z_CreditPer,0),isnull(U_Z_Insurance,0) from OADM")
            Dim dblInsurance As Double
            dblPer1 = oRec.Fields.Item(0).Value
            dblInsurance = oRec.Fields.Item(1).Value
            dblNet1 = 0
            dblNet1 = ((dblTotal1 * dblPer1 / 100) / 360) * dblNet
            oCombobox = aForm.Items.Item("cmbInsuran").Specific
            If oCombobox.Selected.Value = "Y" Then
                dblInsurance = dblTotal1 * dblInsurance / 100
            Else
                dblInsurance = 0
            End If
            oApplication.Utilities.setEdittextvalue(oForm, "40", dblNet1)
            oApplication.Utilities.setEdittextvalue(oForm, "73", dblInsurance)

            dblTotal = dblTotal + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "35"))
            dblTotal = dblTotal + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "38"))
            dblTotal = dblTotal + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "40"))
            dblTotal = dblTotal + oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "73"))
            oApplication.Utilities.setEdittextvalue(aForm, "59", dblTotal)
            dblFinal = Convert.ToDecimal(dblTotal)
            dblFinal = Math.Round(dblFinal, 2)
            Dim strFinal As String
            ' strFinal = dblFinal.ToString(".000")
            strFinal = dblFinal.ToString(".00")
            ' MsgBox(Math.Ceiling(dblFinal))
            strFinal = strFinal.Substring(strFinal.Length - 1)
            Dim intDiff As Integer
            If CInt(strFinal) <= 5 Then
                intDiff = 5 - CInt(strFinal)
                strFinal = "0" & CompanyDecimalSeprator & "0" & intDiff.ToString
                Dim dbvalue As Double
                dbvalue = oApplication.Utilities.getDocumentQuantity(strFinal)
                dblFinal = dblFinal + dbvalue
            Else
                intDiff = 10 - CInt(strFinal)
                strFinal = "0" & CompanyDecimalSeprator & "0" & intDiff.ToString
                Dim dbvalue As Double
                dbvalue = oApplication.Utilities.getDocumentQuantity(strFinal)
                dblFinal = dblFinal + dbvalue
                ' dblFinal = dblFinal + 0.1

            End If

            oApplication.Utilities.setEdittextvalue(aForm, "61", dblFinal)

            Dim C2, B3, C3, C4, C5, C6, B5, B6 As Double
            C2 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "edTin"))
            B3 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "edTraTot"))
            C3 = (C2 * B3) / (100 + B3)
            oApplication.Utilities.setEdittextvalue(aForm, "edTraPr", Math.Round(C3, 2))
            C4 = C2 - C3
            oApplication.Utilities.setEdittextvalue(aForm, "76", Math.Round(C4, 2))
            C4 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "76"))
            B5 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "edRSPTot"))
            C5 = (C4 * B5 / 100)
            oApplication.Utilities.setEdittextvalue(aForm, "edRSPPr", Math.Round(C5, 2))
            C5 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "edRSPPr"))
            C6 = C4 - C5
            oApplication.Utilities.setEdittextvalue(aForm, "78", Math.Round(C6, 2))
            B6 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "80"))
            C6 = (C4 * (B6 / 100))
            oApplication.Utilities.setEdittextvalue(aForm, "81", Math.Round(C6, 2))



            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub

    Private Sub CalculateReferenceValues(ByVal aform As SAPbouiCOM.Form)
        Dim C2, B3, C3, C4, C5, C6, B5, B6 As Double
        Try
            aform.Freeze(True)
            C2 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "edTin"))
            B3 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "edTraTot"))
            C3 = (C2 * B3) / (100 + B3)
            oApplication.Utilities.setEdittextvalue(aform, "edTraPr", Math.Round(C3, 2))
            C4 = C2 - C3
            oApplication.Utilities.setEdittextvalue(aform, "76", Math.Round(C4, 2))
            C4 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "76"))
            B5 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "edRSPTot"))
            C5 = (C4 * B5 / 100)
            oApplication.Utilities.setEdittextvalue(aform, "edRSPPr", Math.Round(C5, 2))
            C5 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "edRSPPr"))
            C6 = C4 - C5
            oApplication.Utilities.setEdittextvalue(aform, "78", Math.Round(C6, 2))
            B6 = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "80"))
            C6 = (C4 * (B6 / 100))
            oApplication.Utilities.setEdittextvalue(aform, "81", Math.Round(C6, 2))
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)

        End Try

    End Sub
#End Region
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_CostSheet Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = 2 Then
                                    Dim otest As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    otest.DoQuery("Delete from [@Z_AL_COST] where Name='" & oApplication.Utilities.getEdittextvalue(oForm, "10") & "' and Name like '%N'")
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID

                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.CharPressed = 9 Then
                                    Select Case pVal.ItemUID

                                        Case "44"
                                            Dim dblTotal, dblPercentage, dblNet As Double
                                            ' CalculateTotal(oForm)

                                        Case "48", "50"
                                            CalculateTotal(oForm)
                                    End Select
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        If AddtoUDT(oForm) = True Then
                                            oMatrix = frmSourceForm.Items.Item("38").Specific
                                            Dim aCode, strCardCode, StrDocCurrency, strdate As String
                                            Dim dblCurrency, dblCostingAmt As Double
                                            Dim dtPostingDate As Date
                                            Dim otest As SAPbobsCOM.Recordset
                                            aCode = oApplication.Utilities.getEdittextvalue(oForm, "10")
                                            aCode = aCode.Replace("N", "")
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Cost_Ref", intSelectedMatrixrow, aCode)
                                            oForm.Close()
                                        End If
                                    Case "62"
                                        CalculateTotal(oForm)
                                End Select


                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "Mapping"
                    If pVal.BeforeAction = False Then
                        Dim aCode As String
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    End If
                Case "Lead"
                    If pVal.BeforeAction = False Then
                        Dim aCode As String
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
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
        If oForm.TypeEx = frm_BPMaster Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                        Try
                            oApplication.SBO_Application.Menus.RemoveEx("Mapping")
                        Catch ex As Exception
                        End Try
                        Try
                            oApplication.SBO_Application.Menus.RemoveEx("Lead")
                        Catch ex As Exception
                        End Try

                        Try
                            oApplication.SBO_Application.Menus.RemoveEx("LeadTime")
                        Catch ex As Exception
                        End Try



                        oCombobox = oForm.Items.Item("40").Specific
                        If oCombobox.Selected.Value = "C" Then
                            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                            Try
                                oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = "Mapping"
                                oCreationPackage.String = "Rebate Percentage Mapping"
                                oCreationPackage.Enabled = True
                                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                oMenus.AddEx(oCreationPackage)
                            Catch ex As Exception

                            End Try
                        ElseIf oCombobox.Selected.Value = "S" Then
                            Try
                                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                                oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = "Lead"
                                oCreationPackage.String = "Lead Time Mapping"
                                oCreationPackage.Enabled = True
                                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                oMenus.AddEx(oCreationPackage)
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Else
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Try
                            oApplication.SBO_Application.Menus.RemoveEx("Mapping")
                            oApplication.SBO_Application.Menus.RemoveEx("Lead")
                        Catch ex As Exception

                        End Try

                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
