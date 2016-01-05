Public Class clsSalesOrder
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
    Private oBP As SAPbobsCOM.BusinessPartners
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SalesOrder Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    Dim otest, oTest1 As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTest1.DoQuery("Select isnull(U_Z_CardCode,''),isnull(U_Z_BranchDB,'') from OCRD where CardCode='" & oApplication.Utilities.getEdittextvalue(oForm, "4") & "'")
                                    If oTest1.Fields.Item(0).Value <> "" And oTest1.Fields.Item(1).Value <> "" Then ' And oDoc.UserFields.Fields.Item("U_Z_Exported").Value <> "Y" Then
                                        otest.DoQuery("Select * from [@Z_AL_OADM] where U_Z_BraDB='" & oTest1.Fields.Item(1).Value & "'")
                                        Dim oRemCompany As SAPbobsCOM.Company
                                        oRemCompany = New SAPbobsCOM.Company
                                        If oApplication.Utilities.CheckConnection(otest.Fields.Item("U_Z_BraDB").Value, otest.Fields.Item("U_Z_SAPUID").Value, otest.Fields.Item("U_Z_SAPPWD").Value) = False Then
                                            oApplication.Utilities.Message("Check the Login setup ....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub

                                        End If
                                    End If
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED


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
                Case mnu_InvSO
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region


    Private Sub ExportSalesOrer(ByVal aDocEntry As Integer, ByVal aremoteCompany As SAPbobsCOM.Company)
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.Documents
        Dim strPath, strFilename, strMessage As String
        Dim strFileName1, strBranchWhs, strMainWhs, strSQL As String
        Dim objremoteRec, objMainRec As SAPbobsCOM.Recordset
        Dim oItem As SAPbobsCOM.Items
        objremoteRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select DocEntry from ODLN where DocEntry=" & aDocEntry & " and U_Z_Exported='N'")
        'Header fields – customer code, customer name, posting date, due date, document date, customer reference number, remarks, Total before discount, discount %, Tax, Total, Applied Amount, Balance Due
        'Line fields – item code, item name, barcode, quantity, unit price, discount %, price after discount, vat code, Gross price, Total (LC)
        For intRemoteLoop As Integer = 1 To 1 'objremoteRec.RecordCount - 1
            objremoteDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

            If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                objMainDoc = aremoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                objMainDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes
                '  objMainDoc.Address = objremoteDoc.Address
                ' objMainDoc.Address2 = objremoteDoc.Address2
                objMainRec.DoQuery("Select U_Z_CardCode ,Currency from OCRD where CardCode='" & objremoteDoc.CardCode & "'")
                objMainDoc.CardCode = objMainRec.Fields.Item(0).Value
                Dim strdoccur As String = objremoteDoc.DocCurrency
                Dim dblDocRate As Double
                Dim oTestRS As SAPbobsCOM.Recordset
                oTestRS = aremoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTestRS.DoQuery("Select * from ortt where Currency='" & strdoccur & "' and Ratedate='" & objMainDoc.DocDate.ToString("yyyy-MM-dd") & "'")
                If oTestRS.RecordCount > 0 Then
                    dblDocRate = oTestRS.Fields.Item("Rate").Value
                Else
                    dblDocRate = 1
                End If
                If aremoteCompany.GetCompanyService.GetAdminInfo.LocalCurrency = strdoccur Then
                    strdoccur = aremoteCompany.GetCompanyService.GetAdminInfo.LocalCurrency
                ElseIf aremoteCompany.GetCompanyService.GetAdminInfo.SystemCurrency = strdoccur Then
                    strdoccur = aremoteCompany.GetCompanyService.GetAdminInfo.SystemCurrency
                Else
                    strdoccur = strdoccur
                End If
                objMainDoc.DocCurrency = strdoccur 'objremoteDoc.DocCurrency
                objMainDoc.DocRate = dblDocRate
                objMainDoc.Comments = objremoteDoc.Comments
                objMainDoc.DocDate = objremoteDoc.DocDate
                objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                objMainDoc.DocType = objremoteDoc.DocType
                objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                objMainDoc.Comments = objremoteDoc.Comments
                objMainDoc.TaxDate = objremoteDoc.TaxDate
                Try
                    objMainDoc.UserFields.Fields.Item("U_DeliveryNo").Value = objremoteDoc.DocNum
                Catch ex As Exception

                End Try

                If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                    objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                Else
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                End If
                Try
                    objMainDoc.UserFields.Fields.Item("U_DG").Value = "No"
                Catch ex As Exception
                    objMainDoc.UserFields.Fields.Item("U_DG").Value = "NO"
                End Try

                'objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                'objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.DocEntry
                'objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = (objremoteDoc.DocNum)
                'objMainDoc.UserFields.Fields.Item("U_Branch").Value = oApplication.Company.CompanyName
                'For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                '    If objremoteDoc.Expenses.LineTotal > 0 Then
                '        If IntExp > 0 Then
                '            objMainDoc.Expenses.Add()
                '            objMainDoc.Expenses.SetCurrentLine(IntExp)
                '        End If
                '        objremoteDoc.Expenses.SetCurrentLine(IntExp)
                '        objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                '        objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                '        objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                '        objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                '        objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                '        objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                '        objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                '        '  objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                '        objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                '        'objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                '        'objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                '    End If
                'Next
                For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                    If intLoop > 0 Then
                        objMainDoc.Lines.Add()
                        objMainDoc.Lines.SetCurrentLine(intLoop)
                    End If
                    objremoteDoc.Lines.SetCurrentLine(intLoop)
                    Dim currency As String = objremoteDoc.Lines.Currency
                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                    objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                    objMainDoc.Lines.Currency = strdoccur
                    Dim oRec1 As SAPbobsCOM.Recordset
                    oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec1.DoQuery("Select isnull(AvgPrice,0) from OITW where Itemcode='" & objremoteDoc.Lines.ItemCode & "' and whscode='" & objremoteDoc.Lines.WarehouseCode & "'")
                    If oRec1.RecordCount > 0 Then
                        objMainDoc.Lines.UnitPrice = oRec1.Fields.Item(0).Value
                    Else
                        If oItem.GetByKey(objremoteDoc.Lines.ItemCode) Then
                            objMainDoc.Lines.UnitPrice = oItem.AvgStdPrice '* dblDocRate
                        Else
                            objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice '* dblDocRate
                        End If
                    End If
                   
                    objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                    strBranchWhs = objremoteDoc.Lines.WarehouseCode
                    objMainDoc.Lines.CostingCode = objremoteDoc.Lines.CostingCode
                    objMainDoc.Lines.CostingCode2 = objremoteDoc.Lines.CostingCode2
                    objMainDoc.Lines.CostingCode3 = objremoteDoc.Lines.CostingCode3
                    objMainDoc.Lines.CostingCode4 = objremoteDoc.Lines.CostingCode4
                    objMainDoc.Lines.CostingCode5 = objremoteDoc.Lines.CostingCode5
                    ' objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                    If objremoteDoc.Lines.WarehouseCode = "MW" Then
                        objMainDoc.Lines.WarehouseCode = "MPW"
                    Else
                        objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                    End If


                    For intSer As Integer = 0 To objremoteDoc.Lines.SerialNumbers.Count - 1
                        If intSer > 0 Then
                            objMainDoc.Lines.SerialNumbers.Add()
                            objMainDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                        End If
                        objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                        objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                        objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                        objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                        objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                        objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                        objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                        objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate
                        ' objMainDoc.Lines.SerialNumbers.SystemSerialNumber = objremoteDoc.Lines.SerialNumbers.SystemSerialNumber
                    Next

                    For intSer As Integer = 0 To objremoteDoc.Lines.BatchNumbers.Count - 1
                        If intSer > 0 Then
                            objMainDoc.Lines.BatchNumbers.Add()
                            objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                        End If
                        objremoteDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                        objMainDoc.Lines.BatchNumbers.AddmisionDate = objremoteDoc.Lines.BatchNumbers.AddmisionDate
                        objMainDoc.Lines.BatchNumbers.BaseLineNumber = objremoteDoc.Lines.BatchNumbers.BaseLineNumber
                        objMainDoc.Lines.BatchNumbers.BatchNumber = objremoteDoc.Lines.BatchNumbers.BatchNumber
                        objMainDoc.Lines.BatchNumbers.ExpiryDate = objremoteDoc.Lines.BatchNumbers.ExpiryDate
                        objMainDoc.Lines.BatchNumbers.InternalSerialNumber = objremoteDoc.Lines.BatchNumbers.InternalSerialNumber
                        objMainDoc.Lines.BatchNumbers.Location = objremoteDoc.Lines.BatchNumbers.Location
                        objMainDoc.Lines.BatchNumbers.ManufacturingDate = objremoteDoc.Lines.BatchNumbers.ManufacturingDate
                        objMainDoc.Lines.BatchNumbers.Notes = objremoteDoc.Lines.BatchNumbers.Notes
                        objMainDoc.Lines.BatchNumbers.Quantity = objremoteDoc.Lines.BatchNumbers.Quantity
                    Next
                Next
                'objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                If objMainDoc.Add <> 0 Then
                    oApplication.Utilities.Message("Failed to create invoice docuemnt :" & objremoteDoc.DocNum & " Error : " & aremoteCompany.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    Dim strDocNum As String
                    aremoteCompany.GetNewObjectCode(strDocNum)
                    oApplication.Utilities.Message("GRPO Created Successfully.Draft Number : " & strDocNum, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objremoteDoc.UserFields.Fields.Item("U_Z_Exported").Value = "Y"
                    objremoteDoc.Update()
                End If
            End If
            objremoteRec.MoveNext()
        Next
    End Sub
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_Delivery Then
                    Dim oDoc As SAPbobsCOM.Documents
                    Dim oDoc1 As SAPbobsCOM.Documents
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    oDoc1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                    If oDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        Dim otest, oTest1 As SAPbobsCOM.Recordset
                        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTest1.DoQuery("Select isnull(U_Z_CardCode,''),isnull(U_Z_BranchDB,'') from OCRD where CardCode='" & oDoc.CardCode & "'")
                        Dim strExported As String = oDoc.UserFields.Fields.Item("U_Z_Exported").Value
                        If oTest1.Fields.Item(0).Value <> "" And oTest1.Fields.Item(1).Value <> "" And oDoc.UserFields.Fields.Item("U_Z_Exported").Value <> "Y" Then
                            otest.DoQuery("Select * from [@Z_AL_OADM] where U_Z_BraDB='" & oTest1.Fields.Item(1).Value & "'")
                            Dim oRemCompany As SAPbobsCOM.Company
                            oRemCompany = New SAPbobsCOM.Company
                            oRemCompany = oApplication.Utilities.ConnectRemoteCompany(otest.Fields.Item("U_Z_BraDB").Value, otest.Fields.Item("U_Z_SAPUID").Value, otest.Fields.Item("U_Z_SAPPWD").Value)
                            ExportSalesOrer(oDoc.DocEntry, oRemCompany)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
End Class
