Public Class clsDocumentPrint
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oCheckBoxColumn As SAPbouiCOM.CheckBoxColumn
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
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_ALPrint, frm_ALPrint)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataSources.UserDataSources.Add("Doc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("RptView", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("frmDoc", SAPbouiCOM.BoDataType.dt_DATE)
        oForm.DataSources.UserDataSources.Add("ToDoc", SAPbouiCOM.BoDataType.dt_DATE)
        oCombobox = oForm.Items.Item("16").Specific
        oCombobox.ValidValues.Add("W", "Window")
        oCombobox.ValidValues.Add("P", "PDF")
        oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        oEditText = oForm.Items.Item("9").Specific
        oEditText.DataBind.SetBound(True, "", "frmDoc")
        oEditText = oForm.Items.Item("11").Specific
        oEditText.DataBind.SetBound(True, "", "ToDoc")
        oCombobox = oForm.Items.Item("7").Specific
        oCombobox.DataBind.SetBound(True, "", "Doc")
        oCombobox.ValidValues.Add("", "")
        oCombobox.ValidValues.Add("Inv", "Invoice")
        oCombobox.ValidValues.Add("Del", "Delivery")
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oForm.Items.Item("7").DisplayDesc = True
        oGrid = oForm.Items.Item("12").Specific
        oGrid.DataTable.ExecuteQuery("Select DocEntry,DocNum,CardCode,CardName,DocTotal, ' ' 'Select' from OINV where 1=2")
        oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oGrid.AutoResizeColumns()
        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub
#Region "Bind Choose From List"
    Private Sub binChooseFromList(ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        oCombobox = aform.Items.Item("7").Specific
        If oCombobox.Selected.Value = "Inv" Then
            'oEditText = aform.Items.Item("9").Specific
            'oEditText.ChooseFromListUID = "CFL_2"
            'oEditText.ChooseFromListAlias = "DocNum"
            'oEditText.String = ""
            'oEditText = aform.Items.Item("11").Specific
            'oEditText.ChooseFromListUID = "CFL_3"
            'oEditText.ChooseFromListAlias = "DocNum"
            'oEditText.String = ""
            aform.Title = "Document Printing - Invoices"
        Else
            'oEditText = aform.Items.Item("9").Specific
            'oEditText.ChooseFromListUID = "CFL_4"
            'oEditText.ChooseFromListAlias = "DocNum"
            'oEditText.String = ""
            'oEditText = aform.Items.Item("11").Specific
            'oEditText.ChooseFromListUID = "CFL_5"
            'oEditText.ChooseFromListAlias = "DocNum"
            'oEditText.String = ""
            aform.Title = "Document Printing - Delivery"
        End If
        aform.Freeze(False)
    End Sub

    Private Sub BindData(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strDoctype, strFromDoc, strtoDoc, strsql, strQuery As String
            oCombobox = aform.Items.Item("7").Specific
            strDoctype = oCombobox.Selected.Value
            strFromDoc = oApplication.Utilities.getEdittextvalue(aform, "9")
            strtoDoc = oApplication.Utilities.getEdittextvalue(aform, "11")
            Dim dtFrom, dtTo As Date
            If strFromDoc <> "" Then
                dtFrom = oApplication.Utilities.GetDateTimeValue(strFromDoc)
                strsql = " T0.DocDate >='" & dtFrom.ToString("yyyy-MM-dd") & "'"
            Else
                strsql = " 1 = 1"
            End If
            If strtoDoc <> "" Then
                dtTo = oApplication.Utilities.GetDateTimeValue(strtoDoc)
                strsql = strsql & " and T0.DocDate<='" & dtTo.ToString("yyyy-MM-dd") & "'"
            Else
                strsql = strsql & " and 1=1"
            End If
            If strDoctype = "Inv" Then
                strQuery = "Select T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName,T1.City,T0.DocCur,T0.DocTotalFC,T0.DocTotal,' ' 'Select' from OINV  T0 inner join OCRD T1 on T1.CardCode=T0.CardCode where " & strsql & " order by T1.City"
            Else
                strQuery = "Select T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName,T1.City,T0.DocCur,T0.DocTotalFC,T0.DocTotal,' ' 'Select' from ODLN T0   inner join OCRD T1 on T1.CardCode=T0.CardCode where " & strsql & " order by T1.City"
            End If
            oGrid = aform.Items.Item("12").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)
            oGrid.Columns.Item(0).Editable = False
            oEditTextColumn = oGrid.Columns.Item(0)
            If strDoctype = "Inv" Then
                oEditTextColumn.LinkedObjectType = "13"
            Else
                oEditTextColumn.LinkedObjectType = "15"
            End If
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("Select").Editable = True
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)

        End Try
    End Sub
    Private Sub SelectAll(ByVal aForm As SAPbouiCOM.Form, ByVal aflag As Boolean)
        aForm.Freeze(True)
        oGrid = aForm.Items.Item("12").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCheckBoxColumn = oGrid.Columns.Item("Select")
            oCheckBoxColumn.Check(intRow, aflag)
        Next
        aForm.Freeze(False)
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ALPrint Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "4"
                                        If oForm.PaneLevel = 1 Then
                                            oCombobox = oForm.Items.Item("7").Specific
                                            If oCombobox.Selected.Value = "" Then
                                                oApplication.Utilities.Message("Document Type is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If

                                End Select
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                'oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                Select Case pVal.ItemUID
                                    Case "3"
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                    Case "4"
                                        If oForm.PaneLevel = 2 Then
                                            oCombobox = oForm.Items.Item("7").Specific
                                            If oCombobox.Selected.Value = "" Then
                                                oApplication.Utilities.Message("Document Type is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                oForm.PaneLevel = 2
                                                Exit Sub
                                            End If
                                        End If
                                        If oForm.PaneLevel = 1 Then
                                            binChooseFromList(oForm)
                                            oForm.PaneLevel = oForm.PaneLevel + 1
                                        Else
                                            BindData(oForm)
                                            oForm.PaneLevel = oForm.PaneLevel + 1
                                        End If
                                    Case "5"
                                        If oApplication.SBO_Application.MessageBox("Do you want to print the selected Documents?", , "Yes", "No") = 2 Then
                                            Exit Sub
                                        End If
                                        Dim oObj As New clsPrint
                                        Dim strPrintOption As String
                                        oCombobox = oForm.Items.Item("16").Specific
                                        strPrintOption = oCombobox.Selected.Value
                                        oCombobox = oForm.Items.Item("7").Specific
                                        oGrid = oForm.Items.Item("12").Specific
                                        oObj.PrintInvoice(oGrid, oCombobox.Selected.Value, strPrintOption)
                                    Case "13"
                                        SelectAll(oForm, True)
                                    Case "14"
                                        SelectAll(oForm, False)

                                End Select

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val, val2 As String
                                Dim intChoice As Integer
                                Dim codebar As String
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
                                        If pVal.ItemUID = "9" Or pVal.ItemUID = "11" Then
                                            val = oDataTable.GetValue("DocNum", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
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
                Case mnu_Print
                    loadform()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

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
