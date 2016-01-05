Public Class clsGoodsReceipt
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
        Dim oTest, oTest1, otest2 As SAPbobsCOM.Recordset
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode, aDocNum, LineRef As String
        Try
            aform.Freeze(True)
            Dim oRec, oRec1, oRec2 As SAPbobsCOM.Recordset
            Dim oItem As SAPbobsCOM.Items
            Dim strItemcode, strRefCode As String
            aDocNum = oApplication.Utilities.getEdittextvalue(aform, "8")

            oApplication.Utilities.LoadForm(xml_GRREceipt, frm_GRReceipt)
            Dim objForm As SAPbouiCOM.Form
            objForm = oApplication.SBO_Application.Forms.ActiveForm()
            If objForm.TypeEx = frm_GRReceipt Then
                Try
                    objForm.Freeze(True)
                    Dim ogrid As SAPbouiCOM.Grid
                    oApplication.Utilities.setEdittextvalue(objForm, "5", aDocNum)
                    oApplication.Utilities.setEdittextvalue(objForm, "7", oApplication.Utilities.getEdittextvalue(aform, "4"))
                    ogrid = objForm.Items.Item("4").Specific
                    objForm.Items.Item("1").Enabled = True
                    Dim strstring As String
                    ' strstring = "SELECT T1.DocEntry,T1.[LineNum], T1.[ItemCode], T1.[Dscription], T1.[Quantity],T1.[Currency], T1.[Price], T1.[WhsCode],T1.U_IS 'IS',T1.U_IA 'IA',T1.U_RVD 'RVD',T1.U_MD 'MD' FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry WHERE  T0.DocEntry=" & aDocNum
                    strstring = " SELECT T1.DocEntry,T1.[LineNum], T1.[ItemCode], T1.[Dscription], T1.[Quantity],T1.[Currency], T1.[Price], T1.[WhsCode], T1.U_IS 'IS',U_MN 'MN',U_RVD 'RVD',U_MD  'MD',isnull(U_IA,'Pending') 'IA',U_MR 'MR',U_Desc 'Desc' FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry WHERE  T0.DocNum=" & aDocNum
                    'strstring = strstring & " where Code in (" & LineRef & ")"
                    ogrid.DataTable.ExecuteQuery(strstring)
                    'ogrid.DataTable.ExecuteQuery("SElect * from [@Z_AL_COST] where U_Z_DocNum='" & aDocNum & "' and Code in (" & LineRef & ")")
                    ogrid.Columns.Item("DocEntry").TitleObject.Caption = "Document Number"
                    ogrid.Columns.Item("LineNum").TitleObject.Caption = "Line Number"
                    ogrid.Columns.Item("ItemCode").TitleObject.Caption = "Item Code"
                    ogrid.Columns.Item("Dscription").TitleObject.Caption = "Item Name"
                    ogrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
                    ogrid.Columns.Item("Price").TitleObject.Caption = "Price"
                    ogrid.Columns.Item("Currency").TitleObject.Caption = "Currency"

                    ogrid.Columns.Item("WhsCode").TitleObject.Caption = "Warehouse"
                    ogrid.Columns.Item("IS").TitleObject.Caption = "Item Status"
                    ogrid.Columns.Item("IA").TitleObject.Caption = "Approval"
                    ogrid.Columns.Item("RVD").TitleObject.Caption = "Received Date"
                    ogrid.Columns.Item("MD").TitleObject.Caption = "Municipality Date"
                    ogrid.Columns.Item("MN").TitleObject.Caption = "Municipality Number"
                    ogrid.Columns.Item("Desc").TitleObject.Caption = "Description"
                    ogrid.Columns.Item("MR").TitleObject.Caption = "Municipality Rejection"
                    'ogrid.Columns.Item("MR").TitleObject.Caption = "Municipality Rejection"
                    'ogrid.Columns.Item("MR").TitleObject.Caption = "Municipality Rejection"


                    For introw As Integer = 0 To 7
                        ogrid.Columns.Item(introw).Editable = False
                    Next
                    Dim oEditTextColumn As SAPbouiCOM.EditTextColumn = ogrid.Columns.Item("ItemCode")
                    oEditTextColumn.LinkedObjectType = "4"
                    ogrid.Columns.Item("IA").Editable = True
                    ogrid.Columns.Item("IA").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                    Dim oComboColumn As SAPbouiCOM.ComboBoxColumn
                    oComboColumn = ogrid.Columns.Item("IA")
                    oComboColumn.ValidValues.Add("", "")
                    oComboColumn.ValidValues.Add("Yes", "Yes")
                    oComboColumn.ValidValues.Add("No", "No")
                    oComboColumn.ValidValues.Add("Pending", "Pending")
                    oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                    ogrid.Columns.Item("RVD").Editable = True
                    ogrid.Columns.Item("MD").Editable = True
                    ogrid.Columns.Item("MR").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                    oComboColumn = ogrid.Columns.Item("MR")
                    oComboColumn.ValidValues.Add("", "")
                    oComboColumn.ValidValues.Add("LI", "Lab Issue")
                    oComboColumn.ValidValues.Add("DI", "Documentation Issue")
                    oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

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
            If pVal.FormTypeEx = frm_GoodsReceipt Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                               
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oApplication.Utilities.AddControls(oForm, "btnView", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", , , "2", "Municipality Follow up", 120)
                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "btnView" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    Populate(oForm)
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
                'Case mnu_DuplicateRow
                '    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                '    If pVal.BeforeAction = False Then
                '        If oForm.TypeEx = frm_PurchaseQuatation Then
                '            oMatrix = oForm.Items.Item("38").Specific
                '            If intCurrentRow <> 10000 Then
                '                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Cost_Ref", intCurrentRow + 1, "")
                '            End If
                '        End If
                '    Else
                '        If oForm.TypeEx = frm_PurchaseQuatation Then
                '            'oMatrix = oForm.Items.Item("38").Specific
                '            'If intCurrentRow <> 10000 Then
                '            '    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Ref", intCurrentRow + 1, "")
                '            'End If
                '        End If
                '    End If

                'Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
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
            'If oForm.TypeEx = frm_PurchaseQuatation And eventInfo.ItemUID = "38" Then
            '    oMatrix = oForm.Items.Item("38").Specific
            '    intCurrentRow = eventInfo.Row
            'End If
        End If

    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                'If BusinessObjectInfo.FormTypeEx = frm_PurchaseQuatation Then
                '    Dim oDoc As SAPbobsCOM.Documents
                '    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations)
                '    If oDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                '        Dim orec As SAPbobsCOM.Recordset
                '        Dim strCode As String
                '        orec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                '        'If BusinessObjectInfo.Type <> "112" Then
                '        '    For intRow As Integer = 0 To oDoc.Lines.Count - 1
                '        '        oDoc.Lines.SetCurrentLine(intRow)
                '        '        strCode = oDoc.Lines.UserFields.Fields.Item("U_Z_Cost_Ref").Value
                '        '        If strCode <> "" Then
                '        '            orec.DoQuery("update [@Z_AL_COST] set U_Z_LineId=" & oDoc.Lines.LineNum & ", U_Z_DocEntry=" & oDoc.DocEntry & ", U_Z_DocNum=" & oDoc.DocNum & " where Code='" & strCode & "'")
                '        '        End If
                '        '    Next
                '        '    'orec.DoQuery("Delete from [@Z_COST] where U_Z_DocNum=" & oDoc.DocNum & " and isnull(U_Z_DocEntry,9999)=9999")

                '    Else

                '    End If
                '  End If



                '  End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
