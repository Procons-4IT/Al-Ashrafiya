Public Class clsSalesTarget
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
    Private oColumn As SAPbouiCOM.Column
    Private InvBase As DocumentType
    Private MatrixId As String
    Private RowtoDelete As Integer
    Private InvBaseDocNo As String
    Private InvForConsumedItems, count As Integer
    Private blnFlag As Boolean = False
    Dim oDataSrc_Line As SAPbouiCOM.DBDataSource

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_SalesTarget, frm_SalesTarget)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "4"
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        '  AddChooseFromList(oForm)
        'databind(oForm)
        oMatrix = oForm.Items.Item("16").Specific
        oColumn = oMatrix.Columns.Item("SlpCode")
        oColumn.ChooseFromListUID = "CFL1"
        oColumn.ChooseFromListAlias = "SlpCode"
        oCombobox = oForm.Items.Item("11").Specific
        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow)
        Next
        oCombobox.ValidValues.Add("", "")
        For intRow As Integer = 2010 To 2050
            oCombobox.ValidValues.Add(intRow, intRow)
        Next
        oForm.Items.Item("11").DisplayDesc = True
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_AL_SAL1")
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_AL_SAL2")
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oForm.PaneLevel = 1
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oForm.Freeze(False)
    End Sub


#Region "Add Choose From List"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        oMatrix = aForm.Items.Item("7").Specific
        oColumn = oMatrix.Columns.Item("V_0")
        oColumn.ChooseFromListUID = "CFL1"
        oColumn.ChooseFromListAlias = "FormatCode"
        oColumn = oMatrix.Columns.Item("V_1")
        oColumn.ChooseFromListUID = "CFL2"
        oColumn.ChooseFromListAlias = "FormatCode"

        oColumn = oMatrix.Columns.Item("V_3")
        oColumn.ChooseFromListUID = "CFL3"
        oColumn.ChooseFromListAlias = "FormatCode"

        oColumn = oMatrix.Columns.Item("V_4")
        oColumn.ChooseFromListUID = "CFL4"
        oColumn.ChooseFromListAlias = "FormatCode"

        Dim oColum As SAPbouiCOM.Column
        oColum = oMatrix.Columns.Item("V_2")
        Try
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetCompanyList
            For intRow As Integer = 0 To otemp.RecordCount - 1
                oColum.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(0).Value)
                otemp.MoveNext()
            Next
        Catch ex As Exception
        End Try
        oColum.DisplayDesc = True
    End Sub
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFL = oCFLs.Item("CFL1")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL2")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL3")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL4")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region
#Region "Methods"
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("16").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_AL_SAL1")
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            oMatrix = aForm.Items.Item("17").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_AL_SAL2")
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub

#End Region
#Region "Validations"
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strcode, stCode1, stCode As String
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                AddMode(aForm)
            End If
            'If oApplication.Utilities.getEdittextvalue(aForm, "4") = "" Then
            '    oApplication.Utilities.Message("Code is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            'If oApplication.Utilities.getEdittextvalue(aForm, "6") = "" Then
            '    oApplication.Utilities.Message("Description is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                'Dim oTemp As SAPbobsCOM.Recordset
                'oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oTemp.DoQuery("Select * from [@Z_HR_OTRAPL]  where U_Z_TraCode='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' ")
                'If oTemp.RecordCount > 0 Then
                '    oApplication.Utilities.Message("Travel Code is already mapped...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
                'End If
            End If
            oMatrix = aForm.Items.Item("16").Specific
            If oMatrix.RowCount = 0 Then
                oApplication.Utilities.Message("Line Details missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Dim strSlpCode, strSlpCode1 As String
            For intRow As Integer = 1 To oMatrix.RowCount
                strSlpCode = oApplication.Utilities.getMatrixValues(oMatrix, "SlpCode", intRow)
                If strSlpCode <> "" Then
                    For intLoop As Integer = intRow + 1 To oMatrix.RowCount
                        strSlpCode1 = oApplication.Utilities.getMatrixValues(oMatrix, "SlpCode", intLoop)
                        If strSlpCode1 <> "" Then
                            If strSlpCode = strSlpCode1 Then
                                oApplication.Utilities.Message("Sales person Details already exists : " & strSlpCode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If
                    Next
                End If

            Next

            AssignLineNo(oForm)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case 1
                    oMatrix = aForm.Items.Item("16").Specific
                    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_AL_SAL1")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    End If
                    If oMatrix.Columns.Item("SlpCode").Cells.Item(oMatrix.RowCount).Specific.value <> "" Then
                        oMatrix.AddRow()
                        oMatrix.ClearRowData(oMatrix.RowCount)
                    End If
                Case 2
                    oMatrix = aForm.Items.Item("17").Specific
                    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_AL_SAL2")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    End If
                    If oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.value <> "" Then
                        oMatrix.AddRow()
                        oMatrix.ClearRowData(oMatrix.RowCount)
                    End If
            End Select
            
            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
            End If

            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            AssignLineNo(aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)

        End Try
    End Sub
#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)
        If Me.MatrixId = "16" Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_AL_SAL1")
        End If
        If Me.MatrixId = "17" Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_AL_SAL2")
        End If
        'oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
        If intSelectedMatrixrow <= 0 Then
            Exit Sub

        End If
        Me.RowtoDelete = intSelectedMatrixrow
        oDataSrc_Line.RemoveRecord(Me.RowtoDelete - 1)
        oMatrix = frmSourceMatrix
        oMatrix.FlushToDataSource()
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        If oMatrix.RowCount > 0 Then
            oMatrix.DeleteRow(oMatrix.RowCount)
        End If
        aForm.Freeze(False)
    End Sub
#End Region

    Private Sub AddMode(ByVal aform As SAPbouiCOM.Form)
        Dim strCode As String
        strCode = oApplication.Utilities.getMaxCode("@Z_AL_OSAL", "DocEntry")
        aform.Items.Item("4").Enabled = True
        aform.Items.Item("6").Enabled = True
        oApplication.Utilities.setEdittextvalue(aform, "4", strCode)
        aform.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oApplication.Utilities.setEdittextvalue(aform, "6", "t")
        oApplication.SBO_Application.SendKeys("{TAB}")
        ' aform.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        aform.Items.Item("4").Enabled = False
        aform.Items.Item("6").Enabled = False
    End Sub


#Region "Create Journal Voucher"
    Public Function getSAPAccount(ByVal aCode As String, ByVal aCompany As SAPbobsCOM.Company) As String
        Dim oRS As SAPbobsCOM.Recordset
        oRS = aCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery("Select isnull(AcctCode,'') from OACT where Formatcode='" & aCode & "'")
        Dim stcode As String = oRS.Fields.Item(0).Value
        Return oRS.Fields.Item(0).Value
    End Function
    Private Function CreateJournalVoucher(ByVal aDocEntry As String) As Boolean
        Try
            Dim oJV, oJV1 As SAPbobsCOM.JournalVouchers
            Dim oRemCompany As SAPbobsCOM.Company
            Dim oRecSet, oRecSet1, oRecSet2, oRecset3 As SAPbobsCOM.Recordset
            Dim dtDate As Date
            Dim blnLineExists As Boolean = False
            oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSet1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSet2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecset3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If CreateJournalVoucher_Branch(aDocEntry) = True Then
                oRecSet.DoQuery("Select * from [@Z_AL_OCASH] where isnull(U_Z_Status,'O')='O' and DocEntry=" & CDbl(aDocEntry))
                If oRecSet.Fields.Item("U_Z_JVNo").Value = "" Then
                    dtDate = oRecSet.Fields.Item("CreateDate").Value
                    oRecSet.DoQuery("Select * from [@Z_AL_CASH1] where DocEntry=" & CDbl(aDocEntry))
                    oJV = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)
                    oJV.JournalEntries.DueDate = dtDate
                    Dim intCount As Integer = 0
                    For introw As Integer = 0 To oRecSet.RecordCount - 1
                        If oRecSet.Fields.Item("U_Z_Amount").Value > 0 Then
                            If intCount > 0 Then
                                oJV.JournalEntries.Lines.Add()
                            End If
                            oJV.JournalEntries.Lines.SetCurrentLine(intCount)
                            oJV.JournalEntries.Lines.AccountCode = getSAPAccount(oRecSet.Fields.Item("U_Z_FrmDAcc").Value, oApplication.Company)
                            oJV.JournalEntries.Lines.Debit = oRecSet.Fields.Item("U_Z_Amount").Value
                            blnLineExists = True
                            intCount = intCount + 1
                            If intCount > 0 Then
                                oJV.JournalEntries.Lines.Add()
                            End If
                            oJV.JournalEntries.Lines.SetCurrentLine(intCount)
                            oJV.JournalEntries.Lines.AccountCode = getSAPAccount(oRecSet.Fields.Item("U_Z_FrmCAcc").Value, oApplication.Company)
                            oJV.JournalEntries.Lines.Credit = oRecSet.Fields.Item("U_Z_Amount").Value
                            intCount = intCount + 1
                            blnLineExists = True
                        End If
                        oRecSet.MoveNext()
                    Next
                    If blnLineExists = True Then
                        If oJV.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        Else
                            Dim st As String
                            oApplication.Company.GetNewObjectCode(st)
                            oRecSet.DoQuery("Update [@Z_AL_OCASH] set U_Z_Status='C' where DocEntry=" & CDbl(aDocEntry))
                            '     MsgBox(st)
                            'oRecSet1.DoQuery("Select U_Z_Branch,Count(*) from [@Z_AL_CASH1] where DocEntry=" & CDbl(aDocEntry) & " group by U_Z_Branch")
                            'For intllop As Integer = 0 To oRecSet1.RecordCount - 1
                            '    Dim strBranch As String = oRecSet1.Fields.Item(0).Value
                            '    oRecSet2.DoQuery("Select * from [@Z_AL_OADM] where U_Z_BraDB='" & oRecSet1.Fields.Item(0).Value & "'")
                            '    If oRecSet2.RecordCount > 0 Then
                            '        oRemCompany = New SAPbobsCOM.Company
                            '        oRemCompany = oApplication.Utilities.ConnectRemoteCompany(oRecSet2.Fields.Item("U_Z_BraDB").Value, oRecSet2.Fields.Item("U_Z_SAPUID").Value, oRecSet2.Fields.Item("U_Z_SAPPWD").Value)
                            '        oJV1 = oRemCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)
                            '        oRecset3.DoQuery("Select * from [@Z_AL_CASH1] where U_Z_Branch='" & strBranch & "' and  DocEntry=" & CDbl(aDocEntry))
                            '        blnLineExists = False
                            '        oJV1 = oRemCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)
                            '        oJV1.JournalEntries.DueDate = dtDate
                            '        intCount = 0
                            '        For introw As Integer = 0 To oRecset3.RecordCount - 1
                            '            If oRecset3.Fields.Item("U_Z_Amount").Value > 0 Then


                            '                If intCount > 0 Then
                            '                    oJV1.JournalEntries.Lines.Add()
                            '                End If
                            '                oJV1.JournalEntries.Lines.SetCurrentLine(intCount)
                            '                oJV1.JournalEntries.Lines.AccountCode = getSAPAccount(oRecset3.Fields.Item("U_Z_ToDAcc").Value, oRemCompany)
                            '                oJV1.JournalEntries.Lines.Debit = oRecset3.Fields.Item("U_Z_Amount").Value
                            '                blnLineExists = True
                            '                intCount = intCount + 1
                            '                If intCount > 0 Then
                            '                    oJV1.JournalEntries.Lines.Add()
                            '                End If
                            '                oJV1.JournalEntries.Lines.SetCurrentLine(intCount)
                            '                oJV1.JournalEntries.Lines.AccountCode = getSAPAccount(oRecset3.Fields.Item("U_Z_ToCAcc").Value, oRemCompany)
                            '                oJV1.JournalEntries.Lines.Credit = oRecset3.Fields.Item("U_Z_Amount").Value
                            '                intCount = intCount + 1
                            '                blnLineExists = True
                            '            End If
                            '            oRecset3.MoveNext()
                            '        Next
                            '        If blnLineExists = True Then
                            '            If oJV1.Add <> 0 Then
                            '                oApplication.Utilities.Message(oRemCompany.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '                Return False
                            '            Else
                            '                oRemCompany.GetNewObjectCode(st)
                            '                '  MsgBox(st)
                            '            End If

                            '        End If
                            '    End If
                            '    oRemCompany.Disconnect()

                            '    oRecSet1.MoveNext()
                            'Next
                        End If
                    End If
                End If
            End If


            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function


    Private Function CreateJournalVoucher_Branch(ByVal aDocEntry As String) As Boolean
        Dim oJV, oJV1 As SAPbobsCOM.JournalVouchers
        Dim oRemCompany As SAPbobsCOM.Company
        Dim oRecSet, oRecSet1, oRecSet2, oRecset3 As SAPbobsCOM.Recordset
        Dim dtDate As Date
        Dim blnLineExists As Boolean = False
        oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSet1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSet2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecset3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRecSet.DoQuery("Select * from [@Z_AL_OCASH] where isnull(U_Z_Status,'O')='O' and DocEntry=" & CDbl(aDocEntry))
            If oRecSet.Fields.Item("U_Z_JVNo").Value = "" Then
                dtDate = oRecSet.Fields.Item("CreateDate").Value
                oRecSet.DoQuery("Select * from [@Z_AL_CASH1] where DocEntry=" & CDbl(aDocEntry))
                If 1 = 1 Then
                    If 1 = 2 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    Else
                        Dim st As String
                        Dim intCount As Integer
                        '    oApplication.Company.GetNewObjectCode(st)
                        oRecSet1.DoQuery("Select U_Z_Branch,Count(*) from [@Z_AL_CASH1] where DocEntry=" & CDbl(aDocEntry) & " group by U_Z_Branch")
                        For intllop As Integer = 0 To oRecSet1.RecordCount - 1
                            Dim strBranch As String = oRecSet1.Fields.Item(0).Value
                            If strBranch <> "" Then


                                oRecSet2.DoQuery("Select * from [@Z_AL_OADM] where U_Z_BraDB='" & oRecSet1.Fields.Item(0).Value & "'")
                                If oRecSet2.RecordCount > 0 Then
                                    oRemCompany = New SAPbobsCOM.Company
                                    oRemCompany = oApplication.Utilities.ConnectRemoteCompany(oRecSet2.Fields.Item("U_Z_BraDB").Value, oRecSet2.Fields.Item("U_Z_SAPUID").Value, oRecSet2.Fields.Item("U_Z_SAPPWD").Value)
                                    If oRemCompany.InTransaction Then
                                        oRemCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                    End If
                                    oRemCompany.StartTransaction()
                                    oJV1 = oRemCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)
                                    oRecset3.DoQuery("Select * from [@Z_AL_CASH1] where U_Z_Branch='" & strBranch & "' and  DocEntry=" & CDbl(aDocEntry))
                                    blnLineExists = False
                                    oJV1 = oRemCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)
                                    oJV1.JournalEntries.DueDate = dtDate
                                    intCount = 0
                                    For introw As Integer = 0 To oRecset3.RecordCount - 1
                                        If oRecset3.Fields.Item("U_Z_Amount").Value > 0 Then
                                            If intCount > 0 Then
                                                oJV1.JournalEntries.Lines.Add()
                                            End If
                                            oJV1.JournalEntries.Lines.SetCurrentLine(intCount)
                                            oJV1.JournalEntries.Lines.AccountCode = getSAPAccount(oRecset3.Fields.Item("U_Z_ToDAcc").Value, oRemCompany)
                                            oJV1.JournalEntries.Lines.Debit = oRecset3.Fields.Item("U_Z_Amount").Value
                                            blnLineExists = True
                                            intCount = intCount + 1
                                            If intCount > 0 Then
                                                oJV1.JournalEntries.Lines.Add()
                                            End If
                                            oJV1.JournalEntries.Lines.SetCurrentLine(intCount)
                                            oJV1.JournalEntries.Lines.AccountCode = getSAPAccount(oRecset3.Fields.Item("U_Z_ToCAcc").Value, oRemCompany)
                                            oJV1.JournalEntries.Lines.Credit = oRecset3.Fields.Item("U_Z_Amount").Value
                                            intCount = intCount + 1
                                            blnLineExists = True
                                        End If
                                        oRecset3.MoveNext()
                                    Next
                                    If blnLineExists = True Then
                                        If oJV1.Add <> 0 Then
                                            oApplication.Utilities.Message(oRemCompany.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            If oRemCompany.InTransaction Then
                                                oRemCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                            End If
                                            Return False
                                        Else
                                            oRemCompany.GetNewObjectCode(st)
                                            '  MsgBox(st)
                                        End If

                                    End If
                                End If
                                If oRemCompany.InTransaction Then
                                    oRemCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                End If
                                oRemCompany.Disconnect()

                            End If
                            oRecSet1.MoveNext()
                        Next

                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region


#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SalesTarget Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "11" Or pVal.ItemUID = "13" Or pVal.ItemUID = "4" Or pVal.ItemUID = "6") And pVal.CharPressed <> 9 Then
                                    'BubbleEvent = False
                                    'Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                If (pVal.ItemUID = "16" Or pVal.ItemUID = "17") And pVal.Row > 0 Then
                                    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = pVal.ItemUID
                                    frmSourceMatrix = oMatrix
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN


                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "8"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "9"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                    Case "14"
                                        oForm.PaneLevel = 1
                                    Case "15"
                                        oForm.PaneLevel = 2

                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val As String
                                Dim val2 As Integer
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
                                        'If pVal.ItemUID = "7" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1" Or pVal.ColUID = "V_3" Or pVal.ColUID = "V_4") Then
                                        '    val = oDataTable.GetValue("FormatCode", 0)
                                        '    oMatrix = oForm.Items.Item("7").Specific
                                        '    Try
                                        '        oApplication.Utilities.SetMatrixValues(oMatrix, pVal.ColUID, pVal.Row, val)
                                        '    Catch ex As Exception

                                        '    End Try

                                        'End If
                                        If pVal.ItemUID = "16" And pVal.ColUID = "SlpCode" Then
                                            val = oDataTable.GetValue("SlpCode", 0)
                                            val1 = oDataTable.GetValue("SlpName", 0)
                                            oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "SlpName", pVal.Row, val1)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "SlpCode", pVal.Row, val)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
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
                Case mnu_SalesTarget
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("6").Enabled = False
                        oForm.Items.Item("4").Enabled = False
                    End If
                Case mnu_ADD_ROW

                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddRow(oForm)
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    Else
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        AddMode(oForm)
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("4").Enabled = True
                        oForm.Items.Item("6").Enabled = True
                    End If

            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_CashTransfer Then
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oForm.Items.Item("4").Enabled = False
                    oForm.Items.Item("6").Enabled = True
                End If
            End If
            If BusinessObjectInfo.BeforeAction = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                ' strDocEntry = oApplication.Utilities.getEdittextvalue(oForm, "4")
            End If
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Class
