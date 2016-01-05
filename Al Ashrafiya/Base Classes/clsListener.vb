Public Class clsListener
    Inherits Object
    Private ThreadClose As New Threading.Thread(AddressOf CloseApp)
    Private WithEvents _SBO_Application As SAPbouiCOM.Application
    Private _Company As SAPbobsCOM.Company
    Private _RemoteCompany As SAPbobsCOM.Company
    Private _Utilities As clsUtilities
    Private _Collection As Hashtable
    Private _LookUpCollection As Hashtable
    Private _FormUID As String
    Private _Log As clsLog_Error
    Private oMenuObject As Object
    Private oItemObject As Object
    Private oSystemForms As Object
    Dim objFilters As SAPbouiCOM.EventFilters
    Dim objFilter As SAPbouiCOM.EventFilter
#Region "New"
    Public Sub New()
        MyBase.New()
        Try
            _Company = New SAPbobsCOM.Company
            _Utilities = New clsUtilities
            _Collection = New Hashtable(10, 0.5)
            _LookUpCollection = New Hashtable(10, 0.5)
            oSystemForms = New clsSystemForms
            _Log = New clsLog_Error

            SetApplication()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Public Properties"

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return _SBO_Application
        End Get
    End Property

    Public ReadOnly Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
    End Property

    Public ReadOnly Property RemoteCompany() As SAPbobsCOM.Company
        Get
            Return _RemoteCompany
        End Get
    End Property

    Public ReadOnly Property Utilities() As clsUtilities
        Get
            Return _Utilities
        End Get
    End Property

    Public ReadOnly Property Collection() As Hashtable
        Get
            Return _Collection
        End Get
    End Property

    Public ReadOnly Property LookUpCollection() As Hashtable
        Get
            Return _LookUpCollection
        End Get
    End Property

    Public ReadOnly Property Log() As clsLog_Error
        Get
            Return _Log
        End Get
    End Property
#Region "Filter"

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        oApplication.SetFilter(Filters)
    End Sub
    Public Sub SetFilter()
        Try
            ''Form Load
            objFilters = New SAPbouiCOM.EventFilters

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            objFilter.AddEx(frm_Import)
            objFilter.AddEx(frm_Export)
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
            objFilter.AddEx(frm_Import)
            objFilter.AddEx(frm_Export)
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            objFilter.AddEx(frm_Import)
            objFilter.AddEx(frm_Export)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            objFilter.AddEx(frm_itemmaster)
            objFilter.AddEx(frm_BPMaster)
            objFilter.AddEx(frm_SalesOrder)
            objFilter.AddEx(frm_PurchaseOrder)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            objFilter.AddEx(frm_itemmaster)
            objFilter.AddEx(frm_BPMaster)
            objFilter.AddEx(frm_SalesOrder)
            objFilter.AddEx(frm_PurchaseOrder)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            objFilter.AddEx(frm_Import)
            objFilter.AddEx(frm_Export)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region

#End Region

#Region "Menu Event"

    Private Sub _SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.FormDataEvent
        Select Case BusinessObjectInfo.FormTypeEx
            Case frm_Delivery
                Dim objInvoice As clsSalesOrder
                objInvoice = New clsSalesOrder
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_CashTransfer
                Dim objInvoice As clsCashTransfer
                objInvoice = New clsCashTransfer
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            Case frm_PurchaseQuatation
                Dim objInvoice As clsPurchaseQuatation
                objInvoice = New clsPurchaseQuatation
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            Case frm_Setup
                Dim objInvoice As clsLogin
                objInvoice = New clsLogin
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)


            Case frm_ALPrint
                Dim objInvoice As clsDocumentPrint
                objInvoice = New clsDocumentPrint
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)

        End Select
        '  End If
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID

                    Case mnu_Setup
                        oMenuObject = New clsLogin
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_CashTransfer
                        oMenuObject = New clsCashTransfer
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_SalesTarget
                        oMenuObject = New clsSalesTarget
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Print
                        oMenuObject = New clsDocumentPrint
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ADD, mnu_FIND, mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD_ROW, mnu_DELETE_ROW
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If

                End Select

            Else
                Select Case pVal.MenuUID
                    Case mnu_CLOSE, mnu_ADD_ROW, mnu_DELETE_ROW
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If
                End Select

            End If

        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oMenuObject = Nothing
        End Try
    End Sub
#End Region

#Region "Item Event"
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ItemEvent
        Try
            _FormUID = FormUID
            If pVal.FormTypeEx = frm_CostView Then
                Dim oForm As SAPbouiCOM.Form
                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                If oForm.TypeEx = frm_CostView Then
                    If pVal.ItemUID = "3" And pVal.Before_Action = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        Dim ogrid As SAPbouiCOM.Grid
                        Dim oTest As SAPbobsCOM.Recordset
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ogrid = oForm.Items.Item("1").Specific
                        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        For intRow As Integer = 0 To ogrid.DataTable.Rows.Count - 1
                            oTest.DoQuery("Update [@Z_AL_COST] set U_Z_PropMarging='" & ogrid.DataTable.GetValue("Proposed", intRow) & "' where Code='" & ogrid.DataTable.GetValue("Code", intRow) & "'")
                        Next
                        oTest.DoQuery("Update [@Z_AL_COST] set U_Z_Margin= ((U_Z_PropMarging-U_Z_LandedCost)/isnull(U_Z_PropMarging,1))*100 where U_Z_PropMarging >0 ")
                        oTest.DoQuery("Update [@Z_AL_COST] set U_Z_Margin= 0 where U_Z_PropMarging <=0 ")
                        oApplication.Utilities.Message("Operation compleated successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        oForm.Close()
                    End If

                    If pVal.ItemUID = "1" And pVal.Before_Action = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN And pVal.CharPressed = 9 Then
                        Dim ogrid As SAPbouiCOM.Grid
                        Dim oTest As SAPbobsCOM.Recordset
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ogrid = oForm.Items.Item("1").Specific
                        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Dim dblPropMarging, dblLandedCost As Double
                        dblPropMarging = ogrid.DataTable.GetValue(pVal.ColUID, pVal.Row)
                        If dblPropMarging <= 0 Then
                            dblPropMarging = 1
                        End If
                        dblLandedCost = ogrid.DataTable.GetValue("Landed Cost in KWD", pVal.Row)
                        dblPropMarging = (dblPropMarging - dblLandedCost) / dblPropMarging
                        dblPropMarging = dblPropMarging * 100
                        ogrid.DataTable.SetValue("Margin%", pVal.Row, dblPropMarging)

                    End If
                End If
            End If
            If pVal.FormTypeEx = frm_GRReceipt Then
                Dim oForm As SAPbouiCOM.Form
                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                If oForm.TypeEx = frm_GRReceipt And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                    If pVal.ItemUID = "3" And pVal.Before_Action = False Then
                        Dim ogrid As SAPbouiCOM.Grid
                        Dim oTest As SAPbobsCOM.Recordset
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ogrid = oForm.Items.Item("4").Specific
                        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Dim oCombo, oCombo1 As SAPbouiCOM.ComboBoxColumn
                        oCombo = ogrid.Columns.Item("IA")
                        oCombo1 = ogrid.Columns.Item("MR")
                        Dim strDate1, strDate2 As String
                        Dim strApproval, strMR As String

                        For intRow As Integer = 0 To ogrid.DataTable.Rows.Count - 1
                            Try
                                strApproval = oCombo.GetSelectedValue(intRow).Value
                            Catch ex As Exception
                                strApproval = "No"
                            End Try

                            Try
                                strMR = oCombo1.GetSelectedValue(intRow).Value
                            Catch ex As Exception
                                strMR = "LI"
                            End Try
                          
                            strDate1 = ogrid.DataTable.GetValue("RVD", intRow)
                            strDate2 = ogrid.DataTable.GetValue("MD", intRow)
                            Dim dtDate1, dtdate2 As Date

                            oTest.DoQuery("Update [PDN1] set U_MR='" & strMR & "',U_MN='" & ogrid.DataTable.GetValue("MN", intRow) & "',U_desc='" & ogrid.DataTable.GetValue("Desc", intRow) & "', U_IS='" & ogrid.DataTable.GetValue("IS", intRow) & "' , U_IA='" & strApproval & "' where DocEntry=" & ogrid.DataTable.GetValue("DocEntry", intRow) & " and LineNum=" & ogrid.DataTable.GetValue("LineNum", intRow))
                            If strDate1 <> "" Then
                                dtDate1 = ogrid.DataTable.GetValue("RVD", intRow)
                                oTest.DoQuery("Update [PDN1] set U_RVD='" & dtDate1.ToString("yyyy-MM-dd") & "' where DocEntry=" & ogrid.DataTable.GetValue("DocEntry", intRow) & " and LineNum=" & ogrid.DataTable.GetValue("LineNum", intRow))

                            End If
                            If strDate2 <> "" Then
                                dtDate1 = ogrid.DataTable.GetValue("MD", intRow)
                                oTest.DoQuery("Update [PDN1] set U_MD='" & dtDate1.ToString("yyyy-MM-dd") & "' where DocEntry=" & ogrid.DataTable.GetValue("DocEntry", intRow) & " and LineNum=" & ogrid.DataTable.GetValue("LineNum", intRow))
                            End If
                        Next
                        oApplication.Utilities.Message("Operation compleated successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        oForm.Close()
                    End If
                End If
            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.Before_Action = False Then
                Select Case pVal.FormTypeEx
                    Case frm_Setup
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsLogin
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_GoodsReceipt
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsGoodsReceipt
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_CashTransfer
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCashTransfer
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_SalesTarget
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSalesTarget
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_PurchaseQuatation
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsPurchaseQuatation
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_CostSheet
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCosting
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)

                        End If
                    Case frm_ALPrint
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDocumentPrint
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)

                        End If
                End Select
            End If

            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                Select Case pVal.FormTypeEx
                    Case frm_FuturaSetup
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsAcctMapping
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_CashTransfer
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCashTransfer
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_SalesTarget
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSalesTarget
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ALPrint
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDocumentPrint
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)

                        End If
                End Select
            End If

            If _Collection.ContainsKey(FormUID) Then
                oItemObject = _Collection.Item(FormUID)
                If oItemObject.IsLookUpOpen And pVal.BeforeAction = True Then
                    _SBO_Application.Forms.Item(oItemObject.LookUpFormUID).Select()
                    BubbleEvent = False
                    Exit Sub
                End If
                _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.BeforeAction = False Then
                If _LookUpCollection.ContainsKey(FormUID) Then
                    oItemObject = _Collection.Item(_LookUpCollection.Item(FormUID))
                    If Not oItemObject Is Nothing Then
                        oItemObject.IsLookUpOpen = False
                    End If
                    _LookUpCollection.Remove(FormUID)
                End If

                If _Collection.ContainsKey(FormUID) Then
                    _Collection.Item(FormUID) = Nothing
                    _Collection.Remove(FormUID)
                End If

            End If

        Catch ex As Exception
            If ex.Message.Contains("Form - Invalid Form") Then
            Else
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Application Event"
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles _SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    _Utilities.AddRemoveMenus("RemoveMenus.xml")
                    CloseApp()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Try
    End Sub
#End Region

#Region "Close Application"
    Private Sub CloseApp()
        Try
            If Not _SBO_Application Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_SBO_Application)
            End If

            If Not _Company Is Nothing Then
                If _Company.Connected Then
                    _Company.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company)
            End If

            _Utilities = Nothing
            _Collection = Nothing
            _LookUpCollection = Nothing

            ThreadClose.Sleep(10)
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Throw ex
        Finally
            oApplication = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Set Application"
    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                SboGuiApi = New SAPbouiCOM.SboGuiApi
                SboGuiApi.Connect(sConnectionString)
                _SBO_Application = SboGuiApi.GetApplication()
            Else
                Throw New Exception("Connection string missing.")
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SboGuiApi = Nothing
        End Try
    End Sub
#End Region

#Region "Finalize"
    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            '            CloseApp()

            oMenuObject = Nothing
            oItemObject = Nothing
            oSystemForms = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Addon Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

End Class
