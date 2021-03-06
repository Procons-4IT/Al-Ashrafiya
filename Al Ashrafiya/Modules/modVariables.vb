Public Module modVariables

    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String
    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public frmSourceFormUD As String
    Public frmSourceForm As SAPbouiCOM.Form
    Public frmSourcePMForm As SAPbouiCOM.Form
    Public frmSourceQCOR As SAPbouiCOM.Form

    Public LoalDB As String
    Public intCurrentRow As Integer = 10000


    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public strCardCode As String = ""
    Public blnDraft As Boolean = False
    Public blnError As Boolean = False
    Public strDocEntry As String
    Public strImportErrorLog As String = ""
    Public companyStorekey As String = ""
    Public strDatabasename As String = "WMS"
    Public strSKUExportTable As String = "[WMS].dbo.[SKU_Export]"
    Public strSOExportTable As String = "[WMS].dbo.[SO_Export]"
    Public strARCRExportTable As String = "[WMS].dbo.[ARCR_Export]"
    Public strPOExportTable As String = "[WMS].dbo.[PO_Export]"

    Public intSelectedMatrixrow As Integer = 0
    Public strSourceformEmpID As String = ""
    Public strApprovalType As String = ""

    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocumentType As Integer
        RENTAL_QUOTATION = 1
        RENTAL_ORDER
        RENTAL_RETURN
    End Enum
    Public Const frm_GoodsReceipt As String = "143"
    Public Const frm_GRReceipt As String = "frm_GoodsReceiptCost"
    Public Const xml_GRREceipt As String = "frm_GoodsReceiptCost.xml"

    Public Const xml_CostView As String = "xml_Al_CostView.xml"
    Public Const frm_CostView As String = "frm_AL_CostView"
    Public Const frm_PurchaseQuatationUDF As String = "-540000988"
    Public Const frm_PurchaseQuatation As String = "540000988"
    Public Const frm_CostSheet As String = "frm_Cost"
    Public Const xml_CostingSheet As String = "frm_CostSheet.xml"
    Public Const frm_CostElement As String = "frm_CostElement"

    Public Const mnu_Print As String = "Z_mnu_Al006"
    Public Const frm_ALPrint As String = "frm_ALPrint"
    Public Const xml_ALPrint As String = "xml_AL_Print.xml"

    Public Const mnu_SalesTarget As String = "Z_mnu_Al005"
    Public Const frm_SalesTarget As String = "frm_ALSalesT"
    Public Const xml_SalesTarget As String = "xml_AL_SalesTarget.xml"

    Public Const mnu_CashTransfer As String = "Z_mnu_Al004"
    Public Const frm_CashTransfer As String = "frm_ALCash"
    Public Const xml_CashTransfer As String = "xml_AL_CashTransfer.xml"


    Public Const mnu_Setup As String = "Z_mnu_Al003"
    Public Const frm_Setup As String = "frm_ALSetup"
    Public Const xml_Setup As String = "xml_AL_Setup.xml"
    Public Const frm_Delivery As String = "140"


    Public Const frm_WAREHOUSES As Integer = 62
    Public Const frm_FuturaSetup As String = "frm_FuturaSetup"
    Public Const frm_StockRequest As String = "frm_StRequest"
    Public Const frm_itemmaster As String = "150"
    Public Const frm_BPMaster As String = "134"
    Public Const frm_InvSO As String = "frm_InvSO"
    Public Const frm_Warehouse As String = "62"
    Public Const frm_SalesOrder As String = "139"
    Public Const frm_ARCreditMemo As String = "179"
    Public Const frm_PurchaseOrder As String = "142"
    Public Const frm_Invoice As String = "133"
    Public Const frm_Import As String = "frm_Import"
    Public Const frm_Export As String = "frm_Export"

  
    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"
    Public Const mnu_CloseOrderLines As String = "DABT_910"
    Public Const mnu_InvSO As String = "DABT_911"
    Public Const mnu_DuplicateRow As String = "1294"

    Public Const mnu_Import As String = "Z_mnu_D003"
    Public Const mnu_Export As String = "Z_mnu_D002"
    Public Const mnu_Mapping As String = "Z_mnu_FU003"
    
    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"
    Public Const xml_StRequest As String = "StRequest.xml"
    Public Const xml_InvSO As String = "frm_InvSO.xml"
    Public Const xml_Import As String = "frm_Import.xml"
    Public Const xml_Export As String = "frm_Export.xml"
    Public Const xml_Futurasetup As String = "frm_FuturaSetup.xml"
   

End Module
