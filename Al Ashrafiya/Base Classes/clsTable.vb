Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try
            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab
                oUserTablesMD.TableDescription = strDesc
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            oUserTablesMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "OADM" Or strTab = "IGN1" Or strTab = "OITM" Or strTab = "OCRD" Or strTab = "INV1" Or strTab = "RDR1" Or strTab = "OINV" Or strTab = "ORDR") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                oUserFieldMD.Description = strDesc
                oUserFieldMD.Name = strCol
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory
                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If



            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    '  MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            Else


            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            objUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try


    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)

            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

            If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                '// Set the table name and the key name
                oUserKeysMD.TableName = strTab
                oUserKeysMD.KeyName = strKey

                '// Set the column's alias
                oUserKeysMD.Elements.ColumnAlias = strColumn
                oUserKeysMD.Elements.Add()
                oUserKeysMD.Elements.ColumnAlias = "RentFac"

                '// Determine whether the key is unique or not
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                '// Add the key
                If oUserKeysMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
            oUserKeysMD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document, Optional ByVal strChildTb2 As String = "")

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.LogTableName = ""
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If

                If strChildTb2 <> "" Then
                    oUserObjectMD.ChildTables.Add()
                    oUserObjectMD.ChildTables.TableName = strChildTb2
                End If

                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub

#End Region

#Region "Public Functions"
    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************
    Public Sub CreateTables()
        Dim oProgressBar As SAPbouiCOM.ProgressBar
        Try

            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            AddFields("OITM", "Pack", "Packing", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddFields("OCRD", "Z_CardCode", "Branch Supplier Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OCRD", "Z_BranchDB", "Branch DB Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("ODLN", "Z_Exported", "Exported", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("OPDN", "DG", "DG", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_Address, "Yes,No", "Yes,No", "No")
            AddFields("OPDN", "DeliveryNo", "Delivery Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)


            addField("OBTF", "DJV", "DJV", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_Address, "Yes,No", "Yes,No", "No")
            AddFields("OITM", "COO", "COO", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("OINV", "Z_LandedCost", "Landed Cost in KWD", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)

            AddFields("OINV", "DN", "Driver Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("IGN1", "IA", "Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, SAPbobsCOM.BoFldSubTypes.st_Address, "Yes,No", "Yes,No", "No")
            AddFields("IGN1", "RVD", "Received Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("IGN1", "MD", "Municipality Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("IGN1", "IS", "Item Status", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            AddFields("IGN1", "MN", "Municipality Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("IGN1", "desc", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("IGN1", "MR", "Municipality Rejection", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_Address, "LI,DI", "Lab Issue,Documentation Issue", "LI")



            AddTables("Z_AL_OADM", "Branch DB Details", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_AL_OADM", "Z_BraDB", "Branch DB Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_AL_OADM", "Z_BraSQL", "Branch SQL Server Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_AL_OADM", "Z_SQLUID", "Branch SQL User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_AL_OADM", "Z_SQLPWD", "Branch SQL Password", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_AL_OADM", "Z_SAPUID", "Branch DB SAP User Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AL_OADM", "Z_SAPPWD", "Branch DB SAP Password", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)


            AddTables("Z_AL_OCASH", "Cash transfer to Branch", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("Z_AL_CASH1", "Cash Transfer Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            addField("@Z_AL_OCASH", "Z_Status", "Posting Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "O,C", "Open , Close", "O")
            AddFields("Z_AL_OCASH", "Z_JVNo", "Journal Voucher Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AL_OCASH", "Z_Currency", "Transfer Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)

            AddFields("Z_AL_CASH1", "Z_frmCAcc", "Main Branch Credit Account ", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AL_CASH1", "Z_frmDAcc", "Main Branch Debit Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AL_CASH1", "Z_Branch", "Branch DB Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_AL_CASH1", "Z_ToCAcc", "Transfer Branch Credit Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AL_CASH1", "Z_ToDAcc", "Transfer Branch Debit Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            AddFields("Z_AL_CASH1", "Z_Amount", "Transfer Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_CASH1", "Z_Remarks", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_AL_CASH1", "Z_JVNo", "Branch JV Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)


            AddTables("Z_AL_OSAL", "Sales Target", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("Z_AL_SAL1", "Sales Target Month Wise", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("Z_AL_SAL2", "Commission Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("Z_AL_SAL1", "Z_SlpCode", "Sales Person Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AL_SAL1", "Z_SlpName", "Sales Person Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_AL_OSAL", "Z_FnYear", "Fiancial Year", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            AddFields("Z_AL_SAL1", "Z_Jan", "January ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL1", "Z_Feb", "Feburary ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL1", "Z_Mar", "March ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL1", "Z_Apr", "April ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL1", "Z_May", "May ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL1", "Z_Jun", "June ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL1", "Z_Jul", "July ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL1", "Z_Aug", "August ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL1", "Z_Sep", "September ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL1", "Z_Oct", "October", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL1", "Z_Nov", "November ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL1", "Z_Dec", "December ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)

            AddFields("Z_AL_SAL2", "Z_ComFrom", "Commission From ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("Z_AL_SAL2", "Z_ComTo", "Commission End ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("Z_AL_SAL2", "Z_Tar1", "Target from 70-100% ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL2", "Z_Tar2", "100-115%", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_SAL2", "Z_Tar3", "Avoce 115 % ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)

            AddFields("INV1", "Z_Cost_Ref", "Costing Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)

            AddTables("Z_AL_COST", "Costing Analysis", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("OITM", "Z_Origion", "Origion", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("OITM", "SE", "SE", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("OITM", "CE", "CE", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_AL_COST", "Z_DocEntry", "Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_AL_COST", "Z_DocNum", "Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_AL_COST", "Z_ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_AL_COST", "Z_ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_AL_COST", "Z_Origion", "Origin", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_AL_COST", "Z_LineID", "Line Number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_AL_COST", "Z_Unit", "UoM", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_AL_COST", "Z_Packaging", "Packaging", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_AL_COST", "Z_OrderQty", "Purchase Quatation Qty", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_AL_COST", "Z_InStock", "In Stock Qty", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_AL_COST", "Z_AvgMonSal", "Average Monthly Sales", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_COST", "Z_PPrice", "Last Purchase Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_COST", "Z_Price", "Purchase Quatation Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_COST", "Z_Totalvalue", "Total Value in Sup.Cur", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_COST", "Z_LandedCost", "Landed Cost in KWD", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_COST", "Z_AvgSelling", "Avg Selling Price in KWD", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_COST", "Z_Month", "Monthly %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("Z_AL_COST", "Z_PropMarging", "Proposed Marging", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AL_COST", "Z_Margin", "Margin %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("Z_AL_COST", "Z_PurCur", "Purchase Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_AL_COST", "Z_PQCur", "Purchase Quatation Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_AL_COST", "Z_POQty", "Open Purchase Order Qty", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_AL_COST", "Z_CardCode", "Supplier COde", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_AL_COST", "Z_CardName", "Supplier Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_AL_COST", "Z_NumAtCard", "NumAtCard", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            AddFields("Z_AL_COST", "Z_QDate", "Quatation Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_AL_COST", "Z_SE", "SE", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_AL_COST", "Z_CE", "CE", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            '---- User Defined Object's
            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            CreateUDO()

            oApplication.Utilities.Message("Initializing Database Completed...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        Catch ex As Exception
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Function UDOSetup(ByVal strUDO As String, _
                    ByVal strDesc As String, _
                        ByVal strTable As String, _
                            ByVal intFind As Integer, _
                                Optional ByVal strCode As String = "", _
                                    Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""
                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()

                oUserObjects.FormColumns.FormColumnAlias = "U_Z_BraDB"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_BraDB"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_BraSQL"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_BraSQL"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_SQLUID"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_SQLUID"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_SQLPWD"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_SQLPWD"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_SAPUID"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_SAPUID"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_SAPPWD"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_SAPPWD"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                    ' oUserObjects.FindColumns.Add()
                    'oUserObjects.FindColumns.SetCurrentLine(1)
                    'oUserObjects.FindColumns.ColumnAlias = strName
                    'oUserObjects.FindColumns.Add()
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function
    Public Sub CreateUDO()
        Try
            ' UDOSetup("Z_AL_OADM", "Branch DB Setups", "Z_AL_OADM", 1, "U_Z_BraDB", )
            AddUDO("Z_AL_OADM", "Login Setup", "Z_AL_OADM", "DocEntry", "U_Z_BraDB", , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_AL_OSAL", "Sales Target Setup", "Z_AL_OSAL", "DocEntry", "U_Z_FnYear", "Z_AL_SAL1", SAPbobsCOM.BoUDOObjType.boud_Document, "Z_AL_SAL2")
            AddUDO("Z_AL_CASH", "Cash Transfer to branch", "Z_AL_OCASH", "DocNum", "U_Z_Status", "Z_AL_CASH1", SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class
