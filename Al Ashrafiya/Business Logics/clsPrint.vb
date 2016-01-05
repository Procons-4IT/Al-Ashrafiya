Imports System
Imports System.Collections
Imports System.ComponentModel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading
Imports System.Collections.Generic
Public Class clsPrint
    'Private rptaccountreport As New AcctStatement
    Dim cryRpt As New ReportDocument
    Private ds As New dsAL
    Private oDRow As DataRow
#Region "Add Crystal Report"

    Private Sub addCrystal(ByVal ds1 As DataSet, ByVal aChoice As String, ByVal aCode As String)
        Dim strFilename, strCompanyName, stfilepath As String
        Dim blnCrystal As Boolean = False
        Dim strReportFileName As String
        If aChoice = "Inv" Then
            strReportFileName = "Invoice.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Invoices"
            blnCrystal = False
        Else : aChoice = "Del"
            strReportFileName = "Delivery.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Delivery"
            blnCrystal = False
        End If
        strReportFileName = strReportFileName
        strFilename = strFilename & ".pdf"
        '  strFilename = strFilename & ".doc"
        stfilepath = System.Windows.Forms.Application.StartupPath & "\CrystalReports\" & strReportFileName
        If File.Exists(stfilepath) = False Then
            oApplication.Utilities.Message("Report does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        If File.Exists(strFilename) Then
            File.Delete(strFilename)
        End If
        If aCode = "W" Then
            blnCrystal = True
        Else
            blnCrystal = False
        End If
        ' If ds1.Tables.Item("AccountBalance").Rows.Count > 0 Then
        If 1 = 1 Then
            cryRpt.Load(System.Windows.Forms.Application.StartupPath & "\CrystalReports\" & strReportFileName)
            Try
                cryRpt.SetDataSource(ds1)
            Catch ex As Exception
            End Try

            If blnCrystal = True Then
                Dim mythread As New System.Threading.Thread(AddressOf openFileDialog)
                mythread.SetApartmentState(ApartmentState.STA)
                mythread.Start()
                mythread.Join()
                ds1.Clear()
            Else

                'Dim CrExportOptions As ExportOptions
                'Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                ''Dim CrFormatTypeOptions As New DiskFileDestinationOptions
                'CrDiskFileDestinationOptions.DiskFileName = strFilename
                'CrExportOptions = cryRpt.ExportOptions
                'With CrExportOptions
                '    .ExportDestinationType = ExportDestinationType.DiskFile
                '    .ExportFormatType = ExportFormatType.WordForWindows
                '    .DestinationOptions = CrDiskFileDestinationOptions
                '    '  .FormatOptions = CrFormatTypeOptions
                'End With
                'cryRpt.Export()
                'cryRpt.Close()

                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New _
                DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
                CrDiskFileDestinationOptions.DiskFileName = strFilename
                CrExportOptions = cryRpt.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                cryRpt.Export()
                cryRpt.Close()
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                x.FileName = strFilename
                System.Diagnostics.Process.Start(x)
                x = Nothing
                ' objUtility.ShowSuccessMessage("Report exported into PDF File")
            End If

        Else
            ' objUtility.ShowWarningMessage("No data found")
        End If

    End Sub

    Private Sub openFileDialog()
        Dim objPL As New frmReportViewer
        objPL.iniViewer = AddressOf objPL.GenerateReport
        objPL.rptViewer.ReportSource = cryRpt
        objPL.rptViewer.Refresh()
        objPL.WindowState = FormWindowState.Maximized
        objPL.ShowDialog()
        System.Threading.Thread.CurrentThread.Abort()
    End Sub

    Public Sub PrintInvoice(ByVal aGrid As SAPbouiCOM.Grid, ByVal aChoice As String, ByVal aRptOption As String)
        Dim oRec, oRecTemp, oRecBP, oBalanceRs, oTemp As SAPbobsCOM.Recordset
        Dim strfrom, dtPosting, dtdue, dttax, strto, strBranch, strSlpCode, strSlpName, strSMNo, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim intReportChoice As Integer
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        ds.Clear()
        Dim blnLineExist As Boolean = False
        Dim oCheckbox As SAPbouiCOM.CheckBoxColumn
        Dim intDocEntry As Integer
        Dim strTable, strtable1 As String
        If aChoice = "Inv" Then
            strTable = "OINV"
            strtable1 = "INV1"
        Else
            strTable = "ODLN"
            strtable1 = "DLN1"
        End If
        Dim oRec1, oTemp1 As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ' Dim strCity As String
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oCheckbox = aGrid.Columns.Item("Select")
            If oCheckbox.IsChecked(intRow) = True Then
                blnLineExist = True
                strCity = aGrid.DataTable.GetValue("City", intRow)

                intDocEntry = aGrid.DataTable.GetValue("DocEntry", intRow)
                oRec.DoQuery("Select * from " & strTable & " where DocEntry=" & intDocEntry)

                If oRec.RecordCount > 0 Then
                    oDRow = ds.Tables("Header").NewRow()
                    oDRow.Item("DocEntry") = oRec.Fields.Item("DocEntry").Value
                    Try
                        oDRow.Item("Currency") = oRec.Fields.Item("DocCur").Value
                    Catch ex As Exception

                    End Try
                    oDRow.Item("City") = strCity
                    oDRow.Item("DocNum") = oRec.Fields.Item("DocNum").Value ' oRec.Fields.Item("DocEntry").Value
                    oDRow.Item("CardCode") = oRec.Fields.Item("CardCode").Value
                    oDRow.Item("CardName") = oRec.Fields.Item("CardName").Value
                    oDRow.Item("BillTo") = oRec.Fields.Item("Address").Value
                    oDRow.Item("Comments") = oRec.Fields.Item("Comments").Value
                    oTemp.DoQuery("Select isnull(Phone1,'') from OCRD where Cardcode='" & oRec.Fields.Item("CardCode").Value & "'")
                    oDRow.Item("Phone1") = oTemp.Fields.Item(0).Value
                    oDRow.Item("DocDate") = oRec.Fields.Item("DocDate").Value
                    oTemp.DoQuery("Select * from OSLP where SlpCode='" & oRec.Fields.Item("SlpCode").Value & "'")

                    oDRow.Item("SlpName") = oTemp.Fields.Item("SlpName").Value
                    oTemp.DoQuery("Select Sum(LineTotal) from " & strtable1 & "  where DocEntry=" & intDocEntry) '='" & oRec.Fields.Item("SlpCode").Value & "'")
                    oDRow.Item("SubTotal") = oTemp.Fields.Item(0).Value
                    ' oTemp.DoQuery("select isnull(U_Z_LocName,'') from [@Z_OLOC] where DocEntry =" & oRec.Fields.Item("U_Z_ToLOC").Value)
                    oDRow.Item("Discount") = oRec.Fields.Item("DiscPrcnt").Value
                    If oRec.Fields.Item("DocCur").Value <> oApplication.Company.GetCompanyService.GetAdminInfo.LocalCurrency Then
                        oDRow.Item("DiscountValue") = oRec.Fields.Item("DiscSumFC").Value
                        oDRow.Item("Doctotal") = oRec.Fields.Item("DocTotalFC").Value
                    Else
                        oDRow.Item("DiscountValue") = oRec.Fields.Item("DiscSum").Value
                        oDRow.Item("Doctotal") = oRec.Fields.Item("DocTotal").Value
                    End If
                    oTemp1.DoQuery("SELECT T1.[CardCode], isnull(T1.[Building],'') 'Building', isnull(T1.[Address],'') 'Street', isnull(T2.[Name],'') 'Country' FROM " & strTable & " T0  INNER JOIN OCRD T1 ON T0.CardCode = T1.CardCode LEFT OUTER JOIN OCRY T2 ON T0.BnkCntry = T2.Code where T0.DocEntry=" & intDocEntry)
                    If oTemp1.RecordCount > 0 Then
                        oDRow.Item("Building") = oTemp1.Fields.Item("Building").Value ' oRec.Fields.Item("DocEntry").Value
                        oDRow.Item("Street") = oTemp1.Fields.Item("Street").Value
                        oDRow.Item("Country") = oTemp1.Fields.Item("Country").Value
                    End If
                    oDRow.Item("DN") = oRec.Fields.Item("U_DN").Value
                    ds.Tables("Header").Rows.Add(oDRow)
                    oRec1.DoQuery("Select * from " & strtable1 & " where DocEntry=" & intDocEntry)
                    For intLoop As Integer = 0 To oRec1.RecordCount - 1
                        oDRow = ds.Tables("Lines").NewRow()
                        oDRow.Item("DocEntry") = oRec1.Fields.Item("DocEntry").Value
                        oDRow.Item("LineID") = oRec1.Fields.Item("LineNum").Value ' oRec.Fields.Item("DocEntry").Value
                        oDRow.Item("ItemCode") = oRec1.Fields.Item("ItemCode").Value
                        oDRow.Item("ItemName") = oRec1.Fields.Item("Dscription").Value
                        oDRow.Item("Qty") = oRec1.Fields.Item("Quantity").Value
                        oDRow.Item("Unit") = oRec1.Fields.Item("UOMCODE").Value
                        oDRow.Item("UnitPrice") = oRec1.Fields.Item("Price").Value
                        oDRow.Item("LineTotal") = oRec1.Fields.Item("LineTotal").Value
                        ds.Tables("Lines").Rows.Add(oDRow)
                        oRec1.MoveNext()
                    Next
                End If
            End If
        Next
        If blnLineExist = False Then
            oApplication.Utilities.Message("No Document selected", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        addCrystal(ds, aChoice, aRptOption)

        oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)

    End Sub

















#End Region


End Class
