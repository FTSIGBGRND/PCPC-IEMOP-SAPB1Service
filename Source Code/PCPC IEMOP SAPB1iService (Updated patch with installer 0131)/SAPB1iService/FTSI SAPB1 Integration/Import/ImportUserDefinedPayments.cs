using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using SAPbobsCOM;
using System.Windows;
using System.Windows.Forms;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;
using System.Data.SqlClient;
using Microsoft.VisualBasic.FileIO;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace SAPB1iService
{
    class ImportUserDefinedPayments
    {
        private static DateTime dteStart;
        private static string strTransType = "Payment - IEMOP Incoming";
        private static string strMsgBod;
        public static void _ImportUserDefinedPayments()
        {
            importFromFile();
        }
        public static bool importFromObject(int intObjType, string strAPDocEntry, DateTime dteDoc, DateTime dtePayDate, DateTime dteTax)
        {
            string strQuery, strCardCode, strStatus, strMsgBod;

            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Payments oPayments;

            try
            {
                strQuery = string.Format("SELECT OPCH.\"DocEntry\", OPCH.\"DocNum\", OPCH.\"CardCode\", OPCH.\"Comments\", OPCH.\"DocTotal\" " +
                                         "FROM OPCH " +
                                         "WHERE OPCH.\"DocEntry\" = '{0}' ", strAPDocEntry);

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                {

                    strCardCode = oRecordset.Fields.Item("CardCode").Value.ToString();

                    GlobalFunction.getObjType(intObjType);

                    oPayments = null;
                    oPayments = (SAPbobsCOM.Payments)GlobalVariable.oCompany.GetBusinessObject(GlobalVariable.oObjectType);

                    oPayments.CardCode = strCardCode;
                    oPayments.DocDate = dteDoc;
                    oPayments.DueDate = dtePayDate;
                    oPayments.TaxDate = dteTax;
                    oPayments.Remarks = oRecordset.Fields.Item("Comments").Value.ToString();

                    oPayments.TransferDate = dtePayDate;
                    oPayments.TransferSum = Convert.ToDouble(oRecordset.Fields.Item("DocTotal").Value.ToString());

                    oPayments.Invoices.DocEntry = Convert.ToInt32(strAPDocEntry);
                    oPayments.Invoices.InvoiceType = BoRcptInvTypes.it_PurchaseInvoice;
                    oPayments.Invoices.SumApplied = Convert.ToDouble(oRecordset.Fields.Item("DocTotal").Value.ToString());

                    if (oPayments.Add() != 0)
                    {
                        GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                        GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                        return false;
                    }
                    else
                    {
                        GlobalVariable.strDocEntry = GlobalVariable.oCompany.GetNewObjectKey().ToString();
                        GlobalVariable.strDocNum = GlobalFunction.getDocNum(GlobalVariable.strDocEntry);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = string.Format("Error Processing Outgoing Payment. {0}.", ex.Message.ToString());

                return false;
            }
        }
        private static void importFromFile()
        {

            string strStatus = "";

            try
            {
                string[] strFileImport = GlobalVariable.strImpExt.Split(Convert.ToChar("|"));

                foreach (string fileimport in strFileImport)
                {
                    foreach (var strFile in Directory.GetFiles(GlobalVariable.strImpPath, string.Format("*{0}_*{1}", GlobalVariable.strCompany, fileimport)))
                    {
                        GlobalVariable.strFileName = Path.GetFileName(strFile);

                        if (strFile.Contains("Incoming"))
                        {
                            dteStart = DateTime.Now;

                            if (importDIAPIPostPaymentFExcel(strFile))
                                strStatus = "S";
                            else
                                strStatus = "E";

                            TransferFile.transferProcFiles("Import", strStatus, Path.GetFileName(strFile));

                            GlobalFunction.sendAlert(strStatus, "Import", strMsgBod, GlobalVariable.oObjectType, GlobalVariable.oCompany.GetNewObjectKey().ToString());

                            //EmailSender._EmailSender("Import", strStatus, GlobalVariable.strFileName, strPostDocNum, string.Format("Error Code : {0} Description : {1} ", GlobalVariable.intErrNum, GlobalVariable.strErrMsg));
                        }
                    }
                }

                GC.Collect();
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("Import", strTransType, "28", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }
        private static bool importDIAPIPostPaymentFExcel(string strFile)
        {
            string strSTLID = "", strNumAtCard = "", strCardCode = "", strRefNum, strTrnsfrGL, 
                   strQuery, strPostDocNum;

            string strWTCode, strTIN = "", strAddrss = "";

            bool blBPExist = true, blTempErr, blRClss = true;

            double dblTotPymnt, dblDocTotal, dblPayAmnt = 0, dblRunBal, dblTotsales, dblTotWTax, dblTotVatSales, dblTotZroRated, dblTotVAT, dblWTAmnt;

            int intRowInv;

            DataTable oDTSTLID, oDTPayment;

            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Payments oPayments;


            DateTime dteDoc = Convert.ToDateTime("01/01/1900");
            DateTime dteDue = Convert.ToDateTime("01/01/1900");
            DateTime dteTax = Convert.ToDateTime("01/01/1900");
            DateTime dteTrnsfr = Convert.ToDateTime("01/01/1900");
            DateTime dteStartPeriod = Convert.ToDateTime("01/01/1900");
            DateTime dteEndPeriod = Convert.ToDateTime("01/01/1900");

            try
            {

                if (GlobalFunction.importXLSX(Path.GetFullPath(strFile), "NO", "Sheet1"))
                {
                    if (GlobalVariable.oDTImpData.Rows.Count > 0)
                    {
                        blBPExist = true;
                        blTempErr = false;

                        strQuery = string.Format("SELECT IEMOPGL.\"U_PymntGL\" " +
                                                 "FROM \"@FTIEMOPGL\" IEMOPGL " +
                                                 "WHERE IEMOPGL.\"Code\" = 'AR' ");
                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (!(oRecordset.RecordCount > 0))
                        {
                            GlobalVariable.intErrNum = -998;
                            GlobalVariable.strErrMsg = "Please setup Transfer GL Account for Incoming Payment.";

                            strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", strTransType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            return false;
                        }
                        else
                            strTrnsfrGL = oRecordset.Fields.Item("U_PymntGL").Value.ToString();

                        for (int intRowH = 0; intRowH <= 6; intRowH++)
                        {
                            strSTLID = GlobalVariable.oDTImpData.Rows[intRowH][0].ToString();

                            //For Naming Convention Payment Period
                            if (strSTLID.Contains("As of"))
                            {
                                MatchCollection strmatches = Regex.Matches(strSTLID, @"(\w+ \d+, \d+)");
                                string strfrstdate = strmatches[0].Value;
                                string strscnddate = strmatches[1].Value;

                                if (validateDates(strfrstdate.ToString()) && validateDates(strscnddate.ToString())) 
                                {
                                    dteStartPeriod = Convert.ToDateTime(strfrstdate.ToString());
                                    dteEndPeriod = Convert.ToDateTime(strscnddate.ToString());
                                }
                                else
                                {
                                    blTempErr = true;
                                }
                            }


                            if (strSTLID == "Posting Date")
                            {
                                if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString()))
                                    dteDoc = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString());
                                else
                                    blTempErr = true;
                            }

                            if (strSTLID == "Due Date")
                            {
                                if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString()))
                                    dteDue = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString());
                                else
                                    blTempErr = true;
                            }

                            if (strSTLID == "Document Date")
                            {
                                if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString()))
                                    dteTax = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString());
                                else
                                    blTempErr = true;
                            }

                            if (strSTLID == "Transfer Date")
                            {
                                if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString()))
                                    dteTrnsfr = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString());
                                else
                                    blTempErr = true;
                            }
                        }

                        if (blTempErr == true)
                        {
                            GlobalVariable.intErrNum = -998;
                            GlobalVariable.strErrMsg = "Template Error. Please check dates.";

                            strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", strTransType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            return false;
                        }
                        
                        //oDTSTLID = GlobalVariable.oDTImpData.Select(string.Format("{0} IS NOT NULL AND {0} <> 'Received From (Buyer STL ID)' ", GlobalVariable.oDTImpData.Columns[2].ColumnName)).CopyToDataTable().DefaultView.ToTable(true, GlobalVariable.oDTImpData.Columns[2].ColumnName, GlobalVariable.oDTImpData.Columns[4].ColumnName);

                        oDTSTLID = GlobalVariable.oDTImpData.Select(string.Format("{0} IS NOT NULL AND {0} <> 'Received From (Buyer STL ID)' ", GlobalVariable.oDTImpData.Columns[2].ColumnName)).CopyToDataTable().DefaultView.ToTable(true, GlobalVariable.oDTImpData.Columns[2].ColumnName, GlobalVariable.oDTImpData.Columns[4].ColumnName);

                        if (!(GlobalVariable.oCompany.InTransaction))
                            GlobalVariable.oCompany.StartTransaction();

                        for (int intRowID = 0; intRowID <= oDTSTLID.Rows.Count - 1; intRowID++)
                        {
                            strSTLID = oDTSTLID.Rows[intRowID][0].ToString();
                            strNumAtCard = oDTSTLID.Rows[intRowID][1].ToString();

                            if (strNumAtCard.Substring(strNumAtCard.Length - 1) == "S" || strNumAtCard.Substring(strNumAtCard.Length - 1) == "s")
                                strRefNum = strNumAtCard.Remove(strNumAtCard.Length - 1, 1);
                            else
                                strRefNum = strNumAtCard;

                            //strCardCode = "C000001"; //oDTSTLID.Rows[intRowID][2].ToString();

                            strQuery = string.Format("SELECT TOP 1 OCRD.\"CardCode\", OCRD.\"CardName\", OCRD.\"LicTradNum\", OCRD.\"Address\", OCRD.\"CardType\", OCRD.\"U_IemopWtax\" " +
                                "FROM OCRD " + 
                                "INNER JOIN \"@FTIEMOP1\" ON OCRD.\"CardCode\" = \"@FTIEMOP1\".\"Code\" " + "" +
                                "WHERE \"@FTIEMOP1\".\"U_STLID\" = '{0}' OR \"@FTIEMOP1\".\"U_BLLID\" = '{0}' AND OCRD.\"CardType\" = 'C' ", strSTLID);
                            oRecordset = null;
                            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            oRecordset.DoQuery(strQuery);

                            if (oRecordset.RecordCount > 0)
                            {
                                strCardCode = oRecordset.Fields.Item("CardCode").Value.ToString();
                                strTIN = oRecordset.Fields.Item("LicTradNum").Value.ToString();
                                strAddrss = oRecordset.Fields.Item("Address").Value.ToString();
                                strWTCode = oRecordset.Fields.Item("U_IemopWtax").Value.ToString();

                                if (string.IsNullOrEmpty(strTIN) || string.IsNullOrEmpty(strAddrss))
                                {
                                    blRClss = false;

                                    GlobalVariable.intErrNum = -997;
                                    GlobalVariable.strErrMsg = string.Format("TIN and/or Address of Business Partner {0} for STLID {1} not exist. Please check BP Master Data in the template.", strCardCode, strSTLID);

                                    strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    continue;
                                }
                                
                            }
                            else
                            {
                                blBPExist = false;

                                GlobalVariable.intErrNum = -997;
                                GlobalVariable.strErrMsg = string.Format("Business Partner {0} for STLID {1} not exist. Please check BP Master Data in the template.", strCardCode, strSTLID);

                                strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                continue;
                            }

                            if (blBPExist == true && blRClss == true)
                            {                
                                
                                GlobalFunction.getObjType(24);

                                oPayments = null;
                                oPayments = (SAPbobsCOM.Payments)GlobalVariable.oCompany.GetBusinessObject(GlobalVariable.oObjectType);

                                oPayments.CardCode = strCardCode;
                                oPayments.DocDate = dteDoc;
                                oPayments.DueDate = dteDue;
                                oPayments.TaxDate = dteTax;

                                oPayments.UserFields.Fields.Item("U_TransType").Value = "IEMOP";
                                oPayments.UserFields.Fields.Item("U_PymntStrtPeriod").Value = dteStartPeriod;
                                oPayments.UserFields.Fields.Item("U_PymntEndPeriod").Value = dteEndPeriod;

                                strQuery = string.Format("{0} = '{1}' AND {2} = '{3}' ", GlobalVariable.oDTImpData.Columns[2].ColumnName, strSTLID, GlobalVariable.oDTImpData.Columns[4].ColumnName, strNumAtCard);
                                oDTPayment = GlobalVariable.oDTImpData.Select(strQuery).CopyToDataTable().DefaultView.ToTable();

                                dblTotPymnt = 0;
                                dblTotsales = 0;
                                dblTotWTax = 0;
                                dblTotVAT = 0;
                                dblTotVatSales = 0;
                                dblTotZroRated = 0;

                                for (int intRowDP = 0; intRowDP <= oDTPayment.Rows.Count - 1; intRowDP++)
                                {
                                    dblTotPymnt = dblTotPymnt + Math.Abs(Convert.ToDouble(oDTPayment.Rows[intRowDP][10].ToString()));

                                    dblTotsales = dblTotsales +
                                                    Math.Abs(Convert.ToDouble(oDTPayment.Rows[intRowDP][5].ToString())) +
                                                    Math.Abs(Convert.ToDouble(oDTPayment.Rows[intRowDP][6].ToString())) +
                                                    Math.Abs(Convert.ToDouble(oDTPayment.Rows[intRowDP][7].ToString())) +
                                                    Math.Abs(Convert.ToDouble(oDTPayment.Rows[intRowDP][8].ToString()));


                                    dblTotVAT = dblTotVAT + Math.Abs(Convert.ToDouble(oDTPayment.Rows[intRowDP][8].ToString()));

                                    dblTotWTax = dblTotWTax + Math.Abs(Convert.ToDouble(oDTPayment.Rows[intRowDP][9].ToString()));

                                    dblTotVatSales = dblTotVatSales + Math.Abs(Convert.ToDouble(oDTPayment.Rows[intRowDP][5].ToString()));

                                    dblTotZroRated = dblTotZroRated +
                                                        Math.Abs(Convert.ToDouble(oDTPayment.Rows[intRowDP][6].ToString())) +
                                                        Math.Abs(Convert.ToDouble(oDTPayment.Rows[intRowDP][7].ToString()));


                                }

                                oPayments.TransferAccount = strTrnsfrGL;
                                oPayments.TransferSum = Math.Abs(dblTotPymnt);
                                oPayments.TransferDate = dteTrnsfr;

                                oPayments.UserFields.Fields.Item("U_TotalSales").Value = dblTotsales;
                                oPayments.UserFields.Fields.Item("U_VAT").Value = dblTotVAT;
                                oPayments.UserFields.Fields.Item("U_WTax").Value = dblTotWTax;
                                oPayments.UserFields.Fields.Item("U_VatableSales").Value = dblTotVatSales;
                                oPayments.UserFields.Fields.Item("U_ZeroRated").Value = dblTotZroRated;

                                oPayments.Remarks = strRefNum;


                                //CHANGE REQUEST 27/06/2024 - Creditable Withholding Tax – Auto JE from IEMOP Incoming Payment
                                GlobalVariable.strCWTJENum = "0";
                                if (dblTotWTax != 0)
                                {
                                    if (!(ImportUserDefinedJournalEntry.importAutoCWT(strCardCode, strNumAtCard, dblTotWTax, dblTotVatSales, dblTotZroRated, dteDoc, dteTax, dteDue)))
                                    {
                                        strMsgBod = string.Format("Error Posting Journal Entry for {0} in file {1}.\rError Code: {2}\rDescription: {3} ", strCardCode, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                        SystemFunction.transHandler("Import", "Journal Entry - Recognition of CWT", GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), string.Format("{0} - CardCodee: {1} | STLID {2}", GlobalVariable.strErrMsg, strCardCode, strSTLID));

                                        return false;
                                    }
                                }

                                strQuery = string.Format("SELECT \"ObjType\", \"DocEntry\", \"DocNum\", \"DocTotal\", \"Line_ID\" " +
                                                         "FROM (SELECT OINV.\"ObjType\", OINV.\"DocEntry\", OINV.\"DocNum\", (OINV.\"DocTotal\" - OINV.\"PaidToDate\") AS \"DocTotal\", NULL AS \"Line_ID\" " +
                                                         "      FROM OINV " +
                                                         "      WHERE OINV.\"DocStatus\" = 'O' AND OINV.\"NumAtCard\" = '{0}' AND OINV.\"CardCode\" = '{1}' " +
                                                         "      UNION ALL " +
                                                         "      SELECT OJDT.\"ObjType\", OJDT.\"TransId\" AS \"DocEntry\", OJDT.\"Number\" AS \"DocNum\", " +
                                                         "      JDT1.\"Credit\" AS \"DocTotal\", " +
                                                         "      JDT1.\"Line_ID\" " +
                                                         "      FROM OJDT " +
                                                         "      INNER JOIN JDT1 ON OJDT.\"TransId\" = JDT1.\"TransId\" " +
                                                         "      WHERE OJDT.\"TransId\" = '{2}' AND JDT1.\"ShortName\" = '{1}' ) AS QRY " +
                                                         "ORDER BY \"ObjType\" DESC ", strRefNum, strCardCode, GlobalVariable.strCWTJENum);

                                oRecordset = null;
                                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                oRecordset.DoQuery(strQuery);

                                if (oRecordset.RecordCount > 0)
                                {
                                    intRowInv = 0;
                                    dblRunBal = dblTotPymnt;

                                    while (!(oRecordset.EoF))
                                    {
                                        dblDocTotal = Convert.ToDouble(oRecordset.Fields.Item("DocTotal").Value.ToString());

                                        if (dblRunBal > 0)
                                        {
                                            if (intRowInv > 0)
                                                oPayments.Invoices.Add();

                                            oPayments.Invoices.DocEntry = Convert.ToInt32(oRecordset.Fields.Item("DocEntry").Value.ToString());

                                            if (oRecordset.Fields.Item("ObjType").Value.ToString() == "13")
                                            {
                                                dblTotPymnt = Math.Round(dblTotPymnt - dblDocTotal, 2);

                                                oPayments.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;

                                                if (dblDocTotal > dblTotPymnt)
                                                    oPayments.Invoices.SumApplied = dblRunBal;
                                                else
                                                    oPayments.Invoices.SumApplied = dblDocTotal;

                                                

                                            }
                                            else
                                            {
                                                dblTotPymnt = Math.Round(dblTotPymnt + dblDocTotal, 2);

                                                oPayments.Invoices.InvoiceType = BoRcptInvTypes.it_JournalEntry;
                                                oPayments.Invoices.DocLine = Convert.ToInt32(oRecordset.Fields.Item("Line_ID").Value.ToString());
                                                oPayments.Invoices.SumApplied = dblDocTotal * -1;

                                                
                                            }

                                            dblRunBal = dblTotPymnt;

                                            intRowInv++;
                                        }

                                        oRecordset.MoveNext();
                                    }
                                }

                                else
                                {
                                    GlobalVariable.intErrNum = -997;
                                    GlobalVariable.strErrMsg = string.Format("AR Invoice Document Not Found for {0} - {1} - {2}. ", strCardCode, strSTLID, strRefNum);

                                    strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", strTransType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    return false;
                                }

                                if (oPayments.Add() != 0)
                                {
                                    GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                    strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", strTransType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), string.Format("{0} - CardCodee: {1} | STLID {2}", GlobalVariable.strErrMsg, strCardCode, strSTLID));

                                    return false;
                                }
                                /*else
                                {                                                          
                                    strPostDocNum = GlobalFunction.getDocNum(GlobalVariable.oCompany.GetNewObjectKey().ToString());                                    

                                    if (dblTotVAT != 0)
                                    {
                                        if (!(ImportUserDefinedJournalEntry.importFromObject("Post", dteDoc, GlobalVariable.oCompany.GetNewObjectKey().ToString(), strPostDocNum, dblTotVAT, strCardCode, dblTotVatSales, dblTotZroRated)))
                                        {
                                            strMsgBod = string.Format("Error Posting Deferred Tax ReClass for {0} in file {1}.\rError Code: {2}\rDescription: {3} ", strCardCode, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                            SystemFunction.transHandler("Import", "Journal Entry - Deferred Tax ReClass", GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), string.Format("{0} - CardCodee: {1} | STLID {2}", GlobalVariable.strErrMsg, strCardCode, strSTLID));

                                            return false;
                                        }
                                    }
                                }*/
                            }

                        }
                    }

                    if (blBPExist == true && blRClss == true)
                    {
                        strMsgBod = string.Format("Successfully Posted {0} - {1}.", strTransType, GlobalVariable.strFileName);

                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "S", "0", strMsgBod);

                        if (GlobalVariable.oCompany.InTransaction)
                            GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                        GC.Collect();

                        ExportORAutoPDF._ExportORAutoPDF(strNumAtCard);                        

                        return true;                        
                    }
                    else
                        return false;

                }
                else
                    return false;

                GC.Collect();

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                return false;
            }
        }
        private static bool validateDates(string strDate)
        {
            DateTime dteReturn;

            if (!DateTime.TryParse(strDate, out dteReturn))
                return false;
            else
            {
                if (dteReturn == Convert.ToDateTime("01/01/1900"))
                    return false;
                else
                    return true;
            }

        }
    }
}
