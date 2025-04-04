using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using SAPB1iService;
using System.Windows;
using System.Windows.Forms;
using System.Data;
using System.Threading;
using System.ServiceProcess;
using System.Xml;
using SAPbobsCOM;
using System.Data.OleDb;
using DidiSoft.Pgp;

namespace SAPB1iService
{
    class GlobalFunction
    {
        public static void getObjType(int ObjType)
        {
            switch (ObjType)
            {
                case 13:
                    GlobalVariable.strDocType = "AR Invoice";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oInvoices;
                    GlobalVariable.intObjType = 13;
                    GlobalVariable.strTableHeader = "OINV";
                    GlobalVariable.strTableLine1 = "INV1";
                    GlobalVariable.strTableLine3 = "INV3";
                    GlobalVariable.strTableLine5 = "INV5";
                    break;

                case 14:
                    GlobalVariable.strDocType = "AR Credit Memo";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oCreditNotes;
                    GlobalVariable.intObjType = 14;
                    GlobalVariable.strTableHeader = "ORIN";
                    GlobalVariable.strTableLine1 = "RIN1";
                    GlobalVariable.strTableLine3 = "RIN3";
                    GlobalVariable.strTableLine5 = "RIN5";
                    break;

                case 15:
                    GlobalVariable.strDocType = "Delivery";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oDeliveryNotes;
                    GlobalVariable.intObjType = 15;
                    GlobalVariable.strTableHeader = "ODLN";
                    GlobalVariable.strTableLine1 = "DLN1";
                    GlobalVariable.strTableLine3 = "DLN3";
                    GlobalVariable.strTableLine5 = "DLN5";
                    break;

                case 16:
                    GlobalVariable.strDocType = "Sales Return";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oReturns;
                    GlobalVariable.intObjType = 16;
                    GlobalVariable.strTableHeader = "ORDN";
                    GlobalVariable.strTableLine1 = "RDN1";
                    GlobalVariable.strTableLine3 = "RDN3";
                    GlobalVariable.strTableLine5 = "RDN5";

                    break;

                case 17:
                    GlobalVariable.strDocType = "Sales Order";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oOrders;
                    GlobalVariable.intObjType = 17;
                    GlobalVariable.strTableHeader = "ORDR";
                    GlobalVariable.strTableLine1 = "RDR1";
                    GlobalVariable.strTableLine3 = "RDR3";
                    GlobalVariable.strTableLine5 = "RDR5";
                    break;

                case 18:
                    GlobalVariable.strDocType = "AP Invoice";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                    GlobalVariable.intObjType = 18;
                    GlobalVariable.strTableHeader = "OPCH";
                    GlobalVariable.strTableLine1 = "PCH1";
                    GlobalVariable.strTableLine3 = "PCH3";
                    GlobalVariable.strTableLine5 = "PCH5";
                    break;

                case 19:
                    GlobalVariable.strDocType = "AP Debit Memo";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes;
                    GlobalVariable.intObjType = 19;
                    GlobalVariable.strTableHeader = "ORPC";
                    GlobalVariable.strTableLine1 = "RPC1";
                    GlobalVariable.strTableLine3 = "RPC3";
                    GlobalVariable.strTableLine5 = "RPC5";
                    break;

                case 20:
                    GlobalVariable.strDocType = "Goods Receipt PO";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;
                    GlobalVariable.intObjType = 20;
                    GlobalVariable.strTableHeader = "OPDN";
                    GlobalVariable.strTableLine1 = "PDN1";
                    GlobalVariable.strTableLine3 = "PDN3";
                    GlobalVariable.strTableLine5 = "PDN5";
                    break;

                case 21:
                    GlobalVariable.strDocType = "Goods Return";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseReturns;
                    GlobalVariable.intObjType = 21;
                    GlobalVariable.strTableHeader = "ORPD";
                    GlobalVariable.strTableLine1 = "RPD1";
                    GlobalVariable.strTableLine3 = "RPD3";
                    GlobalVariable.strTableLine5 = "RPD5";
                    break;

                case 22:
                    GlobalVariable.strDocType = "Purchase Order";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;
                    GlobalVariable.intObjType = 22;
                    GlobalVariable.strTableHeader = "OPOR";
                    GlobalVariable.strTableLine1 = "POR1";
                    GlobalVariable.strTableLine3 = "POR3";
                    GlobalVariable.strTableLine5 = "POR5";
                    break;

                case 23:
                    GlobalVariable.strDocType = "Sales Quotations";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseQuotations;
                    GlobalVariable.strTableHeader = "OQUT";
                    GlobalVariable.strTableLine1 = "QUT1";
                    GlobalVariable.strTableLine3 = "QUT3";
                    GlobalVariable.strTableLine5 = "QUT5";
                    break;

                case 24:
                    GlobalVariable.strDocType = "Incoming Payment";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oIncomingPayments;
                    GlobalVariable.strTableHeader = "ORCT";
                    break;

                case 30:
                    GlobalVariable.strDocType = "Journal Entry";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oJournalEntries;
                    GlobalVariable.strTableHeader = "OJDT";
                    break;

                case 46:
                    GlobalVariable.strDocType = "Outgoing Payment";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oVendorPayments;
                    GlobalVariable.strTableHeader = "OVPM";
                    break;

                case 59:
                    GlobalVariable.strDocType = "Goods Receipt";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oInventoryGenEntry;
                    GlobalVariable.strTableHeader = "OIGN";
                    GlobalVariable.strTableLine1 = "IGN1";
                    break;

                case 60:
                    GlobalVariable.strDocType = "Goods Issue";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oInventoryGenExit;
                    GlobalVariable.strTableHeader = "OIGE";
                    GlobalVariable.strTableLine1 = "IGE1";
                    break;

                case 67:
                    GlobalVariable.strDocType = "Stock Transfer";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                    GlobalVariable.strTableHeader = "OWTR";
                    GlobalVariable.strTableLine1 = "WTR1";
                    break;

                case 112:
                    GlobalVariable.strDocType = "Draft";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oDrafts;
                    GlobalVariable.strTableHeader = "ODRF";
                    GlobalVariable.strTableLine1 = "DRF1";
                    break;

                case 204:
                    GlobalVariable.strDocType = "AP DownPayment";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments;
                    GlobalVariable.strTableHeader = "ODPO";
                    GlobalVariable.strTableLine1 = "DPO1";
                    break;

                case 1250000001:
                    GlobalVariable.strDocType = "Stock Transfer Request";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest;
                    GlobalVariable.strTableHeader = "OWTQ";
                    GlobalVariable.strTableLine1 = "WTQ1";
                    break;

                case 1470000113:
                    GlobalVariable.strDocType = "Purchase Request";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseRequest;
                    GlobalVariable.strTableHeader = "OPRQ";
                    GlobalVariable.strTableLine1 = "PRQ1";
                    break;

                case 28:
                    GlobalVariable.strDocType = "Journal Voucher";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oJournalVouchers;
                    GlobalVariable.intObjType = 28;
                    GlobalVariable.strTableHeader = "OBTD";
                    break;

                default:
                    GlobalVariable.strDocType = "";
                    GlobalVariable.oObjectType = 0;
                    GlobalVariable.strTableHeader = "";
                    break;

            }
        }
        public static void getBaseType(int ObjType)
        {
            switch (ObjType)
            {
                case 13:
                    GlobalVariable.strBDocType = "AR Invoice";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oInvoices;
                    GlobalVariable.intBObjType = 13;
                    GlobalVariable.strBTableHeader = "OINV";
                    GlobalVariable.strBTableLine1 = "INV1";
                    GlobalVariable.strBTableLine3 = "INV3";
                    GlobalVariable.strBTableLine5 = "INV5";
                    break;

                case 14:
                    GlobalVariable.strBDocType = "AR Credit Memo";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oCreditNotes;
                    GlobalVariable.intBObjType = 14;
                    GlobalVariable.strBTableHeader = "ORIN";
                    GlobalVariable.strBTableLine1 = "RIN1";
                    GlobalVariable.strBTableLine3 = "RIN3";
                    GlobalVariable.strBTableLine5 = "RIN5";
                    break;

                case 15:
                    GlobalVariable.strBDocType = "Delivery";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oDeliveryNotes;
                    GlobalVariable.intBObjType = 15;
                    GlobalVariable.strBTableHeader = "ODLN";
                    GlobalVariable.strBTableLine1 = "DLN1";
                    GlobalVariable.strBTableLine3 = "DLN3";
                    GlobalVariable.strBTableLine5 = "DLN5";
                    break;

                case 16:
                    GlobalVariable.strBDocType = "Sales Return";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oReturns;
                    GlobalVariable.intBObjType = 16;
                    GlobalVariable.strBTableHeader = "ORDN";
                    GlobalVariable.strBTableLine1 = "RDN1";
                    GlobalVariable.strBTableLine3 = "RDN3";
                    GlobalVariable.strBTableLine5 = "RDN5";

                    break;

                case 17:
                    GlobalVariable.strBDocType = "Sales Order";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oOrders;
                    GlobalVariable.intBObjType = 17;
                    GlobalVariable.strBTableHeader = "ORDR";
                    GlobalVariable.strBTableLine1 = "RDR1";
                    GlobalVariable.strBTableLine3 = "RDR3";
                    GlobalVariable.strBTableLine5 = "RDR5";
                    break;

                case 18:
                    GlobalVariable.strBDocType = "AP Invoice";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                    GlobalVariable.intBObjType = 18;
                    GlobalVariable.strBTableHeader = "OPCH";
                    GlobalVariable.strBTableLine1 = "PCH1";
                    GlobalVariable.strBTableLine3 = "PCH3";
                    GlobalVariable.strBTableLine5 = "PCH5";
                    break;

                case 19:
                    GlobalVariable.strBDocType = "AP Debit Memo";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes;
                    GlobalVariable.intBObjType = 19;
                    GlobalVariable.strBTableHeader = "ORPC";
                    GlobalVariable.strBTableLine1 = "RPC1";
                    GlobalVariable.strBTableLine3 = "RPC3";
                    GlobalVariable.strBTableLine5 = "RPC5";
                    break;

                case 20:
                    GlobalVariable.strBDocType = "Goods Receipt PO";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;
                    GlobalVariable.intBObjType = 20;
                    GlobalVariable.strBTableHeader = "OPDN";
                    GlobalVariable.strBTableLine1 = "PDN1";
                    GlobalVariable.strBTableLine3 = "PDN3";
                    GlobalVariable.strBTableLine5 = "PDN5";
                    break;

                case 21:
                    GlobalVariable.strBDocType = "Goods Return";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseReturns;
                    GlobalVariable.intBObjType = 21;
                    GlobalVariable.strBTableHeader = "ORPD";
                    GlobalVariable.strBTableLine1 = "RPD1";
                    GlobalVariable.strBTableLine3 = "RPD3";
                    GlobalVariable.strBTableLine5 = "RPD5";
                    break;

                case 22:
                    GlobalVariable.strBDocType = "Purchase Order";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;
                    GlobalVariable.intBObjType = 22;
                    GlobalVariable.strBTableHeader = "OPOR";
                    GlobalVariable.strBTableLine1 = "POR1";
                    GlobalVariable.strBTableLine3 = "POR3";
                    GlobalVariable.strBTableLine5 = "POR5";
                    break;

                case 23:
                    GlobalVariable.strBDocType = "Sales Quotations";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseQuotations;
                    GlobalVariable.strBTableHeader = "OQUT";
                    GlobalVariable.strBTableLine1 = "QUT1";
                    GlobalVariable.strBTableLine3 = "QUT3";
                    GlobalVariable.strBTableLine5 = "QUT5";
                    break;

                case 24:
                    GlobalVariable.strBDocType = "Incoming Payment";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oIncomingPayments;
                    GlobalVariable.strBTableHeader = "ORCT";
                    break;

                case 30:
                    GlobalVariable.strBDocType = "Journal Entry";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oJournalEntries;
                    GlobalVariable.strBTableHeader = "OJDT";
                    break;

                case 46:
                    GlobalVariable.strBDocType = "Outgoing Payment";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oVendorPayments;
                    GlobalVariable.strBTableHeader = "OVPM";
                    break;

                case 59:
                    GlobalVariable.strBDocType = "Goods Receipt";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oInventoryGenEntry;
                    GlobalVariable.strBTableHeader = "OIGN";
                    GlobalVariable.strBTableLine1 = "IGN1";
                    break;

                case 60:
                    GlobalVariable.strBDocType = "Goods Issue";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oInventoryGenExit;
                    GlobalVariable.strBTableHeader = "OIGE";
                    GlobalVariable.strBTableLine1 = "IGE1";
                    break;

                case 67:
                    GlobalVariable.strBDocType = "Stock Transfer";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                    GlobalVariable.strBTableHeader = "OWTR";
                    GlobalVariable.strBTableLine1 = "WTR1";
                    break;

                case 112:
                    GlobalVariable.strBDocType = "Draft";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oDrafts;
                    GlobalVariable.strBTableHeader = "ODRF";
                    GlobalVariable.strBTableLine1 = "DRF1";
                    break;

                case 204:
                    GlobalVariable.strBDocType = "AP DownPayment";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments;
                    GlobalVariable.strBTableHeader = "ODPO";
                    GlobalVariable.strBTableLine1 = "DPO1";
                    break;

                case 1250000001:
                    GlobalVariable.strBDocType = "Stock Transfer Request";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest;
                    GlobalVariable.strBTableHeader = "OWTQ";
                    GlobalVariable.strBTableLine1 = "WTQ1";
                    break;

                case 1470000113:
                    GlobalVariable.strBDocType = "Purchase Request";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseRequest;
                    GlobalVariable.strBTableHeader = "OPRQ";
                    GlobalVariable.strBTableLine1 = "PRQ1";
                    break;

                default:
                    GlobalVariable.strBDocType = "";
                    GlobalVariable.oBObjectType = 0;
                    GlobalVariable.strBTableHeader = "";
                    break;

            }
        }
        public static void sendAlert(string strStatus, string strProcess, string strMsgTxt, SAPbobsCOM.BoObjectTypes ObjType, string strObjKey)
        {
            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Messages oMessages;

            string strSubject = "FT SAP B1 Services - " + strProcess; 

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT \"USER_CODE\", \"U_NAME\" FROM OUSR WHERE \"U_IntMsg\" = 'Y' ");

            if (oRecordset.RecordCount > 0)
            {
                oMessages = null;
                oMessages = (SAPbobsCOM.Messages)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                oMessages.Subject = strSubject;
                oMessages.MessageText = strMsgTxt;
                if (strStatus != "E")
                    oMessages.AddDataColumn("Document #", strObjKey, ObjType, strObjKey);
                oMessages.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;


                while (!(oRecordset.EoF))
                {
                    if (oRecordset.RecordCount > 1)
                        oMessages.Recipients.Add();

                    oMessages.Recipients.UserCode = oRecordset.Fields.Item("USER_CODE").Value.ToString();
                    oMessages.Recipients.NameTo = oRecordset.Fields.Item("U_NAME").Value.ToString();
                    oMessages.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                    oMessages.Recipients.UserType = SAPbobsCOM.BoMsgRcpTypes.rt_InternalUser;

                    oRecordset.MoveNext();
                }

                if (oMessages.Add() != 0)
                {
                    GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                    SystemFunction.errorAppend(GlobalVariable.intErrNum.ToString() + " - " + GlobalVariable.strErrMsg);
                }
            }
        }
        public static string getDocNum(string strDocEntry)
        {
            string strDocNum;

            SAPbobsCOM.Recordset oRecordset;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT " + GlobalVariable.strTableHeader + ".\"DocNum\" FROM " + GlobalVariable.strTableHeader + " WHERE  " + GlobalVariable.strTableHeader + ".\"DocEntry\" = '" + strDocEntry + "' ");

            if (oRecordset.RecordCount > 0)
                strDocNum = oRecordset.Fields.Item("DocNum").Value.ToString();
            else
                strDocNum = "0";

            return strDocNum;
        }
        public static string getJENum(string strDocEntry)
        {
            string strDocNum;

            SAPbobsCOM.Recordset oRecordset;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT Number FROM OJDT WHERE TransID = '" + strDocEntry + "' ");

            if (oRecordset.RecordCount > 0)
                strDocNum = oRecordset.Fields.Item("Number").Value.ToString();
            else
                strDocNum = "0";

            return strDocNum;
        }
        public static string getJVNum(string strDocEntry)
        {
            string strDocNum;

            SAPbobsCOM.Recordset oRecordset;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT BatchNum FROM OBTD WHERE BatchNum = '" + strDocEntry + "' ");

            if (oRecordset.RecordCount > 0)
                strDocNum = oRecordset.Fields.Item("Number").Value.ToString();
            else
                strDocNum = "0";

            return strDocNum;
        }
        public static void createResponse(string strFileName, string strStatus, string strATECSAPDoc, string strRemarks, string strDate, string strTime)
        {
            string strXMLPath, strALBSAPDoc;

            string[] strFValue;


            try
            {
                strFValue = strFileName.Split(Convert.ToChar("_"));
                strALBSAPDoc = strFValue[3];

                strXMLPath = GlobalVariable.strExpPath + @"\RES_" + strFileName;

                XmlTextWriter xWriter = new XmlTextWriter(strXMLPath, Encoding.UTF8);
                xWriter.Formatting = Formatting.Indented;

                xWriter.WriteStartElement("ResponseFile");

                xWriter.WriteStartElement("ALBSAPDoc");
                xWriter.WriteString(strALBSAPDoc);
                xWriter.WriteEndElement();

                xWriter.WriteStartElement("Status");
                xWriter.WriteString(strStatus);
                xWriter.WriteEndElement();

                xWriter.WriteStartElement("ATECSAPDoc");
                xWriter.WriteString(strATECSAPDoc);
                xWriter.WriteEndElement();

                xWriter.WriteStartElement("Remarks");
                xWriter.WriteString(strRemarks);
                xWriter.WriteEndElement();

                xWriter.WriteStartElement("Date");
                xWriter.WriteString(strDate);
                xWriter.WriteEndElement();

                xWriter.WriteStartElement("Time");
                xWriter.WriteString(strTime);
                xWriter.WriteEndElement();

                xWriter.WriteEndElement();
                xWriter.Close();

            }
            catch (Exception ex)
            {

            }
        }
        public static DateTime getDateTime(string strDateTime, string strOrigFormat, string strRetFormat)
        {
            DateTime dteRetDate = Convert.ToDateTime("01/01/9999");
            string strDateVal;

            if (!(string.IsNullOrEmpty(strDateTime)))
            {
                if (strRetFormat == "MM/DD/YYYY")
                {
                    if (strOrigFormat == "YYYYMMDD")
                    {

                        if (strDateTime.Length == 7)
                        {
                            if (Convert.ToInt32(strDateTime.Substring(4, 2)) > 12)
                                strDateVal = strDateTime.Substring(4, 1) + "/" + strDateTime.Substring(5, 2) + "/" + strDateTime.Substring(0, 4);
                            else
                                strDateVal = strDateTime.Substring(4, 2) + "/" + strDateTime.Substring(6, 1) + "/" + strDateTime.Substring(0, 4);
                        }
                        else if (strDateTime.Length == 6)
                            strDateVal = strDateTime.Substring(4, 1) + "/" + strDateTime.Substring(5, 1) + "/" + strDateTime.Substring(0, 4);
                        else
                            strDateVal = strDateTime.Substring(4, 2) + "/" + strDateTime.Substring(6, 2) + "/" + strDateTime.Substring(0, 4);

                        dteRetDate = Convert.ToDateTime(strDateVal);
                    }
                    else if (strOrigFormat == "MMDDYYYY")
                    {
                        if (strDateTime.Length == 7)
                        {
                            if (Convert.ToInt32(strDateTime.Substring(0, 2)) > 12)
                                strDateVal = strDateTime.Substring(0, 1) + "/" + strDateTime.Substring(1, 2) + "/" + strDateTime.Substring(3, 4);
                            else
                                strDateVal = strDateTime.Substring(0, 2) + "/" + strDateTime.Substring(2, 1) + "/" + strDateTime.Substring(3, 4);
                        }
                        else if (strDateTime.Length == 6)
                            strDateVal = strDateTime.Substring(0, 1) + "/" + strDateTime.Substring(1, 1) + "/" + strDateTime.Substring(2, 4);
                        else
                            strDateVal = strDateTime.Substring(0, 2) + "/" + strDateTime.Substring(2, 2) + "/" + strDateTime.Substring(4, 4);

                        dteRetDate = Convert.ToDateTime(strDateVal);
                    }
                }
                else if (strRetFormat == "YYYY/MM/DD")
                {
                    if (strOrigFormat == "MMDDYYYY")
                    {
                        if (strDateTime.Length == 7)
                        {
                            if (Convert.ToInt32(strDateTime.Substring(0, 2)) > 12)
                                strDateVal = strDateTime.Substring(3, 4) + "/0" + strDateTime.Substring(0, 1) + "/" + strDateTime.Substring(1, 2);                           
                            else
                                strDateVal = strDateTime.Substring(3, 4) + "/" + strDateTime.Substring(0, 2) + "/" + strDateTime.Substring(2, 1);
                        }
                        else if (strDateTime.Length == 6)
                            strDateVal = strDateTime.Substring(2, 4) + "/0" + strDateTime.Substring(0, 1) + "/0" + strDateTime.Substring(1, 1);
                        else
                            strDateVal = strDateTime.Substring(0, 4) + "/" + strDateTime.Substring(4, 2) + "/" + strDateTime.Substring(6, 2);

                        dteRetDate = Convert.ToDateTime(strDateVal);
                    }
                    else if (strOrigFormat == "YYYYMMDD")
                    {
                        if (strDateTime.Length == 7)
                        {
                            if (Convert.ToInt32(strDateTime.Substring(4, 2)) > 12)
                                strDateVal = strDateTime.Substring(0, 4) + "/0" + strDateTime.Substring(4, 1) + "/" + strDateTime.Substring(5, 2);
                            else
                                strDateVal = strDateTime.Substring(0, 4) + strDateTime.Substring(4, 2) + "/0" + strDateTime.Substring(6, 1);
                        }
                        else if (strDateTime.Length == 6)
                            strDateVal = strDateTime.Substring(0, 4) + "/0" + strDateTime.Substring(4, 1) + "/0" + strDateTime.Substring(5, 1);
                        else
                            strDateVal = strDateTime.Substring(0, 4) + "/" + strDateTime.Substring(4, 2) + "/" + strDateTime.Substring(6, 2); 

                        dteRetDate = Convert.ToDateTime(strDateVal);
                    }

                }

            }
            else
                dteRetDate = Convert.ToDateTime("01/01/1900");

            return dteRetDate;


        }
        public static bool importXLSX(string strXLSPath, string strHeader, string strSheet)
        {
            try
            {
                string strConnString;

                GlobalVariable.oDTImpData = new DataTable("ImportData");

                GlobalVariable.oDTImpData.Clear();

                if (Path.GetExtension(strXLSPath) == ".xlsx")      
                    strConnString = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended Properties = 'Excel 12.0; HDR={1};' ", strXLSPath, strHeader);
                else
                    strConnString = string.Format("Provider = Microsoft.Jet.OleDb.4.0; Data Source = {0}; Extended Properties = 'Excel 8.0; HDR={1};' ", strXLSPath, strHeader);

                using (OleDbConnection oledbConn = new OleDbConnection(strConnString))
                {
                    oledbConn.Open();

                    using (OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}$]", strSheet), oledbConn))
                    {
                        using (OleDbDataAdapter oleda = new OleDbDataAdapter(cmd))
                        {
                            oleda.Fill(GlobalVariable.oDTImpData);
                        }
                    }

                    oledbConn.Close();
                }

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.errorAppend(string.Format("Error retrieving data from Excel ({0} - {1}). Description : {2} ", strXLSPath, strSheet, ex.Message.ToString()));
                SystemFunction.transHandler("Import", "", GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", DateTime.Now, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);
                //TransferFile.transferProcFiles("Import", "E", Path.GetFileName(strXLSPath));
                return false;
            }

        }
        public static bool importCSV(string strFilePath, string strCSVPath, string strHeader, string strDlmtd)
        {
            try
            {

                GlobalVariable.oDTImpData.Clear();

                string connString = string.Format("Provider = Microsoft.Jet.OleDb.4.0; Data Source = {0}; Extended Properties = 'text; HDR = {1}; FMT = Delimited ({2})' ", strFilePath, strHeader, strDlmtd);

                OleDbConnection oledbConn = new OleDbConnection(connString);

                oledbConn.Open();

                OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}]", strCSVPath), oledbConn);

                OleDbDataAdapter oleda = new OleDbDataAdapter();

                oleda.SelectCommand = cmd;

                oleda.Fill(GlobalVariable.oDTImpData);

                oledbConn.Close();

                return true;
            }
            catch (Exception ex)
            {
                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.errorAppend(string.Format("Error retrieving data from CSV File ({0}). Description : {1} ", strCSVPath, ex.Message.ToString()));

                return false;
            }

        }
        public static bool decryptPGP(string strEncrytdFilePath, string strPrivateKeyPath, string strPassKey, string strDcryptdFilePath)
        {


            try
            {
                PGPLib pgp = new PGPLib();

                string originalFileName = pgp.DecryptFile(strEncrytdFilePath,
                                          strPrivateKeyPath,
                                          strPassKey,
                                          strDcryptdFilePath);

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.errorAppend(string.Format("Error Decrypting Payroll File ({0}). Description : {1} ", Path.GetFileName(strEncrytdFilePath), ex.Message.ToString()));
                return false;
            }
        }
    }
}
