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
using ClosedXML.Excel;
using DidX.BouncyCastle.Crypto.Tls;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;


namespace SAPB1iService
{
    class ImportUserDefinedDocuments
    {
        private static DateTime dteStart;   
        private static string strTransType;
        private static string strMsgBod;
        private static string strPostDocEnt, strPostDocNum;

        private static DataTable oDTBPPurchase, oDTBPSales, oDTEWT;
        public static void _ImportUserDefinedDocuments()
        {
            importFromFile();
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

                        if (strFile.Contains("Purchases") || strFile.Contains("Sales"))
                        {
                            dteStart = DateTime.Now;

                            if (importDIAPIPostDocumentFExcel(strFile))
                                strStatus = "S";
                            else
                                strStatus = "E";

                            TransferFile.transferProcFiles("Import", strStatus, Path.GetFileName(strFile));

                            GlobalFunction.sendAlert(strStatus, "Import", strMsgBod, GlobalVariable.oObjectType, GlobalVariable.oCompany.GetNewObjectKey().ToString());

                            ExportUserDefinedObjects.exportBIR2307();

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
        private static bool importDIAPIPostDocumentFExcel(string strFile)
        {
            string strQuery;

            SAPbobsCOM.Recordset oRecordset;

            try
            {
                if (GlobalFunction.importXLSX(Path.GetFullPath(strFile), "NO", "Sheet1"))
                {
                    if (strFile.Contains("Purchases"))
                    {
                        strQuery = string.Format("SELECT \"DocEntry\" " +
                                                 "FROM \"OPCH\" " +
                                                 "WHERE \"U_FileName\" = '{0}' ", GlobalVariable.strFileName);

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (oRecordset.RecordCount > 0)
                        {
                            SystemFunction.transHandler("Import", "Documents - IEMOP Purchases", "19", GlobalVariable.strFileName, "", "", dteStart, "E", "-999", string.Format("File {0} already uploaded.", GlobalVariable.strFileName));
                            return false;
                        }
                        else
                        {
                            if (!(importPurchases()))
                                return false;
                            else
                                return true;
                        }
                    }
                    else if (strFile.Contains("Sales"))
                    {
                        strQuery = string.Format("SELECT \"DocEntry\" " +
                                                 "FROM \"OINV\" " +
                                                 "WHERE \"U_FileName\" = '{0}' ", GlobalVariable.strFileName);

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (oRecordset.RecordCount > 0)
                        {
                            SystemFunction.transHandler("Import", "Documents - IEMOP Sales", "13", GlobalVariable.strFileName, "", "", dteStart, "E", "-999", string.Format("File {0} already uploaded.", GlobalVariable.strFileName));
                            return false;
                        }
                        else
                        {

                            if (!(importSales()))
                                return false;
                            else
                                return true;
                        }
                    }
                }
                else
                {
                    return false;
                }

                GC.Collect();

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                return false;
            }
        }
        private static bool importPurchases()
        {
            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Documents oDocuments;

            string strSTLID = "", strNumAtCard = "", strCardCode, strRemarks = "", strWTCode, strDscrption, strQuery;

            string strVatPur = "", strTaxVatPur = "", strZroPur = "", strTaxZroPur = "", strZroEcoPur = "", strTaxZroEcoPur = "", strVatOnPur = "", strTaxVatOnPur = "";

            string strBIR2307File, strBllDate, strUpDate;


            DateTime dteDoc = Convert.ToDateTime("01/01/1900");
            DateTime dteDue = Convert.ToDateTime("01/01/1900");
            DateTime dteTax = Convert.ToDateTime("01/01/1900");

            DateTime dteDocV = Convert.ToDateTime("01/01/1900");
            DateTime dteDueV = Convert.ToDateTime("01/01/1900");
            DateTime dteTaxV = Convert.ToDateTime("01/01/1900");

            DateTime dteBllDate = Convert.ToDateTime("01/01/1900");
            DateTime dteUpDate = Convert.ToDateTime("01/01/1900");

            DataTable oDTSTLID, oDTBLLID, oDTInvoice;

            bool blBPExist, blTempErr, blWithErr, blEWT;

            double dblVatPur, dblZroPur, dblZroEcoPur, dblVatOnPur, dblEWT, dblWTBaseAmt, dblWTAmt;

            int intCtr = 0, intSeries;

            try
            {
                strTransType = "Documents - IEMOP Purchases";

                blBPExist = true;
                blTempErr = false;

                strQuery = string.Format("SELECT \"U_VatSlePur\", \"U_TaxVatSlePur\", \"U_ZerSlePur\", \"U_TaxZerSlePur\", \"U_ZerEcoSlePur\", " +
                                         "       \"U_TaxZerEcoSlePur\", \"U_VatOnSlePur\", \"U_TaxVatOnSlePur\" " +
                                         "FROM \"@FTIEMOPGL\" " +
                                         "WHERE \"Code\" = 'AP' ");

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (!(oRecordset.RecordCount > 0))
                {

                    GlobalVariable.intErrNum = -999;
                    GlobalVariable.strErrMsg = "IEMOP GL Account Setup is missing.";

                    strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    return false;

                }
                else
                {
                    strVatPur = oRecordset.Fields.Item("U_VatSlePur").Value.ToString();
                    strTaxVatPur = oRecordset.Fields.Item("U_TaxVatSlePur").Value.ToString();
                    strZroPur = oRecordset.Fields.Item("U_ZerSlePur").Value.ToString();
                    strTaxZroPur = oRecordset.Fields.Item("U_TaxZerSlePur").Value.ToString();
                    strZroEcoPur = oRecordset.Fields.Item("U_ZerEcoSlePur").Value.ToString();
                    strTaxZroEcoPur = oRecordset.Fields.Item("U_TaxZerEcoSlePur").Value.ToString();
                    strVatOnPur = oRecordset.Fields.Item("U_VatOnSlePur").Value.ToString();
                    strTaxVatOnPur = oRecordset.Fields.Item("U_TaxVatOnSlePur").Value.ToString();
                }

                strQuery = string.Format("SELECT NNM1.\"Series\",NNM1.\"SeriesName\", GL.\"U_SeriesName\" FROM \"NNM1\" " +
                                            "INNER JOIN \"@FTIEMOPGL\" GL ON NNM1.\"SeriesName\" = GL.\"U_SeriesName\" " +
                                            "WHERE GL.\"Code\" = 'AP' AND NNM1.\"ObjectCode\" = '18'");
                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                    intSeries = Convert.ToInt32(oRecordset.Fields.Item("Series").Value.ToString());
                else
                    intSeries = 0;

                for (int intRowH = 0; intRowH <= 3; intRowH++)
                {
                    strSTLID = GlobalVariable.oDTImpData.Rows[intRowH][0].ToString();

                    if (strSTLID == "TRANSACTION_NO")
                        strNumAtCard = GlobalVariable.oDTImpData.Rows[intRowH][1].ToString();

                    if (strSTLID == "Posting Date")
                    {
                        if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString()))
                            dteDoc = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString());
                        else
                            blTempErr = true;

                        if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][2].ToString()))
                            dteDocV = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][2].ToString());
                        else
                            blTempErr = true;

                        if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][3].ToString()))
                            dteBllDate = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][3].ToString());
                        else
                            blTempErr = true;

                    }

                    if (strSTLID == "Due Date")
                    {
                        if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString()))
                            dteDue = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString());
                        else
                            blTempErr = true;

                        if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][2].ToString()))
                            dteDueV = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][2].ToString());
                        else
                            blTempErr = true;

                        if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][3].ToString()))
                            dteUpDate = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][3].ToString());
                        else
                            blTempErr = true;
                    }

                    if (strSTLID == "Document Date")
                    {
                        if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString()))
                            dteTax = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString());
                        else
                            blTempErr = true;

                        if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][2].ToString()))
                            dteTaxV = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][2].ToString());
                        else
                            blTempErr = true;
                    }
                }

                if (blTempErr == true)
                {
                    GlobalVariable.intErrNum = -998;
                    GlobalVariable.strErrMsg = "Template Error. Please check dates.";

                    strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    return false;
                }

                try
                {
                    //oDTSTLID = GlobalVariable.oDTImpData.DefaultView.ToTable(true, GlobalVariable.oDTImpData.Columns[0].ColumnName);
                    strQuery = string.Format("({0} <> '0' OR {1} <> '0' OR {2} <> '0' OR {3} <> '0') AND " +
                                             "({4} <> 'TRANSACTION_NO' AND {4} <> 'Posting Date' AND " +
                                             " {4} <> 'Due Date' AND {4} <> 'Document Date' AND {4} <> 'STL ID') ",  GlobalVariable.oDTImpData.Columns[10].ColumnName,
                                                                                                                     GlobalVariable.oDTImpData.Columns[11].ColumnName,
                                                                                                                     GlobalVariable.oDTImpData.Columns[12].ColumnName,
                                                                                                                     GlobalVariable.oDTImpData.Columns[14].ColumnName,
                                                                                                                     GlobalVariable.oDTImpData.Columns[0].ColumnName);

                    oDTSTLID = GlobalVariable.oDTImpData.Select(strQuery).CopyToDataTable().DefaultView.ToTable(true, GlobalVariable.oDTImpData.Columns[0].ColumnName);
                }
                catch (Exception ex)
                {
                    GlobalVariable.intErrNum = -991;
                    GlobalVariable.strErrMsg = string.Format("Error Processing Purchases File {0}. Purchases Details Not Found.", GlobalVariable.strFileName);

                    strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    return false;
                }

                if (oDTSTLID.Rows.Count > 0)
                {
                    initDataTable();

                    if (!(GlobalVariable.oCompany.InTransaction))
                        GlobalVariable.oCompany.StartTransaction();

                    for (int intRowID = 0; intRowID <= oDTSTLID.Rows.Count - 1; intRowID++)
                    {

                        strSTLID = oDTSTLID.Rows[intRowID][0].ToString();

                        strQuery = string.Format("SELECT OCRD.\"CardCode\" " +
                                                    "FROM \"@FTIEMOP1\"  IEMOP1 INNER JOIN \"@FTOIEMOP\" OIEMOP ON IEMOP1.\"Code\" = OIEMOP.\"Code\"  " +
                                                    "                           INNER JOIN OCRD ON OIEMOP.\"Code\" = OCRD.\"CardCode\" " +
                                                    "WHERE IEMOP1.\"U_STLID\" = '{0}' AND OCRD.\"CardType\" = 'S' ", strSTLID);
                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (!(oRecordset.RecordCount > 0))
                        {
                            blBPExist = false;

                            oDTBLLID = GlobalVariable.oDTImpData.Select(string.Format("{0} = '{1}' ", GlobalVariable.oDTImpData.Columns[0].ColumnName, strSTLID)).CopyToDataTable().DefaultView.ToTable(true, GlobalVariable.oDTImpData.Columns[1].ColumnName);

                            for (int intBID = 0; intBID <= oDTBLLID.Rows.Count - 1; intBID++)
                                oDTBPPurchase.Rows.Add("", "", strSTLID, oDTBLLID.Rows[intBID][0].ToString());

                            continue;
                        }
                        else
                            strCardCode = oRecordset.Fields.Item("CardCode").Value.ToString();

                        intCtr = 0;
                        dblWTAmt = 0;
                        dblWTBaseAmt = 0;
                        blWithErr = false;
                        blEWT = false;

                        if (blBPExist == true)
                        {
                            strQuery = string.Format("{0} = '{1}' AND ({2} <> '0' OR {3} <> '0' OR {4} <> '0' OR {5} <> '0') ", GlobalVariable.oDTImpData.Columns[0].ColumnName,
                                                                                                                                strSTLID,
                                                                                                                                GlobalVariable.oDTImpData.Columns[10].ColumnName,
                                                                                                                                GlobalVariable.oDTImpData.Columns[11].ColumnName,
                                                                                                                                GlobalVariable.oDTImpData.Columns[12].ColumnName,
                                                                                                                                GlobalVariable.oDTImpData.Columns[14].ColumnName);

                            oDTInvoice = GlobalVariable.oDTImpData.Select(strQuery).CopyToDataTable().DefaultView.ToTable();

                            if (oDTInvoice.Rows.Count > 0)
                            {

                                GlobalFunction.getObjType(18);

                                oDocuments = null;
                                oDocuments = (SAPbobsCOM.Documents)GlobalVariable.oCompany.GetBusinessObject(GlobalVariable.oObjectType);

                                if (intSeries != 0)
                                    oDocuments.Series = intSeries;

                                oDocuments.DocType = BoDocumentTypes.dDocument_Service;
                                oDocuments.CardCode = strCardCode;
                                oDocuments.DocDate = dteDoc;
                                oDocuments.DocDueDate = dteDue;
                                oDocuments.TaxDate = dteTax;
                                oDocuments.NumAtCard = strNumAtCard.Replace("-", "");

                                oDocuments.UserFields.Fields.Item("U_FileName").Value = GlobalVariable.strFileName;
                                oDocuments.UserFields.Fields.Item("U_WithBase").Value = "Y";
                                oDocuments.UserFields.Fields.Item("U_TransType").Value = "IEMOP";
                                oDocuments.UserFields.Fields.Item("U_STLID").Value = strSTLID;
                                oDocuments.UserFields.Fields.Item("U_BllDate").Value = dteBllDate;
                                oDocuments.UserFields.Fields.Item("U_UpDate").Value = dteUpDate;

                                for (int intRowInv = 0; intRowInv <= oDTInvoice.Rows.Count - 1; intRowInv++)
                                {
                                    strRemarks = oDTInvoice.Rows[intRowInv][15].ToString().Replace(",", "");
                                    strDscrption = oDTInvoice.Rows[intRowInv][1].ToString() + "_" + oDTInvoice.Rows[intRowInv][2].ToString();

                                    dblVatPur = Math.Abs(Convert.ToDouble(oDTInvoice.Rows[intRowInv][10].ToString()));
                                    dblZroPur = Math.Abs(Convert.ToDouble(oDTInvoice.Rows[intRowInv][11].ToString()));
                                    dblZroEcoPur = Math.Abs(Convert.ToDouble(oDTInvoice.Rows[intRowInv][12].ToString()));
                                    dblEWT = Math.Abs(Convert.ToDouble(oDTInvoice.Rows[intRowInv][14].ToString()));

                                    /***** Vatable Purchases *****/

                                    if (dblVatPur != 0)
                                    {
                                        if (intCtr > 0)
                                            oDocuments.Lines.Add();

                                        oDocuments.Lines.ItemDescription = strDscrption;
                                        oDocuments.Lines.AccountCode = strVatPur;
                                        oDocuments.Lines.UnitPrice = dblVatPur;
                                        oDocuments.Lines.VatGroup = strTaxVatPur;
                                        oDocuments.DiscountPercent = 0;

                                        if (dblEWT != 0)
                                        {
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tYES;
                                            dblWTBaseAmt = dblWTBaseAmt + dblVatPur;
                                        }
                                        else
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tNO;

                                        oDocuments.Lines.UserFields.Fields.Item("U_LegalText").Value = oDTInvoice.Rows[intRowInv][16].ToString();
                                        oDocuments.Lines.UserFields.Fields.Item("U_TransType").Value = "VP";

                                        intCtr++;
                                    }

                                    /***** Zero Rated Purchases *****/
                                    if (dblZroPur != 0)
                                    {
                                        if (intCtr > 0)
                                            oDocuments.Lines.Add();

                                        oDocuments.Lines.ItemDescription = strDscrption;
                                        oDocuments.Lines.AccountCode = strZroPur;
                                        oDocuments.Lines.UnitPrice = dblZroPur;
                                        oDocuments.Lines.VatGroup = strTaxZroPur;
                                        oDocuments.DiscountPercent = 0;

                                        if (dblEWT != 0)
                                        {
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tYES;
                                            dblWTBaseAmt = dblWTBaseAmt + dblZroPur;
                                        }
                                        else
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tNO;

                                        oDocuments.Lines.UserFields.Fields.Item("U_BudgetCode").Value = "B5180100078";
                                        oDocuments.Lines.UserFields.Fields.Item("U_LegalText").Value = oDTInvoice.Rows[intRowInv][16].ToString();
                                        oDocuments.Lines.UserFields.Fields.Item("U_TransType").Value = "ZP";

                                        intCtr++;
                                    }

                                    /***** Zero Rated Econzone Purchases *****/
                                    if (dblZroEcoPur != 0)
                                    {
                                        if (intCtr > 0)
                                            oDocuments.Lines.Add();

                                        oDocuments.Lines.ItemDescription = strDscrption;
                                        oDocuments.Lines.AccountCode = strZroEcoPur;
                                        oDocuments.Lines.UnitPrice = dblZroEcoPur;
                                        oDocuments.Lines.VatGroup = strTaxZroEcoPur;
                                        oDocuments.DiscountPercent = 0;

                                        if (dblEWT != 0)
                                        {
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tYES;
                                            dblWTBaseAmt = dblWTBaseAmt + dblZroEcoPur;
                                        }
                                        else
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tNO;

                                        oDocuments.Lines.UserFields.Fields.Item("U_BudgetCode").Value = "B5180200079";
                                        oDocuments.Lines.UserFields.Fields.Item("U_LegalText").Value = oDTInvoice.Rows[intRowInv][16].ToString();
                                        oDocuments.Lines.UserFields.Fields.Item("U_TransType").Value = "EP";

                                        intCtr++;
                                    }

                                    /***** EWT *****/
                                    if (dblEWT != 0)
                                    {
                                        blEWT = true;
                                        dblWTAmt = dblWTAmt + dblEWT;
                                    }
                                }

                                if (dblWTBaseAmt != 0)
                                {
                                    strQuery = string.Format("SELECT \"WTCode\" FROM OCRD WHERE \"CardCode\" = '{0}' ", strCardCode);

                                    oRecordset = null;
                                    oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    oRecordset.DoQuery(strQuery);

                                    if (oRecordset.RecordCount > 0)
                                    {
                                        strWTCode = oRecordset.Fields.Item("WTCode").Value.ToString();

                                        if (!(string.IsNullOrEmpty(strWTCode)))
                                        {

                                            oDocuments.WithholdingTaxData.WTCode = oRecordset.Fields.Item("WTCode").Value.ToString();
                                            oDocuments.WithholdingTaxData.TaxableAmount = Math.Abs(dblWTBaseAmt);
                                            oDocuments.WithholdingTaxData.WTAmount = Math.Abs(dblWTAmt);
                                        }
                                        else
                                        { 
                                            GlobalVariable.intErrNum = -996;
                                            GlobalVariable.strErrMsg = string.Format("Business Partner {0} Withholding Tax Setup not exist. Please check BP Master Data and IEMOP Business Partner Setup.", strCardCode);

                                            strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                            SystemFunction.transHandler("Import", "Documents", GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        GlobalVariable.intErrNum = -996;
                                        GlobalVariable.strErrMsg = string.Format("Business Partner {0} Withholding Tax Setup not exist. Please check BP Master Data and IEMOP Business Partner Setup.", strCardCode);

                                        strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                        SystemFunction.transHandler("Import", "Documents", GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                        return false;

                                    }
                                }

                                oDocuments.Comments = strRemarks;

                                if (oDocuments.Add() != 0)
                                {
                                    GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                    strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} - {4}", strTransType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg, strCardCode);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), strCardCode + " - " + GlobalVariable.strErrMsg);

                                    return false;
                                }
                                else
                                { 
                                    if (!(ImportUserDefinedPayments.importFromObject(46, GlobalVariable.oCompany.GetNewObjectKey().ToString(), dteDocV, dteDueV, dteTaxV)))
                                    {
                                        strMsgBod = string.Format("Error Posting Outgoing Payment for {0} in file {1}.\rError Code: {2}\rDescription: {3} ", strCardCode, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), strCardCode + " - " + GlobalVariable.strErrMsg);

                                        return false;
                                    }
                                }
                            }
                        }

                        strQuery = string.Format("UPDATE \"@FTISSP\" SET \"U_NumAtCard\" = '{0}' WHERE \"Code\" = '2' ", strNumAtCard.Replace("-", ""));
                        if (!(SystemFunction.executeQuery(strQuery)))
                        {

                            GlobalVariable.intErrNum = -899;
                            GlobalVariable.strErrMsg = string.Format("Error updating Intgeration Setup for BIR 2307.");

                            SystemFunction.transHandler("Crystal Report", strTransType, "", "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            return false;
                        }
                    }

                    if (blBPExist == false)
                    {
                        GlobalVariable.intErrNum = -997;
                        GlobalVariable.strErrMsg = "Business Partner Setup not exist. Please check BP Master Data and IEMOP Business Partner Setup.";

                        strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        exportExcelBP("Purchase", oDTBPPurchase);

                        return false;

                    }

                    strMsgBod = string.Format("Successfully Posted {0}.", GlobalVariable.strFileName);

                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "S", "0", strMsgBod);

                    if (GlobalVariable.oCompany.InTransaction)
                        GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                }

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
        private static bool importSales()
        {
            
            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Documents oDocuments;

            string strPostDocNum = "";

            string strSTLID, strBLLID, strNumAtCard = "", strCardCode, strRemarks = "", strDscrption, strBllPriod, strQuery;

            string strCardName, strAddress, strSupplier, strTIN, strCardType, strECVatGroup;

            string strVatSal = "", strTaxVatSal = "", strZroSal = "", strTaxZroSal = "", strZroEcoSal = "", strTaxZroEcoSal = "", strVatOnSal = "", strTaxVatOnSal = "";

            DateTime dteDoc = Convert.ToDateTime("01/01/1900");
            DateTime dteDue = Convert.ToDateTime("01/01/1900");
            DateTime dteTax = Convert.ToDateTime("01/01/1900");

            DataTable oDTSTLID, oDTInvoice;

            bool blBPExist, blTempErr;

            double dblVatSal, dblZroSal, dblZroEcoSal, dblVatOnSal, dblEWT, dblWTBaseAmt, dblWTAmt;

            int intCtr = 0, intSeries;

            try
            {
                strTransType = "Documents - IEMOP Sales";

                blBPExist = true;
                blTempErr = false;

                strQuery = string.Format("SELECT \"U_VatSlePur\", \"U_TaxVatSlePur\", \"U_ZerSlePur\", \"U_TaxZerSlePur\", \"U_ZerEcoSlePur\", " +
                                         "       \"U_TaxZerEcoSlePur\", \"U_VatOnSlePur\", \"U_TaxVatOnSlePur\" " +
                                         "FROM \"@FTIEMOPGL\" " +
                                         "WHERE \"Code\" = 'AR' ");

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (!(oRecordset.RecordCount > 0))
                {

                    GlobalVariable.intErrNum = -995;
                    GlobalVariable.strErrMsg = "IEMOP GL Account Setup is missing.";

                    strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    return false;

                }
                else
                {
                    strVatSal = oRecordset.Fields.Item("U_VatSlePur").Value.ToString();
                    strTaxVatSal = oRecordset.Fields.Item("U_TaxVatSlePur").Value.ToString();
                    strZroSal = oRecordset.Fields.Item("U_ZerSlePur").Value.ToString();
                    strTaxZroSal = oRecordset.Fields.Item("U_TaxZerSlePur").Value.ToString();
                    strZroEcoSal = oRecordset.Fields.Item("U_ZerEcoSlePur").Value.ToString();
                    strTaxZroEcoSal = oRecordset.Fields.Item("U_TaxZerEcoSlePur").Value.ToString();
                    strVatOnSal = oRecordset.Fields.Item("U_VatOnSlePur").Value.ToString();
                    strTaxVatOnSal = oRecordset.Fields.Item("U_TaxVatOnSlePur").Value.ToString();
                }

                strQuery = string.Format("SELECT NNM1.\"Series\",NNM1.\"SeriesName\", GL.\"U_SeriesName\" FROM \"NNM1\" " +
                                            "INNER JOIN \"@FTIEMOPGL\" GL ON NNM1.\"SeriesName\" = GL.\"U_SeriesName\" " +
                                            "WHERE GL.\"Code\" = 'AR' AND NNM1.\"ObjectCode\" = '13'");
                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                    intSeries = Convert.ToInt32(oRecordset.Fields.Item("Series").Value.ToString());
                else
                    intSeries = 0;

                for (int intRowH = 0; intRowH <= 3; intRowH++)
                {
                    strSTLID = GlobalVariable.oDTImpData.Rows[intRowH][0].ToString();
                    strBLLID = GlobalVariable.oDTImpData.Rows[intRowH][1].ToString();

                    if (strSTLID == "TRANSACTION_NO")
                        strNumAtCard = GlobalVariable.oDTImpData.Rows[intRowH][1].ToString();

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
                }

                if (blTempErr == true)
                {
                    GlobalVariable.intErrNum = -994;
                    GlobalVariable.strErrMsg = "Template Error. Please check dates.";

                    strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    return false;
                }

                try
                {
                    strQuery = string.Format("({0} > '0' OR {1} > '0' OR {2} > '0' OR {3} > '0' OR {4} > '0') AND " +
                                             "({5} <> 'TRANSACTION_NO' AND {5} <> 'Posting Date' AND " +
                                             " {5} <> 'Due Date' AND {5} <> 'Document Date' AND {5} <> 'STL ID') ", GlobalVariable.oDTImpData.Columns[6].ColumnName,
                                                                                                                    GlobalVariable.oDTImpData.Columns[7].ColumnName,
                                                                                                                    GlobalVariable.oDTImpData.Columns[8].ColumnName,
                                                                                                                    GlobalVariable.oDTImpData.Columns[9].ColumnName,
                                                                                                                    GlobalVariable.oDTImpData.Columns[14].ColumnName,
                                                                                                                    GlobalVariable.oDTImpData.Columns[0].ColumnName);

                    oDTSTLID = GlobalVariable.oDTImpData.Select(strQuery)?.CopyToDataTable().DefaultView.ToTable(true, GlobalVariable.oDTImpData.Columns[0].ColumnName, GlobalVariable.oDTImpData.Columns[1].ColumnName);
                    //oDTSTLID = GlobalVariable.oDTImpData.DefaultView.ToTable(true, GlobalVariable.oDTImpData.Columns[0].ColumnName, GlobalVariable.oDTImpData.Columns[1].ColumnName);
                }
                catch (Exception ex)
                {
                    GlobalVariable.intErrNum = -992;
                    GlobalVariable.strErrMsg = string.Format("Error Processing Sales File {0}. Sales Details Not Found. {1}.", GlobalVariable.strFileName);

                    strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    return false;
                }

                if (oDTSTLID.Rows.Count > 0)
                {
                    initDataTable();

                    if (!(GlobalVariable.oCompany.InTransaction))
                        GlobalVariable.oCompany.StartTransaction();

                    for (int intRowID = 0; intRowID <= oDTSTLID.Rows.Count - 1; intRowID++)
                    {
                        strSTLID = oDTSTLID.Rows[intRowID][0].ToString();
                        strBLLID = oDTSTLID.Rows[intRowID][1].ToString();


                        strQuery = string.Format("SELECT OCRD.\"CardCode\", OCRD.\"ECVatGroup\" " +
                                                 "FROM \"@FTIEMOP1\"  IEMOP1 INNER JOIN \"@FTOIEMOP\" OIEMOP ON IEMOP1.\"Code\" = OIEMOP.\"Code\"  " +
                                                 "                           INNER JOIN OCRD ON OIEMOP.\"Code\" = OCRD.\"CardCode\" " +
                                                 "WHERE IEMOP1.\"U_STLID\" = '{0}' AND IEMOP1.\"U_BLLID\" = '{1}' AND OCRD.\"CardType\" = 'C' ", strSTLID, strBLLID);
                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (!(oRecordset.RecordCount > 0))
                        {
                            blBPExist = false;

                            oDTBPSales.Rows.Add("", "", strSTLID, strBLLID);

                            continue;
                        }
                        else
                            strCardCode = oRecordset.Fields.Item("CardCode").Value.ToString();
                            strECVatGroup = oRecordset.Fields.Item("ECVatGroup").Value.ToString();

                        intCtr = 0;
                        dblWTAmt = 0;
                        dblWTBaseAmt = 0;

                        if (blBPExist == true)
                        {
                            strQuery = string.Format(string.Format("{0} = '{1}' AND {2} = '{3}' AND " +
                                                                    "({4} <> '0' OR {5} <> '0' OR {6} <> '0' OR {7} <> '0' OR {8} <> '0') ", GlobalVariable.oDTImpData.Columns[0].ColumnName, 
                                                                                                                                            strSTLID, 
                                                                                                                                            GlobalVariable.oDTImpData.Columns[1].ColumnName, 
                                                                                                                                            strBLLID, 
                                                                                                                                            GlobalVariable.oDTImpData.Columns[6].ColumnName,
                                                                                                                                            GlobalVariable.oDTImpData.Columns[7].ColumnName,
                                                                                                                                            GlobalVariable.oDTImpData.Columns[8].ColumnName,
                                                                                                                                            GlobalVariable.oDTImpData.Columns[9].ColumnName, 
                                                                                                                                            GlobalVariable.oDTImpData.Columns[14].ColumnName));


                            oDTInvoice = GlobalVariable.oDTImpData.Select(strQuery).CopyToDataTable().DefaultView.ToTable();

                            if (oDTInvoice.Rows.Count > 0)
                            {
                                GlobalFunction.getObjType(13);

                                oDocuments = null;
                                oDocuments = (SAPbobsCOM.Documents)GlobalVariable.oCompany.GetBusinessObject(GlobalVariable.oObjectType);

                                if (intSeries != 0)
                                    oDocuments.Series = intSeries;

                                oDocuments.DocType = BoDocumentTypes.dDocument_Service;
                                oDocuments.CardCode = strCardCode;
                                oDocuments.DocDate = dteDoc;
                                oDocuments.DocDueDate = dteDue;
                                oDocuments.TaxDate = dteTax;
                                oDocuments.NumAtCard = strNumAtCard;

                                strBllPriod = dteDoc.ToString("MMM") + "-" + dteDoc.Year.ToString();

                                oDocuments.UserFields.Fields.Item("U_FileName").Value = GlobalVariable.strFileName;
                                oDocuments.UserFields.Fields.Item("U_BillingPeriod").Value = strBllPriod;

                                for (int intRowInv = 0; intRowInv <= oDTInvoice.Rows.Count - 1; intRowInv++)
                                {
                                    strRemarks = oDTInvoice.Rows[intRowInv][15].ToString() + "_" + oDTInvoice.Rows[intRowInv][1].ToString();
                                    strDscrption = oDTInvoice.Rows[intRowInv][1].ToString() + "_" + oDTInvoice.Rows[intRowInv][2].ToString();

                                    dblVatSal = Math.Abs(Convert.ToDouble(oDTInvoice.Rows[intRowInv][6].ToString()));
                                    dblZroSal = Math.Abs(Convert.ToDouble(oDTInvoice.Rows[intRowInv][7].ToString()));
                                    dblZroEcoSal = Math.Abs(Convert.ToDouble(oDTInvoice.Rows[intRowInv][8].ToString()));
                                    dblVatOnSal = Math.Abs(Convert.ToDouble(oDTInvoice.Rows[intRowInv][9].ToString()));
                                    dblEWT = Math.Abs(Convert.ToDouble(oDTInvoice.Rows[intRowInv][14].ToString()));

                                    /***** Vatable Sales *****/

                                    if (dblVatSal != 0)
                                    {
                                        if (intCtr > 0)
                                            oDocuments.Lines.Add();

                                        oDocuments.Lines.ItemDescription = strDscrption;
                                        oDocuments.Lines.AccountCode = strVatSal;
                                        oDocuments.Lines.UnitPrice = dblVatSal;
                                        oDocuments.Lines.VatGroup = strECVatGroup;
                                        oDocuments.DiscountPercent = 0;

                                        if (dblEWT != 0)
                                        {
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tYES;
                                            dblWTBaseAmt = dblWTBaseAmt + dblVatSal;
                                        }
                                        else
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tNO;

                                        oDocuments.Lines.UserFields.Fields.Item("U_TransType").Value = "VS";

                                        intCtr++;
                                    }

                                    /***** Zero Rated Sales *****/
                                    if (dblZroSal != 0)
                                    {
                                        if (intCtr > 0)
                                            oDocuments.Lines.Add();

                                        oDocuments.Lines.ItemDescription = strDscrption;
                                        oDocuments.Lines.AccountCode = strZroSal;
                                        oDocuments.Lines.UnitPrice = dblZroSal;
                                        oDocuments.Lines.VatGroup = strECVatGroup;
                                        oDocuments.DiscountPercent = 0;

                                        if (dblEWT != 0)
                                        {
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tYES;
                                            dblWTBaseAmt = dblWTBaseAmt + dblZroSal;
                                        }
                                        else
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tNO;

                                        oDocuments.Lines.UserFields.Fields.Item("U_TransType").Value = "ZS";

                                        intCtr++;
                                    }

                                    /***** Zero Rated Econzone Sales *****/
                                    if (dblZroEcoSal != 0)
                                    {
                                        if (intCtr > 0)
                                            oDocuments.Lines.Add();

                                        oDocuments.Lines.ItemDescription = strDscrption;
                                        oDocuments.Lines.AccountCode = strZroEcoSal;
                                        oDocuments.Lines.UnitPrice = dblZroEcoSal;
                                        oDocuments.Lines.VatGroup = strECVatGroup;
                                        oDocuments.DiscountPercent = 0;

                                        if (dblEWT != 0)
                                        {
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tYES;
                                            dblWTBaseAmt = dblWTBaseAmt + dblZroEcoSal;
                                        }
                                        else
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tNO;

                                        oDocuments.Lines.UserFields.Fields.Item("U_TransType").Value = "ES";

                                        intCtr++;
                                    }

                                    /***** VAT On Sales *****/
                                    if (dblVatOnSal != 0)
                                    {
                                        if (intCtr > 0)
                                            oDocuments.Lines.Add();

                                        oDocuments.Lines.ItemDescription = strDscrption;
                                        oDocuments.Lines.AccountCode = strVatOnSal;
                                        oDocuments.Lines.UnitPrice = dblVatOnSal;
                                        oDocuments.Lines.VatGroup = strECVatGroup;
                                        oDocuments.Lines.WTLiable = BoYesNoEnum.tNO;
                                        oDocuments.DiscountPercent = 0;

                                        oDocuments.Lines.UserFields.Fields.Item("U_TransType").Value = "VOS";

                                        //CHANGE REQUEST Auto-populate relevant BIR Add On fields 20/06/2024

                                        strQuery = string.Format("SELECT OCRD.\"CardCode\",OCRD.\"CardName\",OCRD.\"U_TIN1\",OCRD.\"CardType\", CRD1.\"Street\" FROM OCRD INNER JOIN CRD1 ON OCRD.\"CardCode\" = CRD1.\"CardCode\" WHERE OCRD.\"CardCode\" = '{0}'", strCardCode);

                                        oRecordset = null;
                                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRecordset.DoQuery(strQuery);

                                        strCardCode = oRecordset.Fields.Item("CardCode").Value.ToString();
                                        strCardName = oRecordset.Fields.Item("CardName").Value.ToString();
                                        strAddress = oRecordset.Fields.Item("Street").Value.ToString();
                                        strTIN = oRecordset.Fields.Item("U_TIN1").Value.ToString();
                                        strCardType = oRecordset.Fields.Item("CardType").Value.ToString();
                                        oDocuments.Lines.UserFields.Fields.Item("U_xWTCode").Value = strTaxVatOnSal;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xWTVendor").Value = strCardCode;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xTaxbleAmnt").Value = dblVatSal;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xSupplierName").Value = strCardName;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xAddress").Value = strAddress;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xTINnumber").Value = strTIN;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xCardType").Value = strCardType;


                                        intCtr++;
                                    }

                                    //CHANGE REQUEST 2nd line added if Zero Rated Econzone Sales is >0

                                    if (dblZroEcoSal != 0)
                                    {
                                        if (intCtr > 0)
                                            oDocuments.Lines.Add();

                                        oDocuments.Lines.ItemDescription = strDscrption;
                                        oDocuments.Lines.AccountCode = strVatOnSal;
                                        oDocuments.Lines.UnitPrice = 0;
                                        oDocuments.Lines.VatGroup = strECVatGroup;
                                        oDocuments.Lines.WTLiable = BoYesNoEnum.tNO;
                                        oDocuments.DiscountPercent = 0;

                                        oDocuments.Lines.UserFields.Fields.Item("U_TransType").Value = "ES";                                        

                                        strQuery = string.Format("SELECT OCRD.\"CardCode\",OCRD.\"CardName\",OCRD.\"U_TIN1\",OCRD.\"CardType\", CRD1.\"Street\" FROM OCRD INNER JOIN CRD1 ON OCRD.\"CardCode\" = CRD1.\"CardCode\" WHERE OCRD.\"CardCode\" = '{0}'", strCardCode);

                                        oRecordset = null;
                                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRecordset.DoQuery(strQuery);

                                        strCardCode = oRecordset.Fields.Item("CardCode").Value.ToString();
                                        strCardName = oRecordset.Fields.Item("CardName").Value.ToString();
                                        strAddress = oRecordset.Fields.Item("Street").Value.ToString();
                                        strTIN = oRecordset.Fields.Item("U_TIN1").Value.ToString();
                                        strCardType = oRecordset.Fields.Item("CardType").Value.ToString();
                                        oDocuments.Lines.UserFields.Fields.Item("U_xWTCode").Value = strTaxZroEcoSal;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xWTVendor").Value = strCardCode;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xTaxbleAmnt").Value = dblZroEcoSal;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xSupplierName").Value = strCardName;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xAddress").Value = strAddress;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xTINnumber").Value = strTIN;
                                        oDocuments.Lines.UserFields.Fields.Item("U_xCardType").Value = strCardType;


                                        intCtr++;
                                    }


                                    /***** EWT *****/
                                    if (dblEWT != 0)
                                        dblWTAmt = dblWTAmt + dblEWT;

                                }

                                //COMMENTED DATE: -20 / 06 / 2024 CHANGE REQUEST EOPT


                                //if (dblWTBaseAmt != 0)
                                //{
                                //    strQuery = string.Format("SELECT \"WTCode\" FROM OCRD WHERE \"CardCode\" = '{0}' ", strCardCode);

                                //    oRecordset = null;
                                //    oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                //    oRecordset.DoQuery(strQuery);

                                //    if (!(oRecordset.RecordCount > 0))
                                //    {
                                //        GlobalVariable.intErrNum = -992;
                                //        GlobalVariable.strErrMsg = string.Format("Business Partner {0} Withholding Tax Setup not exist. Please check BP Master Data and IEMOP Business Partner Setup.", strCardCode);

                                //        strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                //        SystemFunction.transHandler("Import", "Documents", GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                //        return false;

                                //    }
                                //    else
                                //    {
                                //        oDocuments.WithholdingTaxData.WTCode = oRecordset.Fields.Item("WTCode").Value.ToString();
                                //        oDocuments.WithholdingTaxData.TaxableAmount = Math.Abs(dblWTBaseAmt);
                                //        oDocuments.WithholdingTaxData.WTAmount = Math.Abs(dblWTAmt);
                                //    }
                                //}



                                oDocuments.Comments = strRemarks;

                                if (oDocuments.Add() != 0)
                                {
                                    GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                    strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    SystemFunction.transHandler("Import", "Documents", GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), strCardCode + " - " + GlobalVariable.strErrMsg);

                                    return false;
                                }
                            }
                        }
                        
                    }

                    if (blBPExist == false)
                    {
                        GlobalVariable.intErrNum = -993;
                        GlobalVariable.strErrMsg = "Business Partner Setup not exist. Please check BP Master Data and IEMOP Business Partner Setup.";

                        strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        exportExcelBP("Sales", oDTBPSales);

                        return false;

                    }

                    strMsgBod = string.Format("Successfully Posted {0}.", GlobalVariable.strFileName);

                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, GlobalVariable.oCompany.GetNewObjectKey(), strPostDocNum, dteStart, "S", "0", strMsgBod);

                    if (GlobalVariable.oCompany.InTransaction)
                        GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                }
                
                GC.Collect();

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

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
        private static bool exportExcelBP(string strType, DataTable oDTBP)
        {
            string strFilePath, strFileName;

            try
            {

                using (XLWorkbook wb = new XLWorkbook())
                {

                    if (strType == "Sales")
                        strFileName = string.Format("{0}_Customer IEMOP Business Partner.xlsx", GlobalVariable.strCompany);
                    else
                        strFileName = string.Format("{0}_Vendor IEMOP Business Partner.xlsx", GlobalVariable.strCompany);

                    var ws = wb.Worksheets.Add("Sheet1");
                    ws.Cell(1, 1).InsertData(oDTBP.Rows);
                    ws.Columns().AdjustToContents();

                    strFilePath = GlobalVariable.strExpPath + strFileName;

                    if (File.Exists(strFilePath))
                        File.Delete(strFilePath);

                    wb.SaveAs(strFilePath);

                }

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                return false;
            }
        }        
        private static void initDataTable()
        {
            oDTBPPurchase = new DataTable("IEMOP BP Purchase");
            oDTBPPurchase.Columns.Add("CardCode", typeof(System.String));
            oDTBPPurchase.Columns.Add("CardName", typeof(System.String));
            oDTBPPurchase.Columns.Add("STLID", typeof(System.String));
            oDTBPPurchase.Columns.Add("BLLID", typeof(System.String));
            oDTBPPurchase.Clear();

            oDTBPSales = new DataTable("IEMOP BP Sales");
            oDTBPSales.Columns.Add("CardCode", typeof(System.String));
            oDTBPSales.Columns.Add("CardName", typeof(System.String));
            oDTBPSales.Columns.Add("STLID", typeof(System.String));
            oDTBPSales.Columns.Add("BLLID", typeof(System.String));
            oDTBPSales.Clear();
           
            oDTBPSales.Rows.Add("CardCode", "CardName", "STLID", "BLLID");
            oDTBPPurchase.Rows.Add("CardCode", "CardName", "STLID", "BLLID");

        }

    }

}
