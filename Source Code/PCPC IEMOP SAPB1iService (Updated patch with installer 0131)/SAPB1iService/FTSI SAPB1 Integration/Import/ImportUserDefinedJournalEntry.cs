using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using SAPbobsCOM;

namespace SAPB1iService
{
    class ImportUserDefinedJournalEntry
    {
        private static DateTime dteStart;
        private static string strTransType;
        public static bool importFromObject(string strPostType, DateTime dtePost, string strDocEntry, string strDocNum, double dblAmount,
                                            string strxCardCode, double dblxOAmount, double dblxZAmount)
        {
            string strQuery, strDefGL, strVatGL, strZerGL;
            string strxOWTCode, strxZWTCode;

            string strAddrss, strTIN;

            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.JournalEntries oJournalEntries;

            try
            {
                dteStart = DateTime.Now;

                strTransType = "Journal Entry - Deferred Tax ReClass";

                strQuery = string.Format("SELECT \"U_VatOnSlePur\", \"U_VATReClassGL\", \"U_ZerReClassGL\",  \"U_ReClsOWTCod\", \"U_ReClsZWTCod\" FROM \"@FTIEMOPGL\" WHERE \"Code\" = 'AR' ");

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                {
                    strDefGL = oRecordset.Fields.Item("U_VatOnSlePur").Value.ToString();
                    strVatGL = oRecordset.Fields.Item("U_VATReClassGL").Value.ToString();
                    strZerGL = oRecordset.Fields.Item("U_ZerReClassGL").Value.ToString();

                    strxOWTCode = oRecordset.Fields.Item("U_ReClsOWTCod").Value.ToString();
                    strxZWTCode = oRecordset.Fields.Item("U_ReClsZWTCod").Value.ToString();

                    if (string.IsNullOrEmpty(strDefGL) || string.IsNullOrEmpty(strVatGL))
                    {
                        GlobalVariable.intErrNum = -699;
                        GlobalVariable.strErrMsg = "AR Invoice IEMOP GL Settings is missing.";

                        return false;
                    }
                    else
                    {
                        strQuery = string.Format("SELECT \"CardName\", \"LicTradNum\", \"Address\", \"CardType\" FROM \"OCRD\" WHERE \"CardCode\" = '{0}' ", strxCardCode);

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        strAddrss = oRecordset.Fields.Item("Address").Value.ToString();
                        strTIN = oRecordset.Fields.Item("LicTradNum").Value.ToString();

                        if (string.IsNullOrEmpty(strTIN) || string.IsNullOrEmpty(strAddrss))
                        {
                            GlobalVariable.intErrNum = -698;
                            GlobalVariable.strErrMsg = string.Format("Please check Address and/or TIN");

                            return false;
                        }

                        oJournalEntries = null;
                        oJournalEntries = (SAPbobsCOM.JournalEntries)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);

                        oJournalEntries.ReferenceDate = dtePost;
                        oJournalEntries.DueDate = dtePost;
                        oJournalEntries.TaxDate = dtePost;

                        oJournalEntries.Memo = string.Format("Deferred Tax ReClass for Payment # {0} ", strDocNum);
                        oJournalEntries.UserFields.Fields.Item("U_ReportEx").Value = "N";

                        oJournalEntries.UserFields.Fields.Item("U_ReClsBseDocEnt").Value = strDocEntry;
                        oJournalEntries.UserFields.Fields.Item("U_ReClsBseDocNum").Value = strDocNum;

                        if (strPostType == "Post")
                        {
                            oJournalEntries.Lines.AccountCode = strDefGL;
                            oJournalEntries.Lines.Debit = dblAmount;
                            oJournalEntries.Lines.Credit = 0;

                            oJournalEntries.Lines.Add();
                            oJournalEntries.Lines.AccountCode = strVatGL;
                            oJournalEntries.Lines.Debit = 0;
                            oJournalEntries.Lines.Credit = dblAmount;

                            oJournalEntries.Lines.UserFields.Fields.Item("U_xWTCode").Value = strxOWTCode;
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xWTVendor").Value = strxCardCode;
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xTaxbleAmnt").Value = dblxOAmount;
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xSupplierName").Value = oRecordset.Fields.Item("CardName").Value.ToString();
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xAddress").Value = strAddrss;
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xTINnumber").Value = strTIN;
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xCardType").Value = oRecordset.Fields.Item("CardType").Value.ToString();

                            if (dblxZAmount > 0)
                            {
                                oJournalEntries.Lines.Add();
                                oJournalEntries.Lines.AccountCode = strZerGL;
                                oJournalEntries.Lines.Debit = 0;
                                oJournalEntries.Lines.Credit = 0;

                                oJournalEntries.Lines.UserFields.Fields.Item("U_xWTCode").Value = strxZWTCode;
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xWTVendor").Value = strxCardCode;
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xTaxbleAmnt").Value = dblxZAmount;
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xSupplierName").Value = oRecordset.Fields.Item("CardName").Value.ToString();
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xAddress").Value = strAddrss;
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xTINnumber").Value = strTIN;
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xCardType").Value = oRecordset.Fields.Item("CardType").Value.ToString();
                            }

                        }
                        else
                        {
                            oJournalEntries.Lines.AccountCode = strDefGL;
                            oJournalEntries.Lines.Credit = dblAmount;
                            oJournalEntries.Lines.Debit = 0;

                            oJournalEntries.Lines.Add();
                            oJournalEntries.Lines.AccountCode = strVatGL;
                            oJournalEntries.Lines.Credit = 0;
                            oJournalEntries.Lines.Debit = dblAmount;

                            oJournalEntries.Lines.UserFields.Fields.Item("U_xWTCode").Value = strxOWTCode;
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xWTVendor").Value = strxCardCode;
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xTaxbleAmnt").Value = dblxOAmount;
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xSupplierName").Value = oRecordset.Fields.Item("CardName").Value.ToString();
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xAddress").Value = strAddrss;
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xTINnumber").Value = strTIN;
                            oJournalEntries.Lines.UserFields.Fields.Item("U_xCardType").Value = oRecordset.Fields.Item("CardType").Value.ToString();

                            if (dblxZAmount > 0)
                            {
                                oJournalEntries.Lines.Add();
                                oJournalEntries.Lines.AccountCode = strZerGL;
                                oJournalEntries.Lines.Debit = 0;
                                oJournalEntries.Lines.Credit = 0;

                                oJournalEntries.Lines.UserFields.Fields.Item("U_xWTCode").Value = strxZWTCode;
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xWTVendor").Value = strxCardCode;
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xTaxbleAmnt").Value = dblxZAmount;
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xSupplierName").Value = oRecordset.Fields.Item("CardName").Value.ToString();
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xAddress").Value = strAddrss;
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xTINnumber").Value = strTIN;
                                oJournalEntries.Lines.UserFields.Fields.Item("U_xCardType").Value = oRecordset.Fields.Item("CardType").Value.ToString();
                            }
                        }

                        if (oJournalEntries.Add() != 0)
                        {
                            GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                            GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                            return false;
                        }
                    }
                }
                else
                {
                    GlobalVariable.intErrNum = -699;
                    GlobalVariable.strErrMsg = "AR Invoice IEMOP GL Settings is missing.";

                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), string.Format("{0} - CardCode:{1}", GlobalVariable.strErrMsg, strxCardCode));

                return false;
            }
        }
        public static void importFromReversal()
        {

            SAPbobsCOM.Recordset oRecordset;

            DateTime dteDoc;

            string strQuery, strDocEntry, strDocnum, strCardCode;

            double dblTotVAT, dblVatableSales, dblTotZero;

            try
            {
                dteStart = DateTime.Now;

                strTransType = "Journal Entry - Deferred Tax ReClass";

                strQuery = string.Format("SELECT ORCT.\"DocEntry\", ORCT.\"DocNum\", ORCT.\"DocDate\", ORCT.\"U_VAT\", ORCT.\"U_ZeroRated\", " +
                                         "       ORCT.\"CardCode\", ORCT.\"U_VatableSales\" " +
                                         "FROM \"ORCT\"  " +
                                         "WHERE ORCT.\"Canceled\" = 'Y' AND ORCT.\"U_ReClsCanSts\" = 'N' AND \"U_TransType\" = 'IEMOP' ");
                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                {
                    while (!(oRecordset.EoF))
                    {
                        dteDoc = Convert.ToDateTime(oRecordset.Fields.Item("DocDate").Value.ToString());

                        strDocEntry = oRecordset.Fields.Item("DocEntry").Value.ToString();
                        strDocnum = oRecordset.Fields.Item("DocNum").Value.ToString();
                        strCardCode = oRecordset.Fields.Item("CardCode").Value.ToString();

                        dblTotZero = Convert.ToDouble(oRecordset.Fields.Item("U_ZeroRated").Value.ToString());
                        dblTotVAT = Convert.ToDouble(oRecordset.Fields.Item("U_VAT").Value.ToString());
                        dblVatableSales = Convert.ToDouble(oRecordset.Fields.Item("U_VatableSales").Value.ToString());

                        if (!(ImportUserDefinedJournalEntry.importFromObject("Cancel", dteDoc, strDocEntry, strDocnum, dblTotVAT, strCardCode, dblVatableSales, dblTotZero)))
                        {
                            strQuery = string.Format("UPDATE \"ORCT\" SET \"U_ReClsCanSts\" = 'E' WHERE \"DocEntry\" = '{0}' ", strDocEntry);
                            if (!(SystemFunction.executeQuery(strQuery)))
                            {

                                GlobalVariable.intErrNum = -699;
                                GlobalVariable.strErrMsg = string.Format("Error updating Cancelation Base Document for Deferred VAT ReClass.");

                                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            }
                        }
                        else
                        {
                            strQuery = string.Format("UPDATE \"ORCT\" SET \"U_ReClsCanSts\" = 'Y' WHERE \"DocEntry\" = '{0}' ", strDocEntry);
                            if (!(SystemFunction.executeQuery(strQuery)))
                            {

                                GlobalVariable.intErrNum = -699;
                                GlobalVariable.strErrMsg = string.Format("Error updating Cancelation Base Document for Deferred VAT ReClass.");

                                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            }
                        }

                        oRecordset.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

            }

            GC.Collect();

        }

        //CHANGE REQUEST 27/06/2024 - Creditable Withholding Tax – Auto JE from IEMOP Incoming Payment
        public static bool importAutoCWT(string strCardCode, string strNumAtCard, double dblTotWTax,double dblTotVatSales, double dblTotZroRated, DateTime dteDoc, DateTime dteTax, DateTime dteDue)
        {           

            SAPbobsCOM.JournalEntries oJournalEntries = null;
            SAPbobsCOM.Recordset oRecordset = null;

            string strCardName, strAddress, strCardType;
            string strwtaxCode, strAcctCode;
            string strTIN;
            string strQuery;       

            try
            {

                strQuery = string.Format("SELECT OCRD.\"CardCode\", CRD1.\"Street\", OWHT.\"Account\", OCRD.\"U_IemopWtax\", OCRD.\"CardName\", OCRD.\"U_TIN1\", OCRD.\"CardType\" " +
                    "FROM OCRD " +
                    "INNER JOIN CRD1 ON OCRD.\"CardCode\" = CRD1.\"CardCode\" " + 
                    "INNER JOIN OWHT ON OCRD.\"U_IemopWtax\" = OWHT.\"WTCode\" WHERE OCRD.\"CardCode\" = '{0}'", strCardCode);

                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (!oRecordset.EoF)
                {
                    strCardCode = oRecordset.Fields.Item("CardCode").Value.ToString();
                    strCardName = oRecordset.Fields.Item("CardName").Value.ToString();
                    strAddress = oRecordset.Fields.Item("Street").Value.ToString();
                    strTIN = oRecordset.Fields.Item("U_TIN1").Value.ToString();
                    strCardType = oRecordset.Fields.Item("CardType").Value.ToString();
                    strwtaxCode = oRecordset.Fields.Item("U_IemopWtax").Value.ToString();
                    strAcctCode = oRecordset.Fields.Item("Account").Value.ToString();
                }
                else
                {
                    return false;
                }

                if (dblTotVatSales != 0)
                {
                    oJournalEntries = (SAPbobsCOM.JournalEntries)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);
                    oJournalEntries.ReferenceDate = dteDoc;
                    oJournalEntries.DueDate = dteDoc;
                    oJournalEntries.TaxDate = dteDoc;
                    oJournalEntries.Memo = "Recognition of CWT";
                    oJournalEntries.Reference = strNumAtCard;
                    oJournalEntries.Reference2 = strCardName;

                    oJournalEntries.Lines.AccountCode = strAcctCode;
                    oJournalEntries.Lines.ShortName = strAcctCode;
                    oJournalEntries.Lines.Credit = 0;
                    oJournalEntries.Lines.Debit = dblTotWTax;

                    oJournalEntries.Lines.UserFields.Fields.Item("U_xWTCode").Value = strwtaxCode;
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xWTVendor").Value = strCardCode;
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xTaxbleAmnt").Value = dblTotVatSales;                    
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xSupplierName").Value = strCardName;
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xAddress").Value = strAddress;
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xTINnumber").Value = strTIN;
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xCardType").Value = strCardType;

                    oJournalEntries.Lines.Add();

                    oJournalEntries.Lines.ShortName = strCardCode;
                    oJournalEntries.Lines.Credit = dblTotWTax;
                    oJournalEntries.Lines.Debit = 0;
                }

                if (dblTotZroRated != 0)
                {
                    oJournalEntries = (SAPbobsCOM.JournalEntries)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);
                    oJournalEntries.ReferenceDate = dteDoc;
                    oJournalEntries.DueDate = dteDoc;
                    oJournalEntries.TaxDate = dteDoc;
                    oJournalEntries.Memo = "Recognition of CWT";
                    oJournalEntries.Reference = strNumAtCard;
                    oJournalEntries.Reference2 = strCardName;

                    oJournalEntries.Lines.AccountCode = strAcctCode;
                    oJournalEntries.Lines.ShortName = strAcctCode;
                    oJournalEntries.Lines.Credit = 0;
                    oJournalEntries.Lines.Debit = dblTotWTax;

                    oJournalEntries.Lines.UserFields.Fields.Item("U_xWTCode").Value = strwtaxCode;
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xWTVendor").Value = strCardCode;
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xTaxbleAmnt").Value = dblTotZroRated;                
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xSupplierName").Value = strCardName;
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xAddress").Value = strAddress;
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xTINnumber").Value = strTIN;
                    oJournalEntries.Lines.UserFields.Fields.Item("U_xCardType").Value = strCardType;

                    oJournalEntries.Lines.Add();

                    oJournalEntries.Lines.ShortName = strCardCode;
                    oJournalEntries.Lines.Credit = dblTotWTax;
                    oJournalEntries.Lines.Debit = 0;
                }

                if (oJournalEntries.Add() != 0)
                {
                    GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                    return false;
                }
                else
                {
                    GlobalVariable.strCWTJENum = GlobalVariable.oCompany.GetNewObjectKey();
                }

                return true;
            }
            catch (Exception ex)
            {
                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), string.Format("{0} - CardCode:{1}", GlobalVariable.strErrMsg, strCardCode));

                return false;
            }
        }
    }

}
