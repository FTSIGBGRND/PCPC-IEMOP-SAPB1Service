using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPB1iService;
using System.IO;

namespace SAPB1iService
{
    class SystemInitialization
    {
        public static bool initTables()
        {

            /******************************* TOUCH ME NOT PLEASE *****************************************************/

            if (SystemFunction.createUDT("FTPISL", "FT IEMOP Integration Log", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                return false;

            if (SystemFunction.createUDT("FTISSP", "FT Integration SetUp", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                return false;

            /****************************** UNTIL HERE - THANK YOU ***************************************************/

            if (SystemFunction.createUDT("FTOIEMOP", "OIEMOP - Business Partner H", SAPbobsCOM.BoUTBTableType.bott_MasterData) == false)
                return false;

            if (SystemFunction.createUDT("FTIEMOP1", "OIEMOP - Business Partner L", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines) == false)
                return false;

            if (SystemFunction.createUDT("FTIEMOPGL", "IEMOP GL Account", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                return false;

            return true;
        }
        public static bool initFields()
        {

            /******************************* TOUCH ME NOT PLEASE *****************************************************/

            #region "FRAMEWORK UDF"

            /******************************* INTEGRATION SERVICE LOG ***********************************************/

            if (SystemFunction.isUDFexists("@FTPISL", "Process") == false)
                if (SystemFunction.createUDF("@FTPISL", "Process", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "TransType") == false)
                if (SystemFunction.createUDF("@FTPISL", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "ObjType") == false)
                if (SystemFunction.createUDF("@FTPISL", "ObjType", "Object Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "TransDate") == false)
                if (SystemFunction.createUDF("@FTPISL", "TransDate", "Transaction Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "FileName") == false)
                if (SystemFunction.createUDF("@FTPISL", "FileName", "FileName", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "TrgtDocKey") == false)
                if (SystemFunction.createUDF("@FTPISL", "TrgtDocKey", "Base Document Key", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "TrgtDocNum") == false)
                if (SystemFunction.createUDF("@FTPISL", "TrgtDocNum", "Base Document No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "StartTime") == false)
                if (SystemFunction.createUDF("@FTPISL", "StartTime", "StartTime", SAPbobsCOM.BoFldSubTypes.st_Time, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "EndTime") == false)
                if (SystemFunction.createUDF("@FTPISL", "EndTime", "EndTime", SAPbobsCOM.BoFldSubTypes.st_Time, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "Status") == false)
                if (SystemFunction.createUDF("@FTPISL", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "ErrorCode") == false)
                if (SystemFunction.createUDF("@FTPISL", "ErrorCode", "Error Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "Remarks") == false)
                if (SystemFunction.createUDF("@FTPISL", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "", "", "") == false)
                    return false;

            /******************************* INTEGRATION SETUP ***********************************************/

            if (SystemFunction.isUDFexists("@FTISSP", "ExportFile") == false)
                if (SystemFunction.createUDF("@FTISSP", "ExportFile", "Export File Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ExportPath") == false)
                if (SystemFunction.createUDF("@FTISSP", "ExportPath", "Export Path", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ImportFile") == false)
                if (SystemFunction.createUDF("@FTISSP", "ImportFile", "Import File Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ImportPath") == false)
                if (SystemFunction.createUDF("@FTISSP", "ImportPath", "Import Path", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "Delimiter") == false)
                if (SystemFunction.createUDF("@FTISSP", "Delimiter", "Delimiter", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ProcessTime") == false)
                if (SystemFunction.createUDF("@FTISSP", "ProcessTime", "Process Time", SAPbobsCOM.BoFldSubTypes.st_Time, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "AlwaysRun") == false)
                if (SystemFunction.createUDF("@FTISSP", "AlwaysRun", "Services Always Running?", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ProcSer") == false)
                if (SystemFunction.createUDF("@FTISSP", "ProcSer", "Process Service", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "RunRep") == false)
                if (SystemFunction.createUDF("@FTISSP", "RunRep", "Reprocess Error File?", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "RepDate") == false)
                if (SystemFunction.createUDF("@FTISSP", "RepDate", "Reprocess Error Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "NumAtCard") == false)
                if (SystemFunction.createUDF("@FTISSP", "NumAtCard", "Vendor Reference for EWT", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            /************************** MARKETING DOCUMENTS ****************************************************************/

            if (SystemFunction.isUDFexists("OINV", "isExtract") == false)
                if (SystemFunction.createUDF("OINV", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, Y -Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "FileName") == false)
                if (SystemFunction.createUDF("OINV", "FileName", "FileName", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "RefNum") == false)
                if (SystemFunction.createUDF("OINV", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "RefNum") == false)
                if (SystemFunction.createUDF("INV1", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "BaseLine") == false)
                if (SystemFunction.createUDF("INV1", "BaseLine", "Base Line", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "BaseRef") == false)
                if (SystemFunction.createUDF("INV1", "BaseRef", "Base Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "BaseType") == false)
                if (SystemFunction.createUDF("INV1", "BaseType", "Base Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "RefNum") == false)
                if (SystemFunction.createUDF("INV3", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "BaseLine") == false)
                if (SystemFunction.createUDF("INV3", "BaseLine", "Base Line", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "BaseRef") == false)
                if (SystemFunction.createUDF("INV3", "BaseRef", "Base Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "BaseType") == false)
                if (SystemFunction.createUDF("INV3", "BaseType", "Base Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV5", "RefNum") == false)
                if (SystemFunction.createUDF("INV5", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            /************************** ITEM MASTER DATA ***************************************************************/

            if (SystemFunction.isUDFexists("OITM", "isExtract") == false)
                if (SystemFunction.createUDF("OITM", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, Y -Yes", "") == false)
                    return false;

            /************************** BUSINESS PARTNER DATA **********************************************************/

            if (SystemFunction.isUDFexists("OCRD", "isExtract") == false)
                if (SystemFunction.createUDF("OCRD", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, Y -Yes", "") == false)
                    return false;

            /************************** ADMINISTRATION ****************************************************************/

            if (SystemFunction.isUDFexists("OUSR", "IntMsg") == false)
                if (SystemFunction.createUDF("OUSR", "IntMsg", "Integration Message", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OADM", "Company") == false)
                if (SystemFunction.createUDF("OADM", "Company", "Company", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            #endregion

            /****************************** UNTIL HERE - THANK YOU ***************************************************/

            if (SystemFunction.isUDFexists("@FTIEMOP1", "STLID") == false)
                if (SystemFunction.createUDF("@FTIEMOP1", "STLID", "STL ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOP1", "BLLID") == false)
                if (SystemFunction.createUDF("@FTIEMOP1", "BLLID", "BILLING ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "VatSlePur") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "VatSlePur", "Vatable Sales/Purchase", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "TaxVatSlePur") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "TaxVatSlePur", "Tax Vatable Sales/Purchase", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "ZerSlePur") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "ZerSlePur", "Zero Rated Sales/Purchase", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "TaxZerSlePur") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "TaxZerSlePur", "Tax Zero Rated Sales/Purchase", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "ZerEcoSlePur") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "ZerEcoSlePur", "Zero Rated Ecozone Sales/Purchase", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "TaxZerEcoSlePur") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "TaxZerEcoSlePur", "Tax Zero Rated Ecozone Sales/Purchase", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "VatOnSlePur") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "VatOnSlePur", "Vat On Sales/Purchase", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "TaxVatOnSlePur") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "TaxVatOnSlePur", "Tax Vat On Sales/Purchase", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "PymntGL") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "PymntGL", "Payment GL Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "VATReClassGL") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "VATReClassGL", "VAT ReClass GL Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "ReClsOWTCod") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "ReClsOWTCod", "ReClass Output WT Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "ZerReClassGL") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "ZerReClassGL", "Zero Rated ReClass GL Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "ReClsZWTCod") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "ReClsZWTCod", "ReClass Zero Rated WT Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTIEMOPGL", "SeriesName") == false)
                if (SystemFunction.createUDF("@FTIEMOPGL", "SeriesName", "Series Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("PCH1", "TransType") == false)
                if (SystemFunction.createUDF("PCH1", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, "", "VS - Vatable Sales, ZS - Zero Rated Sales, ES - Econzone Sales, VOS - Vat On Sales," +
                                                                                                                               "VP - Vatable Purchase, ZP - Zero Rate Purchase, EP - Econzone Purchase, VOP - Vat On Purchase", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("PCH1", "BudgetCode") == false)
                if (SystemFunction.createUDF("PCH1", "BudgetCode", "Budget Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OPCH", "TransType") == false)
                if (SystemFunction.createUDF("OPCH", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OPCH", "STLID") == false)
                if (SystemFunction.createUDF("OPCH", "STLID", "STL ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OPCH", "BllDate") == false)
                if (SystemFunction.createUDF("OPCH", "BllDate", "Billing Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OPCH", "UpDate") == false)
                if (SystemFunction.createUDF("OPCH", "UpDate", "Uploading Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORCT", "TransType") == false)
                if (SystemFunction.createUDF("ORCT", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORCT", "ReClsCanSts") == false)
                if (SystemFunction.createUDF("ORCT", "ReClsCanSts", "ReClass Cancel Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "Y - Yes, N - No, E - Error", "") == false)
                    return false;

            //for incoming payments period
            if (SystemFunction.isUDFexists("ORCT", "PymntStrtPeriod") == false)
                if (SystemFunction.createUDF("ORCT", "PymntStrtPeriod", "Payment Start Period", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORCT", "PymntEndPeriod") == false)
                if (SystemFunction.createUDF("ORCT", "PymntEndPeriod", "Payment End Period", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OJDT", "ReClsBseDocEnt") == false)
                if (SystemFunction.createUDF("OJDT", "ReClsBseDocEnt", "ReClass Base Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OJDT", "ReClsBseDocNum") == false)
                if (SystemFunction.createUDF("OJDT", "ReClsBseDocNum", "ReClass Base Document No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            return true;
        }
        public static bool initUDO()
        {

            if (SystemFunction.createUDO("FTOIEMOP", "IEMOP Business Partner Setup", SAPbobsCOM.BoUDOObjType.boud_MasterData, "FTOIEMOP", "FTIEMOP1", "Code", false, false, false, false) == false)
                return false;

            return true;
        }
        public static bool initFolders()
        {
            try
            {
                string strDate = DateTime.Today.ToString("MMddyyyy") + @"\";

                string strExp = @"Export\" + strDate;
                string strImp = @"Import\" + strDate;

                GlobalVariable.strErrLogPath = GlobalVariable.strFilePath + @"\Error Log";
                if (!Directory.Exists(GlobalVariable.strErrLogPath))
                    Directory.CreateDirectory(GlobalVariable.strErrLogPath);

                GlobalVariable.strSQLScriptPath = GlobalVariable.strFilePath + @"\SQL Scripts\";
                if (!Directory.Exists(GlobalVariable.strSQLScriptPath))
                    Directory.CreateDirectory(GlobalVariable.strSQLScriptPath);

                GlobalVariable.strSAPScriptPath = GlobalVariable.strFilePath + @"\SAP Scripts\";
                if (!Directory.Exists(GlobalVariable.strSAPScriptPath))
                    Directory.CreateDirectory(GlobalVariable.strSAPScriptPath);

                GlobalVariable.strExpSucPath = GlobalVariable.strFilePath + @"\Success Files\" + strExp;
                if (!Directory.Exists(GlobalVariable.strExpSucPath))
                    Directory.CreateDirectory(GlobalVariable.strExpSucPath);

                GlobalVariable.strExpErrPath = GlobalVariable.strFilePath + @"\Error Files\" + strExp;
                if (!Directory.Exists(GlobalVariable.strExpErrPath))
                    Directory.CreateDirectory(GlobalVariable.strExpErrPath);

                GlobalVariable.strImpSucPath = GlobalVariable.strFilePath + @"\Success Files\" + strImp;
                if (!Directory.Exists(GlobalVariable.strImpSucPath))
                    Directory.CreateDirectory(GlobalVariable.strImpSucPath);

                GlobalVariable.strImpErrPath = GlobalVariable.strFilePath + @"\Error Files\" + strImp;
                if (!Directory.Exists(GlobalVariable.strImpErrPath))
                    Directory.CreateDirectory(GlobalVariable.strImpErrPath);

                GlobalVariable.strImpPath = GlobalVariable.strFilePath + @"\Import Files\";
                if (!Directory.Exists(GlobalVariable.strImpPath))
                    Directory.CreateDirectory(GlobalVariable.strImpPath);

                GlobalVariable.strExpPath = GlobalVariable.strFilePath + @"\Export Files\";
                if (!Directory.Exists(GlobalVariable.strExpPath))
                    Directory.CreateDirectory(GlobalVariable.strExpPath);

                GlobalVariable.strConPath = GlobalVariable.strFilePath + @"\Connection Path\";
                if (!Directory.Exists(GlobalVariable.strConPath))
                    Directory.CreateDirectory(GlobalVariable.strConPath);

                GlobalVariable.strTempPath = GlobalVariable.strFilePath + @"\Temp Files\";
                if (!Directory.Exists(GlobalVariable.strTempPath))
                    Directory.CreateDirectory(GlobalVariable.strTempPath);

                GlobalVariable.strAttImpPath = GlobalVariable.strFilePath + @"\Attachment\" + strImp;
                if (!Directory.Exists(GlobalVariable.strAttImpPath))
                    Directory.CreateDirectory(GlobalVariable.strAttImpPath);

                GlobalVariable.strAttExpPath = GlobalVariable.strFilePath + @"\Attachment\" + strExp;
                if (!Directory.Exists(GlobalVariable.strAttExpPath))
                    Directory.CreateDirectory(GlobalVariable.strAttExpPath);

                GlobalVariable.strArcExpPath = GlobalVariable.strFilePath + @"\Archive Files\Export\";
                if (!Directory.Exists(GlobalVariable.strArcExpPath))
                    Directory.CreateDirectory(GlobalVariable.strArcExpPath);

                GlobalVariable.strArcImpPath = GlobalVariable.strFilePath + @"\Archive Files\Import\";
                if (!Directory.Exists(GlobalVariable.strArcImpPath))
                    Directory.CreateDirectory(GlobalVariable.strArcImpPath);

                GlobalVariable.strCRPath = GlobalVariable.strFilePath + @"\Crystal Report\";
                if (!Directory.Exists(GlobalVariable.strCRPath))
                    Directory.CreateDirectory(GlobalVariable.strCRPath);

                GlobalVariable.strExpPath2 = GlobalVariable.strFilePath + @"\Export Files\" + GlobalVariable.strCompany + @"\";
                if (!Directory.Exists(GlobalVariable.strExpPath2))
                    Directory.CreateDirectory(GlobalVariable.strExpPath2);

                return true;
            }
            catch(Exception ex)
            {
                SystemFunction.errorAppend(string.Format("Error initializing program directory. {0}", ex.Message.ToString()));
                return false;
            }
        }
        public static bool initStoreProcedure()
        {
            //if (!(SystemFunction.initStoredProcedures(GlobalVariable.strSAPScriptPath)))
            //    return false;

            return true;
        }
        public static bool initSQLConnection()
        {
            if (File.Exists(GlobalVariable.strSQLSettings))
            {
                if (SystemFunction.connectSQL(GlobalVariable.strSQLSettings))
                    return true;
                else
                    return false;
            }
            else
                return false;
        }

    }
}
