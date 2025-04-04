using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.Odbc;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;

namespace SAPB1iService
{
    class GlobalVariable
    {
        public static SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

        public static SqlConnection mySqlConnection;
        public static OdbcConnection myOdbcConnection;

        public static SqlConnection SqlCon = new SqlConnection();
        public static SqlConnection SapCon = new SqlConnection();

        #region "File Location"

        public static string strFilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);

        public static string strSQLScriptPath, strSAPScriptPath, strImpPath, strExpPath, strErrLogPath, strConPath,
                             strExpSucPath, strImpSucPath, strExpErrPath, strImpErrPath, strFileName, strTempPath,
                             strImpConfPath, strExpConfPath, strAttImpPath, strAttExpPath, strArcImpPath, strArcExpPath,
                             strCRPath, strExpPath2;

        #endregion

        #region "System Variable"

        public static int intErrNum, intRetVal, intObjType, intBObjType;

        public static bool blinstalledUDO;

        public static SAPbobsCOM.BoObjectTypes oObjectType;
        public static SAPbobsCOM.BoObjectTypes oBObjectType;

        public static string strSQLSettings;

        public static string strErrLog, strErrMsg, strServer, strDBType, strDBUserName, strDBPassword, strSBOCompany, strSBOUserName, strSBOPassword;
        public static string strDocType, strBDocType;

        public static string strEncryptKey = "Fasttrack SAP B1 Connection Settings Generator Encryption Program";

        public static string strTableHeader, strTableLine1, strTableLine3, strTableLine5;

        public static string strBTableHeader, strBTableLine1, strBTableLine3, strBTableLine5;

        public static char chrDlmtr;

        public static string strImpExt, strExpExt;

        public static string strSMTPSettings, strSMTPEnable, strSMTPHost, strEmailUserName, strEmailPassword, strEmailSubject, strEmailTo, strEmailCC;

        public static int intEmailPort;

        #endregion

        #region "Program Variable"

        public static string strCompany, strCWTJENum;
        public static string strDocEntry, strDocNum;

        public bool blEWTErr;

        #endregion

        #region "DataTable"

        public static DataTable oDTImpData = new DataTable("ImportData");

        #endregion




    }
}
