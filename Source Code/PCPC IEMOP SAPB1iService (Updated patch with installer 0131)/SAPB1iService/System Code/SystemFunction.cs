using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using SAPbobsCOM;
using SAPB1iService;
using System.Windows;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data.Odbc;
using System.Data;
using System.Security.Cryptography;


namespace SAPB1iService
{
    class SystemFunction
    {
        private static TripleDESCryptoServiceProvider DES = new TripleDESCryptoServiceProvider();
        private static MD5CryptoServiceProvider MD5 = new MD5CryptoServiceProvider();

        public static bool getSMTPCredentials(string strPathConnect)
        {
            string[] strLines;

            try
            {
                strLines = File.ReadAllLines(strPathConnect);

                GlobalVariable.strSMTPEnable = strLines[0].ToString().Substring(strLines[0].IndexOf("=") + 1);
                GlobalVariable.strSMTPHost = strLines[1].ToString().Substring(strLines[1].IndexOf("=") + 1);
                GlobalVariable.intEmailPort = Convert.ToInt32(strLines[2].ToString().Substring(strLines[2].IndexOf("=") + 1));
                GlobalVariable.strEmailUserName = strLines[3].ToString().Substring(strLines[3].IndexOf("=") + 1);
                GlobalVariable.strEmailPassword = strLines[4].ToString().Substring(strLines[4].IndexOf("=") + 1);
                GlobalVariable.strEmailTo = strLines[5].ToString().Substring(strLines[5].IndexOf("=") + 1);
                GlobalVariable.strEmailCC = strLines[6].ToString().Substring(strLines[6].IndexOf("=") + 1);
                GlobalVariable.strEmailSubject = strLines[7].ToString().Substring(strLines[7].IndexOf("=") + 1);

            }        
            catch (Exception ex)
            {
                SystemFunction.errorAppend(ex.Message.ToString());
                return false;
            }
            return true;
        }
        public static bool connectSQL(string strPathConnect)
        {

            string strServer, strSQLDB, strDBUserName, strDBPassword;
            string[] strLines;

            try
            {

                strLines = File.ReadAllLines(strPathConnect);

                strServer = strLines[0].ToString().Substring(strLines[0].IndexOf("=") + 1);
                strSQLDB = strLines[1].ToString().Substring(strLines[1].IndexOf("=") + 1);
                strDBUserName = strLines[2].ToString().Substring(strLines[2].IndexOf("=") + 1);
                strDBPassword = strLines[3].ToString().Substring(strLines[3].IndexOf("=") + 1);

                if (!string.IsNullOrEmpty(strServer))
                {

                    GlobalVariable.SqlCon = new SqlConnection(string.Format("Data Source = {0}; Initial Catalog = {1}; User ID = {2}; Password = {3}", strServer, strSQLDB, strDBUserName, strDBPassword));

                    if (GlobalVariable.SqlCon.State == ConnectionState.Closed)
                        GlobalVariable.SqlCon.Open();

                    if (GlobalVariable.SqlCon.State == ConnectionState.Open)
                        GlobalVariable.SqlCon.Close();

                }

            }
            catch (Exception ex)
            {
                SystemFunction.errorAppend(ex.Message.ToString());
                return false;
            }

            return true;
        }
        public static bool connectSAP(string strPathConnect)
        {

            string[] strLines;

            try
            {

                strLines = File.ReadAllLines(strPathConnect);

                GlobalVariable.strServer = strLines[0].ToString().Substring(strLines[0].IndexOf("=") + 1);
                GlobalVariable.strDBType = strLines[1].ToString().Substring(strLines[1].IndexOf("=") + 1);
                GlobalVariable.strDBUserName = strLines[2].ToString().Substring(strLines[2].IndexOf("=") + 1);
                GlobalVariable.strDBPassword = strLines[3].ToString().Substring(strLines[3].IndexOf("=") + 1);
                GlobalVariable.strSBOCompany = strLines[4].ToString().Substring(strLines[4].IndexOf("=") + 1);
                GlobalVariable.strSBOUserName = strLines[5].ToString().Substring(strLines[5].IndexOf("=") + 1);
                GlobalVariable.strSBOPassword = strLines[6].ToString().Substring(strLines[6].IndexOf("=") + 1);

                GlobalVariable.strDBPassword = SystemFunction.Decrypt(GlobalVariable.strDBPassword, GlobalVariable.strEncryptKey);
                GlobalVariable.strSBOPassword = SystemFunction.Decrypt(GlobalVariable.strSBOPassword, GlobalVariable.strEncryptKey);

            }
            catch (Exception ex)
            {
                SystemFunction.errorAppend(ex.Message.ToString());
                return false;
            }

            SystemFunction.connectDIAPI(GlobalVariable.strServer, GlobalVariable.strDBType, GlobalVariable.strDBUserName, GlobalVariable.strDBPassword, 
                                        GlobalVariable.strSBOCompany, GlobalVariable.strSBOUserName, GlobalVariable.strSBOPassword);

            if (!(GlobalVariable.intRetVal == 0))
            {
                GlobalVariable.oCompany.GetLastError(out GlobalVariable.intErrNum, out GlobalVariable.strErrMsg);
                SystemFunction.errorAppend(DateTime.Now.ToString() + "         " + GlobalVariable.intErrNum.ToString() + " - " + GlobalVariable.strErrMsg);
                return false;
            }
            else
            {
                GlobalVariable.SapCon = new SqlConnection(string.Format("Data Source = {0}; Initial Catalog = {1}; User ID = {2}; Password = {3}", 
                                                                        GlobalVariable.strServer, GlobalVariable.strSBOCompany, GlobalVariable.strDBUserName, GlobalVariable.strSBOPassword));

                SystemFunction.errorAppend("Connected to " + GlobalVariable.strSBOCompany + " - Company Database.");
            }

            return true;
        }
        public static void connectDIAPI(string serverName, string serverType, string dbUserName, string dbPassword,
                                string companyDB, string sapUserName, string sapPassword)
        {
            try
            {
                switch (serverType)
                {
                    case "SQL Server 2000":
                        GlobalVariable.oCompany.DbServerType = BoDataServerTypes.dst_MSSQL;
                        break;
                    case "SQL Server 2005":
                        GlobalVariable.oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2005;
                        break;
                    case "SQL Server 2008":
                        GlobalVariable.oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2008;
                        break;
                    case "SQL Server 2012":
                        GlobalVariable.oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2012;
                        break;
                    case "SQL Server 2014":
                        GlobalVariable.oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2014;
                        break;
                    case "SQL Server 2016":
                        GlobalVariable.oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2016;
                        break;
                    case "SQL Server 2017":
                        GlobalVariable.oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2017;
                        break;
                    //case "SQL Server 2019":
                    //    GlobalVariable.oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2019;
                    //    break;
                    case "HANA DB":
                        GlobalVariable.oCompany.DbServerType = BoDataServerTypes.dst_HANADB;
                        break;
                }

                GlobalVariable.oCompany.Server = serverName;
                GlobalVariable.oCompany.CompanyDB = companyDB;
                GlobalVariable.oCompany.DbUserName = dbUserName;
                GlobalVariable.oCompany.DbPassword = dbPassword;
                GlobalVariable.oCompany.UserName = sapUserName;
                GlobalVariable.oCompany.Password = sapPassword;
                GlobalVariable.oCompany.UseTrusted = false;
                GlobalVariable.oCompany.language = BoSuppLangs.ln_English;

                GlobalVariable.intRetVal = GlobalVariable.oCompany.Connect();

            }
            catch (Exception ex)
            {
                errorAppend(ex.ToString());
            }
        }
        public static void reconnectSAP()
        {
            if (GlobalVariable.oCompany.Connected)
            {
                if (GlobalVariable.oCompany.InTransaction)
                    GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);

                GlobalVariable.oCompany.Disconnect();
            }

            SystemFunction.releaseObj(GlobalVariable.oCompany);

            System.Threading.Thread.Sleep(30000);

            GlobalVariable.oCompany = new SAPbobsCOM.Company();

            GC.Collect();

        }
        public static bool createUDT(string as_tablename, string as_tabledescription, SAPbobsCOM.BoUTBTableType aole_tabletype)
        {

            GlobalVariable.intRetVal = 0;
            SAPbobsCOM.UserTablesMD UserTablesMD;

            try
            {
                UserTablesMD = (SAPbobsCOM.UserTablesMD)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (UserTablesMD.GetByKey(as_tablename) == false)
                {        
                    UserTablesMD.TableName = as_tablename;
                    UserTablesMD.TableDescription = as_tabledescription;
                    UserTablesMD.TableType = aole_tabletype;
                    GlobalVariable.intRetVal = UserTablesMD.Add();
                    if (GlobalVariable.intRetVal != 0)
                    {
                        GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                        GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();                        

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(UserTablesMD);
                        UserTablesMD = null;
                        GC.Collect();
                        return false;
                    }
                    else
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(UserTablesMD);
                        UserTablesMD = null;
                        GC.Collect();
                        GlobalVariable.blinstalledUDO = true;
                        return true;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UserTablesMD);
                UserTablesMD = null;
                GC.Collect();
                return true;
            }
            catch (Exception e)
            {
                SystemFunction.errorAppend("Errmsg -" + e.Message + ".");
                UserTablesMD = null;
                GC.Collect();
                return false;
            }
        }
        public static bool createUDF(string as_tablename, string as_name, string as_description, SAPbobsCOM.BoFieldTypes aole_type, int ai_size, string as_default, string as_options, string as_reltable)
        {
            long al_type;
            al_type = 0;

            switch (aole_type)
            {
                case SAPbobsCOM.BoFieldTypes.db_Alpha:
                    al_type = 0;
                    break;
                case SAPbobsCOM.BoFieldTypes.db_Date:
                    al_type = 3;
                    break;
                case SAPbobsCOM.BoFieldTypes.db_Float:
                    al_type = 4;
                    break;
                case SAPbobsCOM.BoFieldTypes.db_Memo:
                    al_type = 1;
                    break;
                case SAPbobsCOM.BoFieldTypes.db_Numeric:
                    al_type = 2;
                    break;
                default:
                    al_type = 0;
                    break;
            }

            return createUDF(as_tablename, as_name, as_description, al_type, ai_size, as_default, as_options, as_reltable);
        }
        public static bool createUDF(string as_tablename, string as_name, string as_description, SAPbobsCOM.BoFldSubTypes aole_type, int ai_size, string as_default, string as_options, string as_reltable)
        {
            long al_type;
            al_type = 0;

            switch (aole_type)
            {
                case SAPbobsCOM.BoFldSubTypes.st_Address:
                    al_type = 63;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Image:
                    al_type = 73;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Link:
                    al_type = 66;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Measurement:
                    al_type = 77;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_None:
                    al_type = 0;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Percentage:
                    al_type = 37;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Phone:
                    al_type = 35;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Price:
                    al_type = 80;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Quantity:
                    al_type = 81;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Rate:
                    al_type = 82;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Sum:
                    al_type = 83;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Time:
                    al_type = 84;
                    break;
                default:
                    break;
            }
            return createUDF(as_tablename, as_name, as_description, al_type, ai_size, as_default, as_options, as_reltable);
        }
        public static bool createUDF(string as_tablename, string as_name, string as_description, long al_type, int ai_size, string as_default, string as_options, string as_reltable)
        {
            GlobalVariable.intRetVal = 0;
            SAPbobsCOM.UserFieldsMD UserFieldsMD;
            int li_index;
            string ls_data;
            try
            {
                UserFieldsMD = (SAPbobsCOM.UserFieldsMD)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                UserFieldsMD.Description = as_description;
                UserFieldsMD.Name = as_name;
                UserFieldsMD.TableName = as_tablename; 
                if (!string.IsNullOrEmpty(as_reltable))
                    UserFieldsMD.LinkedTable = as_reltable;

                switch (al_type)
                {
                    case 0:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
                        if (ai_size > 0)
                        {
                            UserFieldsMD.EditSize = ai_size;
                        }
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 1:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo;
                        break;
                    case 2:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Numeric;
                        if (ai_size > 0)
                        {
                            UserFieldsMD.EditSize = ai_size;
                        }
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 3:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 4:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 77:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Measurement;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 37:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 80:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 81:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 82:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Rate;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 83:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Sum;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 84:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Time;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 35:
                        //SAPbobsCOM.BoFldSubTypes.st_Phone;
                        break;
                    case 63:
                        //SAPbobsCOM.BoFldSubTypes.st_Address;
                        break;
                    case 73:
                        //SAPbobsCOM.BoFldSubTypes.st_Image;
                        break;
                    case 66:
                        //SAPbobsCOM.BoFldSubTypes.st_Link;
                        break;
                    default:
                        break;

                }
                li_index = 0;
                while (as_options != "")
                {
                    ls_data = getToken(ref as_options, ",");
                    if (ls_data != "")
                    {
                        li_index = li_index + 1;
                        if (li_index > 0)
                        {
                            UserFieldsMD.ValidValues.Add();
                            UserFieldsMD.ValidValues.SetCurrentLine(li_index);

                        }

                        UserFieldsMD.ValidValues.Value = getToken(ref ls_data, "-");
                        UserFieldsMD.ValidValues.Description = getToken(ref ls_data, "-");

                    }
                }
                GlobalVariable.intRetVal = UserFieldsMD.Add();
                if (GlobalVariable.intRetVal != 0)
                {
                    GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                    SystemFunction.errorAppend("Add Field Failed~nTable Name: " + as_tablename + "Field Name: " + as_name + "Field Description: " + as_description + "Error No : " + GlobalVariable.intErrNum.ToString() + "Error Desciption : " + GlobalVariable.strErrMsg);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(UserFieldsMD);
                    UserFieldsMD = null;
                    GC.Collect();
                    return false;
                }
                else
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(UserFieldsMD);
                    UserFieldsMD = null;
                    GC.Collect();
                    return true;
                }

            }
            catch (Exception e)
            {
                SystemFunction.errorAppend("Errmsg -" + e.Message + ".");
                UserFieldsMD = null;
                GC.Collect();
                return false;
            }

        }
        public static bool isUDFexists(string as_tablename, string as_fieldname)
        {
            SAPbobsCOM.Recordset oRecordSet;

            oRecordSet = null;
            oRecordSet = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oRecordSet.DoQuery("select \"AliasID\" from \"CUFD\" where \"TableID\" ='" + as_tablename + "' and \"AliasID\" ='" + as_fieldname + "'");
                if (oRecordSet.RecordCount == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                    oRecordSet = null;
                    GC.Collect();
                    return false;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();

                return true;
            }
            catch (Exception e)
            {
                SystemFunction.errorAppend("Errmsg -" + e.Message + ".");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;

                GC.Collect();
                return false;
            }
        }
        public static bool createUDO(string as_code, string as_name, SAPbobsCOM.BoUDOObjType al_objecttype, string as_tablename, string as_childtables, string as_findcolumns, Boolean ab_manageseries)
        {
            try
            {
                if (al_objecttype == SAPbobsCOM.BoUDOObjType.boud_Document)
                {
                    return createUDO(as_code, as_name, al_objecttype, as_tablename, as_childtables, as_findcolumns, ab_manageseries, true, true, false);
                }
                else
                {
                    return createUDO(as_code, as_name, al_objecttype, as_tablename, as_childtables, as_findcolumns, ab_manageseries, false, false, true);
                }

            }
            catch (Exception e)
            {
               SystemFunction.errorAppend("Errmsg -" + e.Message + ".");
                return false;
            }
        }
        public static bool createUDO(string as_code, string as_name, SAPbobsCOM.BoUDOObjType al_objecttype, string as_tablename, string as_childtables, string as_findcolumns, Boolean ab_manageseries, Boolean ab_cancel, Boolean ab_close, Boolean ab_delete)
        {
            try
            {
                SAPbobsCOM.UserObjectsMD oUserObjectMD;
                SAPbobsCOM.Recordset oRecordset;
                long ErrNumber;
                int li_index;
                string ErrMsg, ls_data;

                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                if (oUserObjectMD.GetByKey(as_code) == false)
                {
                    //CanCancel
                    if (ab_cancel)
                    {
                        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    else
                    {
                        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                    }
                    //CanClose
                    if (ab_close)
                    {
                        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    else
                    {
                        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                    }
                    //CanDelete
                    if (ab_delete)
                    {
                        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    else
                    {
                        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                    }

                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;

                    li_index = 0;
                    while (as_findcolumns != "")
                    {
                        ls_data = getToken(ref as_findcolumns, ",");
                        if (ls_data != "")
                        {
                            li_index = li_index + 1;
                            if (li_index > 0)
                            {
                                oUserObjectMD.FindColumns.Add();
                                oUserObjectMD.FindColumns.SetCurrentLine(li_index);
                            }
                            oUserObjectMD.FindColumns.ColumnAlias = ls_data;
                        }
                    }

                    oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.LogTableName = "";
                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;

                    li_index = 0;
                    while (as_childtables != "")
                    {
                        ls_data = getToken(ref as_childtables, ",");
                        if (ls_data != "")
                        {
                            li_index = li_index + 1;
                            if (li_index > 0)
                            {
                                oUserObjectMD.ChildTables.Add();
                                oUserObjectMD.ChildTables.SetCurrentLine(li_index);
                            }
                            oUserObjectMD.ChildTables.TableName = ls_data;
                        }
                    }


                    oUserObjectMD.ExtensionName = "";

                    if (ab_manageseries && al_objecttype == SAPbobsCOM.BoUDOObjType.boud_Document)
                    {
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    else
                    {
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                    }

                    oUserObjectMD.Code = as_code;
                    oUserObjectMD.Name = as_name;
                    oUserObjectMD.ObjectType = al_objecttype;
                    oUserObjectMD.TableName = as_tablename;

                    if (oUserObjectMD.Add() != 0)
                    {
                        GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                        GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                       SystemFunction.errorAppend("Add UDO Failed - " + "Table Name: " + oUserObjectMD.TableName + " UDO Name: " + oUserObjectMD.Code + " UDO Description: " + oUserObjectMD.Name + " Error No : " + GlobalVariable.intErrNum.ToString() +  " Error Desciption : " + GlobalVariable.strErrMsg);

                        oUserObjectMD = null;
                        GC.Collect();
                        return false;
                    }
                    else
                    {
                        GlobalVariable.blinstalledUDO = true;
                        if (ab_manageseries && al_objecttype == SAPbobsCOM.BoUDOObjType.boud_Document)
                        {
                            oRecordset = null;
                            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            oRecordset.DoQuery("Update NNM1 set Indicator='Default' where ObjectCode='" + as_code + "'");
                            oRecordset = null;
                            GC.Collect();
                        }
                    }
                }
                else
                {
                    oUserObjectMD = null;
                    GC.Collect();
                }

                GC.Collect();
                return true;

            }
            catch (Exception e)
            {
                SystemFunction.errorAppend("Errmsg - " + e.Message + ". Errorcode - " + GlobalVariable.oCompany.GetLastErrorCode().ToString() + "  Message - " + GlobalVariable.oCompany.GetLastErrorDescription());
                return false;
            }
        }
        public static bool executeQuery(string as_string)
        {
            try
            {
                SAPbobsCOM.Recordset executeQuery;
                executeQuery = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                executeQuery.DoQuery(as_string);
                executeQuery = null;
            }
            catch (Exception e)
            {
                SystemFunction.errorAppend("Errmsg -" + e.Message + ".");
                return false;
            }
            return true;
        }
        public static void filewrite()
        {
            GlobalVariable.strErrLog = GlobalVariable.strFilePath  + "\\Error Log\\" + DateTime.Now.ToString("MM-dd-yyyy") + " - ErrorLog.txt";

            if (!File.Exists(GlobalVariable.strErrLog))
            {
                FileInfo errorlog = new FileInfo(GlobalVariable.strErrLog);
                StreamWriter streamwriter = errorlog.CreateText();
                streamwriter.WriteLine("                                                      SAP Business One Integration Error Log");
                streamwriter.Close();
            }
        }
        public static void errorAppend(string text)
        {
            FileInfo errorlog = new FileInfo(GlobalVariable.strErrLog);
            StreamWriter streamwriter = errorlog.AppendText();
            streamwriter.WriteLine(DateTime.Now.TimeOfDay.ToString() + "          " + text);
            streamwriter.Close();
        }
        public static string getToken(ref string as_source, string as_separator)
        {
            string ls_ret;
            int li_pos;

            if (as_source == null || as_separator == "")
            {
                string ls_null;
                ls_null = null;
                return ls_null;
            }

            li_pos = as_source.IndexOf(as_separator);
            if (li_pos == -1)
            {
                ls_ret = as_source;
                as_source = "";
            }
            else
            {
                ls_ret = as_source.Substring(0, li_pos);
                as_source = as_source.Substring(li_pos + 1);
            }

            return ls_ret;


        }
        public static long getNextKey(string tableName)
        {
            long lngNxtKey;
            string strQuery;

            SAPbobsCOM.Recordset oRecordset;

            if (GlobalVariable.strDBType == "HANA DB")
                strQuery = "SELECT IFNULL(Count(\"Code\"),0) + 1 AS \"LineId\" FROM \"" + tableName + "\" ";
            else
                strQuery = "SELECT ISNULL(Count(Code),0) + 1 AS LineId FROM  \"" + tableName +  "\" ";

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(strQuery);

            lngNxtKey = System.Convert.ToInt64(oRecordset.Fields.Item("LineId").Value.ToString());

            return lngNxtKey;
        }
        public static void releaseObj(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                SystemFunction.errorAppend("Error releasing Object." + ex.Message.ToString());
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
        public static void transHandler(string strProcess, string TransType, string ObjType, string FileName, 
                                       string TrgtDocKey, string TrgtDocNum, DateTime tmStart, string Status, string ErrCode, string ErrMsg)
        {
            string strQuery;
            long lngNxtKey, lngNxtKey1;

            try
            {
                if (Status == "E")
                {
                    if (GlobalVariable.oCompany.InTransaction)
                        GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }

                lngNxtKey = SystemFunction.getNextKey("@FTPISL");


                if (GlobalVariable.strDBType != "HANA DB")
                    strQuery = "INSERT INTO \"@FTPISL\" VALUES ('" + lngNxtKey.ToString() + "', '" + lngNxtKey.ToString() + "', '" + strProcess + "', '" + TransType + "', '" + ObjType + "', '" + DateTime.Today.ToString("MM/dd/yyyy") + "', " +
                                                               "'" + FileName + "', '" + TrgtDocKey + "', '" + TrgtDocNum + "', '" + tmStart.ToString("HHmm") + "', '" + DateTime.Now.ToString("HHmm") + "', " +
                                                               "'" + Status + "', '" + ErrCode.ToString() + "', '" + ErrMsg.Replace("'", "''") + "' )";
                else
                    strQuery = "INSERT INTO \"@FTPISL\" VALUES ('" + lngNxtKey.ToString() + "', '" + lngNxtKey.ToString() + "', '" + strProcess + "', '" + TransType + "', '" + ObjType + "',   to_date('" + DateTime.Today.ToShortDateString() + "','MM/DD/YYYY'), " +
                                                               "'" + FileName + "', '" + TrgtDocKey + "', '" + TrgtDocNum + "', '" + tmStart.ToString("HHmm") + "', '" + DateTime.Now.ToString("HHmm") + "', " +
                                                               "'" + Status + "', '" + ErrCode.ToString() + "', '" + ErrMsg.Replace("'", "''") + "' )";

                if (!(SystemFunction.executeQuery(strQuery)))
                    errorAppend("Error while execute transaction history log.   -  " + strQuery);
            }
            catch (Exception ex)
            {
                errorAppend(ex.Message.ToString());
            }

            GC.Collect();

        }
        public static bool initStoredProcedures(string strPath)
        {

            string filepath = "";
            string ls_sql = "";
            string ls_sql2 = "";
            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Recordset oRecordset1;

            try
            {
                if (Directory.Exists(strPath))
                {
                    DirectoryInfo DirInfo = new DirectoryInfo(strPath);
                    FileInfo[] Files = DirInfo.GetFiles("*.sql");

                    if (GlobalVariable.oCompany.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        SystemFunction.createConnSQL(GlobalVariable.oCompany.CompanyDB);

                    else
                    {
                        SystemFunction.createConnODBC(GlobalVariable.oCompany.CompanyDB);
                    }

                    foreach (FileInfo file in Files)
                    {

                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordset1 = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        if (GlobalVariable.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            ls_sql = "select 'FUNCTION' \"TYP\" , \"FUNCTION_NAME\" \"NAME\" from SYS.FUNCTIONS where SCHEMA_NAME = CURRENT_SCHEMA and \"FUNCTION_NAME\" ='" + file.Name.Replace(".sql", "") + "' " +
                                     "union all " +
                                     "select 'VIEW' \"TYP\", \"VIEW_NAME\" \"NAME\" from SYS.VIEWS where SCHEMA_NAME = CURRENT_SCHEMA and \"VIEW_NAME\" ='" + file.Name.Replace(".sql", "") + "' " +
                                     "union all " +
                                     "select 'TRIGGER' \"TYP\", \"TRIGGER_NAME\" \"NAME\" from SYS.TRIGGERS where SCHEMA_NAME = CURRENT_SCHEMA and \"TRIGGER_NAME\" ='" + file.Name.Replace(".sql", "") + "' " +
                                     "union all " +
                                     "select 'PROCEDURE' \"TYP\", \"PROCEDURE_NAME\" \"NAME\" from SYS.PROCEDURES where SCHEMA_NAME = CURRENT_SCHEMA and \"PROCEDURE_NAME\" ='" + file.Name.Replace(".sql", "") + "';";

                            oRecordset.DoQuery(ls_sql);
                            if (oRecordset.RecordCount > 0)
                            {
                                ls_sql2 = "DROP " + oRecordset.Fields.Item("TYP").Value.ToString() + " \"" + oRecordset.Fields.Item("NAME").Value.ToString() + "\"; ";
                                oRecordset1.DoQuery(ls_sql2);
                            }
                        }

                        filepath = DirInfo.FullName + "\\" + file.Name;
                        if (!SystemFunction.execstoredproc(filepath))
                        {
                            return false;
                        }
                        else
                        {
                           //file.Delete();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                SystemFunction.errorAppend("Errmsg -" + e.Message.ToString() + ".");
                return false;
            }

            return true;
        }
        public static SqlConnection createConnSQL(string database)
        {
            // Here you define your server. Values can not be NULL.        //Database Server Name.
            //string myDSN = "SQLSERVER";

            //Local Server Name.
            string mySN = GlobalVariable.oCompany.Server;
            string myUserId = GlobalVariable.oCompany.DbUserName.ToString();
            string myPassword = GlobalVariable.strDBPassword;

            //string myPassword = "B1Admin";//DI.oCompany.DbPassword;
            //Define the type of security, 'TRUE' or 'FALSE'.
            //string mySecType = "TRUE";
            //string mySqlConnectionString = null;

            string mySqlConnectionString = ("Server = " + mySN + "; Database = " + database + "; User Id = " + myUserId + "; Password = " + myPassword + "; ");
            
            //Here you have your connection string you can edit it here.
            // Server = myServerAddress; Database = myDataBase; User ID = myUsername; Password = myPassword; Trusted_Connection = False;
            //string mySqlConnectionString = ("Data Source=" + mySN + ";Initial Catalog=" + database + ";Integrated Security=SSPI;");

            //If you wish to use SQL security, well just make your own connection string...
            // I make sure I have declare what DI.mySqlConnection stand for.

            if (GlobalVariable.mySqlConnection == null)
                GlobalVariable.mySqlConnection = new SqlConnection();

            // Since i will be reusing the connection I will try this it the connection dose not exist.
            if (GlobalVariable.mySqlConnection.ConnectionString == string.Empty || GlobalVariable.mySqlConnection.ConnectionString == null)
            {
                // I use a try catch stament cuz I use 2 set of arguments to connect to the database
                try
                {
                    //First I try with a pool of 5-40 and a connection time out of 4 seconds. then I open the connection.
                    GlobalVariable.mySqlConnection.ConnectionString = "Min Pool Size = 5; Max Pool Size = 40; Connect Timeout = 4; " + mySqlConnectionString + ";";
                    GlobalVariable.mySqlConnection.Open();
                }

                catch (Exception e)
                {
                    errorAppend(e.Message.ToString());
                    //If it did not work i try not using the pool and I give it a 45 seconds timeout.
                    try
                    {
                        if (GlobalVariable.mySqlConnection.State != ConnectionState.Closed)
                            GlobalVariable.mySqlConnection.Close();

                        GlobalVariable.mySqlConnection.ConnectionString = "Pooling = false; Connect Timeout = 45;" + mySqlConnectionString + ";";
                        GlobalVariable.mySqlConnection.Open();
                    }
                    catch (Exception ex)
                    {
                        errorAppend(ex.Message.ToString());
                    }
                }
                return GlobalVariable.mySqlConnection;
            }
            //Here if the connection exsist and is open i try this.
            if (GlobalVariable.mySqlConnection.State != ConnectionState.Open)
            {
                try
                {
                    GlobalVariable.mySqlConnection.ConnectionString = "Min Pool Size = 5; Max Pool Size = 40; Connect Timeout = 4;" + mySqlConnectionString + ";";
                    GlobalVariable.mySqlConnection.Open();
                }
                catch (Exception)
                {
                    if (GlobalVariable.mySqlConnection.State != ConnectionState.Closed)
                        GlobalVariable.mySqlConnection.Close();

                    GlobalVariable.mySqlConnection.ConnectionString = "Pooling = false; Connect Timeout = 45;" + mySqlConnectionString + ";";
                    GlobalVariable.mySqlConnection.Open();
                }
            }

            return GlobalVariable.mySqlConnection;
        }
        public static OdbcConnection createConnODBC(string database)
        {
            string mySN = GlobalVariable.oCompany.Server;
            string myUserId = GlobalVariable.oCompany.DbUserName;
            string myPassword = GlobalVariable.oCompany.DbPassword;

            string myOdbcConnectionString = "";

            if (IntPtr.Size == 8)
                myOdbcConnectionString = "DRIVER = {HDBODBC}; UID = " + myUserId + "; PWD = " + myPassword + "; SERVERNODE = " + mySN + "; CS = " + database + "";
            else
                myOdbcConnectionString = "DRIVER = {HDBODBC32}; UID = " + myUserId + "; PWD = " + myPassword + "; SERVERNODE = " + mySN + "; CS = " + database + "";


            if (GlobalVariable.myOdbcConnection == null)
                GlobalVariable.myOdbcConnection = new OdbcConnection(); 

            // Since i will be reusing the connection I will try this it the connection dose not exist.
            if (GlobalVariable.myOdbcConnection.ConnectionString == string.Empty || GlobalVariable.myOdbcConnection.ConnectionString == null)
            {
                try
                {
                    GlobalVariable.myOdbcConnection.ConnectionString = myOdbcConnectionString;
                    GlobalVariable.myOdbcConnection.Open();
                }

                catch (Exception e)
                {

                    SystemFunction.errorAppend(e.Message.ToString());

                    try
                    {

                        if (GlobalVariable.myOdbcConnection.State != ConnectionState.Closed)
                            GlobalVariable.myOdbcConnection.Close();

                        GlobalVariable.myOdbcConnection.ConnectionString = myOdbcConnectionString;
                        GlobalVariable.myOdbcConnection.Open();

                    }
                    catch (Exception ex)
                    {
                        SystemFunction.errorAppend(ex.Message.ToString());
                    }
                }

                return GlobalVariable.myOdbcConnection;
            }
            //Here if the connection exsist and is open i try this.
            if (GlobalVariable.myOdbcConnection.State != ConnectionState.Open)
            {
                try
                {
                    GlobalVariable.myOdbcConnection.ConnectionString = myOdbcConnectionString;
                    GlobalVariable.myOdbcConnection.Open();
                }
                catch (Exception)
                {
                    if (GlobalVariable.myOdbcConnection.State != ConnectionState.Closed)
                        GlobalVariable.myOdbcConnection.Close();

                    GlobalVariable.myOdbcConnection.ConnectionString = myOdbcConnectionString;
                    GlobalVariable.myOdbcConnection.Open();
                }
            }
            return GlobalVariable.myOdbcConnection;
        }
        public static bool execstoredproc(string path)
        {
            using (StreamReader sr = new StreamReader(path))
            {
                string script = sr.ReadToEnd();
                //Server serverM = new Server(new ServerConnection(mySqlConnection));
                //serverM.ConnectionContext.ExecuteNonQuery(script);
                createStorepProcedure(script);
            }

            return true;
        }
        public static bool createStorepProcedure(string as_sql)
        {
            string ls_sql;
            try
            {

                foreach (var batch in as_sql.Split(new string[] { "\nGO", "\ngo" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    try
                    {
                        if (GlobalVariable.oCompany.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            new SqlCommand(batch, GlobalVariable.mySqlConnection).ExecuteNonQuery();
                        }
                        else
                        {
                            new OdbcCommand(batch, GlobalVariable.myOdbcConnection).ExecuteNonQuery();
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        throw;
                    }
                }
                //GlobalFunction.fileappend("Created Stored Procedure - " + as_sql);
                return true;
            }
            catch (Exception e)
            {

                string errmsg;
                int errcode;

                errcode = GlobalVariable.oCompany.GetLastErrorCode();
                errmsg = GlobalVariable.oCompany.GetLastErrorDescription();
                SystemFunction.errorAppend("Errmsg -" + e.Message.ToString() + ". Errorcode - " + errcode.ToString() + "  Msg - " + errmsg);
                return false;

            }
        }
        public static byte[] MD5Hash(string value)
        {
            return MD5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(value));
        }
        public static string Encrypt(string stringToEncrypt, string key)
        {
            DES.Key = MD5Hash(key);
            DES.Mode = CipherMode.ECB;
            byte[] Buffer = ASCIIEncoding.ASCII.GetBytes(stringToEncrypt);

            return Convert.ToBase64String(DES.CreateEncryptor().TransformFinalBlock(Buffer, 0, Buffer.Length));
        }
        public static string Decrypt(string encryptedString, string key)
        { 
            try
            {
                DES.Key = MD5Hash(key);
                DES.Mode = CipherMode.ECB;
                byte[] Buffer = Convert.FromBase64String(encryptedString);
                return ASCIIEncoding.ASCII.GetString(DES.CreateDecryptor().TransformFinalBlock(Buffer, 0, Buffer.Length));
            }
            catch (Exception ex)
            {
                return "";
            }
        }
    }
}
