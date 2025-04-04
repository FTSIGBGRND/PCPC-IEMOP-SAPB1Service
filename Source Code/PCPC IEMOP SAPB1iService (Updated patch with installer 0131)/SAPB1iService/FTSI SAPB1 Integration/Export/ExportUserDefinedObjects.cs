using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;

namespace SAPB1iService
{
    class ExportUserDefinedObjects
    {
        private static DateTime dteStart;
        private static string strTransType;

        public static void _ExportUserDefinedObjects()
        {
            exportBIR2307();
        }
        public static void exportBIR2307()
        {
            string strQuery;
            string strCardCode, strDocNum, strSTLID, strNumAtCard, strBllDate, strUpDate, strBIR2307File;

            DateTime dteBllDate, dteUpDate;
            SAPbobsCOM.Recordset oRecordset;

            try
            {
                strTransType = "Export - Crystal Report BIR 2307";

                dteStart = DateTime.Now;

                strQuery = string.Format("SELECT OPCH.\"DocNum\", OPCH.\"NumAtCard\", OPCH.\"CardCode\", OPCH.\"U_STLID\", OPCH.\"U_BllDate\", OPCH.\"U_UpDate\" " +
                                         "FROM OPCH INNER JOIN \"@FTISSP\" ISSP ON OPCH.\"NumAtCard\" = ISSP.\"U_NumAtCard\" " +
                                         "          INNER JOIN PCH5 ON OPCH.\"DocEntry\" = PCH5.\"AbsEntry\" " +
                                         "WHERE ISSP.\"Code\" = '2' AND IFNULL(ISSP.\"U_NumAtCard\", '') != '' ");

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                {
                    while (!(oRecordset.EoF))
                    {

                        strDocNum = oRecordset.Fields.Item("DocNum").Value.ToString();
                        strCardCode = oRecordset.Fields.Item("CardCode").Value.ToString();
                        strSTLID = oRecordset.Fields.Item("U_STLID").Value.ToString();

                        strNumAtCard = oRecordset.Fields.Item("NumAtCard").Value.ToString();
                        strNumAtCard = Regex.Replace(strNumAtCard, "[^0-9.]", "");

                        dteBllDate = Convert.ToDateTime(oRecordset.Fields.Item("U_BllDate").Value.ToString());
                        strBllDate = dteBllDate.ToString("MM") + dteBllDate.ToString("yy");

                        dteUpDate = Convert.ToDateTime(oRecordset.Fields.Item("U_UpDate").Value.ToString());
                        strUpDate = dteUpDate.ToString("MMddyyyy");

                        strBIR2307File = string.Format(@"{0}BIR2307_{1}_{2}_{3}_{4}_{5}.pdf", GlobalVariable.strExpPath2, GlobalVariable.strCompany, strSTLID, strNumAtCard, strBllDate, strUpDate);

                        exportBIR2307(strDocNum, strBIR2307File);

                        oRecordset.MoveNext();
                    }
                }

                strQuery = string.Format("UPDATE \"@FTISSP\" SET \"U_NumAtCard\" = '' WHERE \"Code\" = '2' ");
                if (!(SystemFunction.executeQuery(strQuery)))
                {

                    GlobalVariable.intErrNum = -899;
                    GlobalVariable.strErrMsg = string.Format("Error updating Intgeration Setup for BIR 2307.");

                    SystemFunction.transHandler("Export", strTransType, "", "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);
                }
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Export", strTransType, "", "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

            }
        }
        private static bool exportBIR2307(string strDocNum, string strFileName)
        {
            SAPbobsCOM.Recordset oRecordset;

            string strQuery, strCompany, strBIRParam, strBIRName = "2307" + GlobalVariable.strSBOUserName;

            try
            {
                dteStart = DateTime.Now;

                strTransType = "Export - Crystal Report BIR 2307";

                strQuery = string.Format("SELECT Max(CAST(\"Code\" AS INT)) + 1 as \"Code\" FROM \"@FT_BIRPARAM\" ");

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                    strBIRParam = oRecordset.Fields.Item("Code").Value.ToString();
                else
                    strBIRParam = "999";

                strQuery = string.Format("SELECT \"Code\", \"Name\" FROM \"@FT_BIRMASTER\" ");

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                    strCompany = oRecordset.Fields.Item("Code").Value.ToString();
                else
                    strCompany = GlobalVariable.strCompany;

                strQuery = string.Format("DELETE FROM \"@FT_BIRPARAM\" WHERE \"U_UserCode\" = '{0}' AND \"U_RType\" = '2307' ", GlobalVariable.strSBOUserName);
                if (!(SystemFunction.executeQuery(strQuery)))
                {

                    GlobalVariable.intErrNum = -799;
                    GlobalVariable.strErrMsg = string.Format("Error updating BIR 2307 Parameters");

                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    return false;
                }
                else
                {
                    strQuery = string.Format("INSERT INTO \"@FT_BIRPARAM\" (\"Code\", \"Name\", \"U_UserCode\", \"U_RType\", \"U_Company\", \"U_DocType\", \"U_DocNo\") " +
                                             "VALUES ('{0}', '{1}', '{2}', '2307', '{3}', 'AP', '{4}') ", strBIRParam, strBIRName, GlobalVariable.strSBOUserName, strCompany, strDocNum);

                    if (!(SystemFunction.executeQuery(strQuery)))
                    {

                        GlobalVariable.intErrNum = -798;
                        GlobalVariable.strErrMsg = string.Format("Error updating BIR 2307 Parameters");

                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        return false;
                    }
                    else
                    {
                        if (File.Exists(strFileName))
                            File.Delete(strFileName);

                        if (!(CrystalReport.CRBIR2307(strFileName)))                        
                            return false;           
                    }
                }

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Export", strTransType, "", "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                return false;
            }
        }

    }
}
