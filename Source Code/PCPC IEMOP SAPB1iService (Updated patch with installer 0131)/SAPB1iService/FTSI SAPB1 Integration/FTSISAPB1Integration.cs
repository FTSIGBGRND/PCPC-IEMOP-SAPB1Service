using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using System.IO;

namespace SAPB1iService
{
    class FTSISAPB1Integration
    {
        public static void _FTSISAPB1Integration()
        {
            SAPbobsCOM.Recordset oRecordset;

            string strAlRun, strProcSer, strQuery;
            int timeOfDay;

            try
            {
                timeOfDay = Convert.ToInt16(DateTime.Now.ToString("HH:mm").Replace(":", ""));

                strQuery = "SELECT TOP 1 ISSP.\"Code\", ISSP.\"U_ExportFile\", ISSP.\"U_ExportPath\", ISSP.\"U_ImportFile\", ISSP.\"U_ImportPath\", " +
                           "             ISSP.\"U_AlwaysRun\", ISSP.\"U_Delimiter\", ISSP.\"U_ProcSer\", OADM.\"U_Company\" " +
                           "FROM \"@FTISSP\" \"ISSP\", \"OADM\" " +
                           "WHERE ISSP.\"Code\" = '2' AND ISSP.\"U_ProcessTime\" <= " + timeOfDay + " ";

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                {
                    strAlRun = oRecordset.Fields.Item("U_AlwaysRun").Value.ToString();
                    strProcSer = oRecordset.Fields.Item("U_ProcSer").Value.ToString();

                    if (!(string.IsNullOrEmpty(oRecordset.Fields.Item("U_Delimiter").Value.ToString())))
                        GlobalVariable.chrDlmtr = Convert.ToChar(oRecordset.Fields.Item("U_Delimiter").Value.ToString());

                    GlobalVariable.strExpExt = oRecordset.Fields.Item("U_ExportFile").Value.ToString();
                    GlobalVariable.strExpConfPath = oRecordset.Fields.Item("U_ExportPath").Value.ToString();
                    GlobalVariable.strImpExt = oRecordset.Fields.Item("U_ImportFile").Value.ToString();
                    GlobalVariable.strImpConfPath = oRecordset.Fields.Item("U_ImportPath").Value.ToString();
                    GlobalVariable.strCompany = oRecordset.Fields.Item("U_Company").Value.ToString();

                    GlobalVariable.strExpPath = GlobalVariable.strExpPath + GlobalVariable.strCompany + @"\";

                    if (!Directory.Exists(GlobalVariable.strExpPath))
                        Directory.CreateDirectory(GlobalVariable.strExpPath);

                    if (strAlRun == "Y")
                    {
                        Import._Import();
                        Export._Export();
                    }
                    else
                    {
                        if (strProcSer == "Y")
                        {
                            Import._Import();
                            Export._Export();
                        }
                    }
                }
                else
                {
                    SystemFunction.errorAppend("Integration Setup is missing. Go to Tools > User Defined Windows > FTISSP - FT Integration Setup.");
                }
            }
            catch (Exception ex)
            {
                SystemFunction.errorAppend(string.Format("Error validating integration setup.", ex.Message.ToLower()));
            }
        }

    }

}
