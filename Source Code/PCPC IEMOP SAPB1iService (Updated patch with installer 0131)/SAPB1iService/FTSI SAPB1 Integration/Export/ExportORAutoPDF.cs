using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAPB1iService
{
    class ExportORAutoPDF
    {
        private static string strFileName;
        private static DateTime dteStart;
        private static string strTransType;

        public static void _ExportORAutoPDF(string strNumAtCard)
        {
            string strQuery, strDocKey, strSetRef, strIssuer, strRecepient, strBllDate, strUpDate, strORFileName;
            DateTime dteBllDate, dteUpDate;

            SAPbobsCOM.Recordset oRecordset;
            try
            {
                dteStart = DateTime.Now;
                strTransType = "CrystalReport Export -  Official Receipt";

                strQuery = string.Format("SELECT ORCT.\"DocEntry\", ORCT.\"Comments\", ORCT.\"CreateDate\", ORCT.\"U_PymntStrtPeriod\", \"@FTIEMOP1\".\"U_STLID\", \"@FTIEMOP1\".\"U_BLLID\" " +
                    "FROM ORCT " +
                    "INNER JOIN \"@FTIEMOP1\" " +
                    "ON ORCT.\"CardCode\" = \"@FTIEMOP1\".\"Code\" " +
                    "WHERE ORCT.\"Comments\" = '{0}'", strNumAtCard);

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0 ) 
                {
                    while (!oRecordset.EoF)
                    {
                        strDocKey = oRecordset.Fields.Item("DocEntry").Value.ToString();
                        strIssuer = oRecordset.Fields.Item("U_STLID").Value.ToString();
                        strRecepient = oRecordset.Fields.Item("U_BLLID").Value.ToString();
                        strSetRef = oRecordset.Fields.Item("Comments").Value.ToString();

                        dteBllDate = Convert.ToDateTime(oRecordset.Fields.Item("U_PymntStrtPeriod").Value.ToString());
                        strBllDate = dteBllDate.ToString("MM") + dteBllDate.ToString("yy");

                        dteUpDate = Convert.ToDateTime(oRecordset.Fields.Item("CreateDate").Value.ToString());
                        strUpDate = dteUpDate.ToString("MMddyyyy");

                        strORFileName = string.Format(@"{0}OR_{1}_{2}_{3}_{4}_{5}.pdf", GlobalVariable.strExpPath, GlobalVariable.strCompany, strIssuer, strSetRef, strBllDate, strUpDate);

                        CrystalReport.ExportORToPDF(strDocKey, strORFileName);
                        oRecordset.MoveNext();
                    }
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

    }
}
