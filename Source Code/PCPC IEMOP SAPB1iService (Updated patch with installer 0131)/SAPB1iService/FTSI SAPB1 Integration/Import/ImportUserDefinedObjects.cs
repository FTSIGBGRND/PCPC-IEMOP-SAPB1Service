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

namespace SAPB1iService
{
    class ImportUserDefinedObjects
    {
        private static DateTime dteStart;

        private static DataTable oDTPlan, oDataTable;

        private static string strMsgBod, strStatus;
        private static string strTransType = "User Defined Objects - IEMOP Business Partner";
        public static void _ImportUserDefinedObjects()
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

                        if (strFile.Contains("IEMOP Business Partner"))
                        {
                            dteStart = DateTime.Now;

                            if (importIEMOPBusinessPartner(strFile))
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
                SystemFunction.transHandler("Import", strTransType, "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }
        private static bool importIEMOPBusinessPartner(string strFile)
        {
            string strQuery, strCardCode, strCardName, strSTLId, strBLLId;

            DataTable oDTHeader, oDTLines;

            SAPbobsCOM.Recordset oRecordset;
             
            bool blWithErr = false, blExist = false;

            SAPbobsCOM.CompanyService oCmpSrvc = null;
            SAPbobsCOM.GeneralService oGenService = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.GeneralData oGenData = null;
            SAPbobsCOM.GeneralData oChild;
            SAPbobsCOM.GeneralDataCollection oGenDataCol;

            try
            {
                if (GlobalFunction.importXLSX(Path.GetFullPath(strFile), "YES", "Sheet1"))
                {

                    if (!GlobalVariable.oCompany.InTransaction)
                        GlobalVariable.oCompany.StartTransaction();

                    oDTHeader = GlobalVariable.oDTImpData.DefaultView.ToTable(true, "CardCode", "CardName");

                    for (int intRowH = 0; intRowH <= oDTHeader.Rows.Count - 1; intRowH++)
                    {
                        strCardCode = oDTHeader.Rows[intRowH]["CardCode"].ToString();

                        strQuery = string.Format("SELECT \"CardName\" FROM OCRD WHERE \"CardCode\" = '{0}' ", strCardCode);
                        
                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (oRecordset.RecordCount > 0)
                            strCardName = oRecordset.Fields.Item("CardName").Value.ToString();
                        else
                        {
                            blWithErr = true;

                            GlobalVariable.intErrNum = -999;
                            GlobalVariable.strErrMsg = string.Format("{0} Not Found in Business Partner Master Data", strCardCode);

                            SystemFunction.transHandler("Import", strTransType, "FTIEMOP", GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            continue;
                        }

                        if (blWithErr == false)
                        {

                            strQuery = string.Format("SELECT \"Code\" FROM \"@FTOIEMOP\" WHERE \"Code\" = '{0}' ", strCardCode);

                            oRecordset = null;
                            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            oRecordset.DoQuery(strQuery);

                            if (!(oRecordset.RecordCount > 0))
                            {
                                blExist = false;

                                oCmpSrvc = GlobalVariable.oCompany.GetCompanyService();
                                oGenService = (SAPbobsCOM.GeneralService)oCmpSrvc.GetGeneralService("FTOIEMOP");
                                oGenData = (SAPbobsCOM.GeneralData)oGenService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                                oGenData.SetProperty("Code", strCardCode);
                                oGenData.SetProperty("Name", strCardName);
                            }
                            else
                            {
                                blExist = true;

                                oCmpSrvc = GlobalVariable.oCompany.GetCompanyService();
                                oGenService = (SAPbobsCOM.GeneralService)oCmpSrvc.GetGeneralService("FTOIEMOP");

                                oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGenService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                                oGeneralParams.SetProperty("Code", strCardCode);
                                oGenData = oGenService.GetByParams(oGeneralParams);
                            }
                        }

                        oDTLines = GlobalVariable.oDTImpData.Select("CardCode = '" + strCardCode + "' ").CopyToDataTable().DefaultView.ToTable();

                        for (int intRowL = 0; intRowL <= oDTLines.Rows.Count - 1; intRowL++)
                        {    
                            if (blWithErr == false)
                            {
                                strSTLId = oDTLines.Rows[intRowL]["STLID"].ToString();
                                strBLLId = oDTLines.Rows[intRowL]["BLLID"].ToString();

                                strQuery = string.Format("SELECT \"U_BLLID\" FROM \"@FTIEMOP1\" WHERE \"Code\" = '{0}' AND \"U_STLID\" = '{1}' AND \"U_BLLID\" = '{2}' ", strCardCode, strSTLId, strBLLId);

                                oRecordset = null;
                                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                oRecordset.DoQuery(strQuery);

                                if (!(oRecordset.RecordCount > 0))
                                {
                                    oGenDataCol = oGenData.Child("FTIEMOP1");
                                    oChild = oGenDataCol.Add();

                                    oChild.SetProperty("U_STLID", strSTLId);
                                    oChild.SetProperty("U_BLLID", strBLLId);
                                }
                            }
                        }

                        if (blWithErr == false)
                        {
                            try
                            {
                                if (blExist == true)
                                    oGenService.Update(oGenData);
                                else
                                    oGenService.Add(oGenData);
                            }
                            catch (Exception ex)
                            {

                                GlobalVariable.intErrNum = -111;
                                GlobalVariable.strErrMsg = ex.Message.ToString();

                                SystemFunction.transHandler("Import", strTransType, "FTIEMOP", GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                GC.Collect();

                                return false;
                            }
                        }
                    }

                    if (blWithErr == false)
                        SystemFunction.transHandler("Import", strTransType, "FTIEMOP", GlobalVariable.strFileName, "", "", dteStart, "S", "0", "Successfully uploaded IEMOP Business Partner.");

                    if (GlobalVariable.oCompany.InTransaction)
                        GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                    return true;
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

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                return false;
            }
        }
    }
}
