using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAPB1iService
{
    class Export
    {
            
        private static DateTime dteStart;
        private static string strTransType;
        public static void _Export()
        {
            ExportDocuments._ExportDocuments();

            ExportSharedFiles();
        }

        public static void ExportSharedFiles()
        {
            string strFileNewPath, strNewFileName, strXFileNewPath;

            strTransType = "Export - Transfer File";


            try
            {
                strFileNewPath = GlobalVariable.strExpConfPath + @"\" + GlobalVariable.strCompany;

                if (!Directory.Exists(strFileNewPath))
                    Directory.CreateDirectory(strFileNewPath);

                string[] ExcelFiles = Directory.GetFiles(GlobalVariable.strExpPath2, "*.xlsx");
                string[] PdfFiles = Directory.GetFiles(GlobalVariable.strExpPath2, "*.pdf");

                foreach (var strFile in ExcelFiles.Concat(PdfFiles))
                {                    

                    dteStart = DateTime.Now;

                    //strNewFileName = strFileNewPath + @"\" + Path.GetFileName(strFile);
                    string strfileName = Path.GetFileName(strFile);
                    string fileExtension = Path.GetExtension(strFile);

                    if (strfileName.Contains("BIR2307") && fileExtension == ".pdf")
                    {
                        string BIRPath = strFileNewPath + @"\BIR2307\";
                        if (!Directory.Exists(BIRPath))
                            Directory.CreateDirectory(BIRPath);

                        if (File.Exists(BIRPath + strfileName))
                        {
                            strfileName = strfileName.Substring(0, strfileName.Length - 4) +
                                          DateTime.Now.ToString("_MMddyyyy_HHmmss") +
                                          fileExtension;
                        }
                        strNewFileName = BIRPath + strfileName;
                    }
                    else if (strfileName.Contains("OR") && fileExtension == ".pdf")
                    {
                        string ORPath = strFileNewPath + @"\Official Receipts\";
                        if (!Directory.Exists(ORPath))
                            Directory.CreateDirectory(ORPath);

                        if (File.Exists(ORPath + strfileName))
                        {
                            strfileName = strfileName.Substring(0, strfileName.Length - 4) +
                                          DateTime.Now.ToString("_MMddyyyy_HHmmss") +
                                          fileExtension;
                        }

                        strNewFileName = ORPath + strfileName;
                    }
                    else
                    {
                        if (File.Exists(strFileNewPath + @"\" + strfileName))
                        {
                            strfileName = strfileName.Substring(0, strfileName.Length - 4) +
                                          DateTime.Now.ToString("_MMddyyyy_HHmmss") +
                                          fileExtension;
                        }
                        strNewFileName = strFileNewPath + @"\" + strfileName;
                    }

                    File.Move(strFile, strNewFileName);
                   
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
