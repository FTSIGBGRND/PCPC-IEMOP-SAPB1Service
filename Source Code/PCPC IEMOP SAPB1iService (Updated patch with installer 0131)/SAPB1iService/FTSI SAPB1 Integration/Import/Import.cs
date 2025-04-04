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
using SAPB1iService;

namespace SAPB1iService
{
    class Import
    {
        public static void _Import()
        {

            //importFTPFiles();
            importSharedFiles();

            ImportUserDefinedObjects._ImportUserDefinedObjects();

            ImportDocuments._ImportDocuments();

            ImportPayments._ImportPayments();

            ImportJournalEntry._ImportJournalEntry();

            RemoveImport();

        }
        private static void importFTPFiles()
        {
            string[] strImpExt;

            if (!string.IsNullOrEmpty(GlobalVariable.strImpExt))
            {
                strImpExt = GlobalVariable.strImpExt.Split(Convert.ToChar("|"));

                for (int intStr = 0; intStr < strImpExt.Length; intStr++)
                {
                    TransferFile.importSFTPFiles(strImpExt[intStr]);
                }
            }
        }

        public static void importSharedFiles()
        { 
            string strFileNewPath;

            try 
            { 
                foreach (var strFile in Directory.GetFiles(GlobalVariable.strImpConfPath, "*.xlsx"))
                {
                    strFileNewPath = GlobalVariable.strImpPath + @"\" + Path.GetFileName(strFile);
                    File.Copy(strFile, strFileNewPath);
                }
            }
            catch (Exception ex)
            {
                SystemFunction.errorAppend(ex.Message.ToString());
            }
        }

        public static void RemoveImport()
        {
            string[] files = Directory.GetFiles(GlobalVariable.strImpConfPath);

            foreach (string file in files)
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
        }
    }
}
