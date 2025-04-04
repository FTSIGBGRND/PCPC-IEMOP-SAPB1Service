using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using System.Net;
using Ionic.Zip;
using SAPbobsCOM;

namespace SAPB1iService
{
    class ProcServices
    {
        public static void _SAPB1Services()
        {          
            ConnectSAPB1();
        }
        private static void ConnectSAPB1()
        {
            string strFileName;

            foreach (var strFile in Directory.GetFiles(GlobalVariable.strConPath, "*SBO_*.ini"))
            {
                strFileName = Path.GetFileName(strFile);

                if (Initialization.connectSBO(strFile) == true)
                    FTSISAPB1Integration._FTSISAPB1Integration();
                else
                {
                    if (GlobalVariable.oCompany.Connected)
                        if (GlobalVariable.oCompany.InTransaction)
                            GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);

                    SAPB1Service myService = new SAPB1Service();
                    myService.Stop();

                    Environment.Exit(0);
                }
            }
        }
    }
}
