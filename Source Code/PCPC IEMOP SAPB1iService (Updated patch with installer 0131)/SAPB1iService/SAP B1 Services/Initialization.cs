using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPB1iService;

namespace SAPB1iService
{
    class Initialization
    {
        public static bool onInit()
        {
            try
            {
                if (SystemInitialization.initFolders())
                {
                    GlobalVariable.strSQLSettings = GlobalVariable.strFilePath + "\\Connection Path\\SQL_ConnectSettings.ini";
                    GlobalVariable.strSMTPSettings = GlobalVariable.strFilePath + "\\Connection Path\\E-Mail_ConnectSettings.ini";

                    SystemFunction.filewrite();

                    if (!(SystemInitialization.initSQLConnection()))
                    {
                        SystemFunction.errorAppend("Error Connecting SQL Database.");
                        return false;
                    }
                }            
                else
                    return false;

                return true;
            }
            catch (Exception ex)
            {
                SystemFunction.errorAppend(ex.Message.ToString());
                return false;
            }
        }
        public static bool connectSBO(string strConSettings)
        {

            SystemFunction.reconnectSAP();

            if (SystemFunction.connectSAP(strConSettings))
            {
                if (!(SystemInitialization.initTables()))
                {
                    SystemFunction.errorAppend(string.Format("Error Creating User Define Tables using {0} Connection Settings.", strConSettings));
                    return false;
                }

                if (!(SystemInitialization.initFields()))
                {
                    SystemFunction.errorAppend(string.Format("Error Creating User Define Fields using {0} Connection Settings.", strConSettings));
                    return false;
                }

                if (!(SystemInitialization.initUDO()))
                {
                    SystemFunction.errorAppend(string.Format("Error Creating User Define Objects using {0} Connection Settings.", strConSettings));
                    return false;
                }

                if (!(SystemInitialization.initStoreProcedure()))
                {
                    SystemFunction.errorAppend(string.Format("Error Executing SQL Scripts using {0} Connection Settings.", strConSettings));
                    return false;
                }

                SystemInitialization.initFolders();
            }
            else
            {
                return false;
            }

            return true;
        }
    }
}
