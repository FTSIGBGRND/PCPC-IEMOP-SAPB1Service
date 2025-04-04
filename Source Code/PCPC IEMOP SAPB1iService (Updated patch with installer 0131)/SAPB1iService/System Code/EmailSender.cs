using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.IO;

namespace SAPB1iService
{
    class EmailSender
    {
        private static DateTime dteStart;
        private static string strMsgSub = "E-Mail Sender";
        public static void _EmailSender(string strProcess, string strStatus, string strFileName, string strPostDocNum, string strErrMsg)
        {
            string strBody;

            dteStart = DateTime.Now;

            GlobalVariable.strEmailSubject = GlobalVariable.strEmailSubject + " " + strFileName;

            if (strStatus == "S")
               strBody =  string.Format("Good day Team, \r\n\r\n " +
                                        "{0} is successful in integrating the payroll entry in the Journal Voucher no. {1}. ", strFileName, strPostDocNum);
            else
               strBody = string.Format("Good day Team, \r\n\r\n " +
                                       "{0} failed in integrating the payroll entry.\r\n " +
                                       "Cause of Error: {1} ", strFileName, strErrMsg);

            sendSMTPEmail(strProcess, GlobalVariable.strEmailSubject, strBody, GlobalVariable.strAttExpPath);
            
        }
        private static void sendSMTPEmail(string strProcess, string strSubject, string strBody, string strAttPath)
        {
            string strMsgBod;

            try
            {
                SystemFunction.getSMTPCredentials(GlobalVariable.strSMTPSettings);

                if (GlobalVariable.strSMTPEnable == "Yes")
                {
                    MailMessage emailmsg = new MailMessage();
                    SmtpClient smtpServer = new SmtpClient(GlobalVariable.strSMTPHost, GlobalVariable.intEmailPort);

                    emailmsg.From = new MailAddress(GlobalVariable.strEmailUserName);
                    emailmsg.To.Add(GlobalVariable.strEmailTo);
                    emailmsg.CC.Add(GlobalVariable.strEmailCC);

                    emailmsg.Subject = GlobalVariable.strEmailSubject + strSubject;

                    emailmsg.Body = strBody;

                    smtpServer.EnableSsl = true;

                    foreach (var strFile in Directory.GetFiles(GlobalVariable.strImpPath, "*.*"))
                    {
                        System.Net.Mail.Attachment attachment;
                        attachment = new System.Net.Mail.Attachment(strFile);
                        emailmsg.Attachments.Add(attachment);
                    }

                    smtpServer.Credentials = new System.Net.NetworkCredential(GlobalVariable.strEmailUserName, GlobalVariable.strEmailPassword);
                    smtpServer.ServicePoint.MaxIdleTime = 2;
                    smtpServer.Send(emailmsg);

                    updateBaseDoc();
                }    
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                strMsgBod = string.Format("Error send e-mail. {0} ", GlobalVariable.strErrMsg);

                GlobalFunction.sendAlert("E", strMsgSub, strMsgBod, GlobalVariable.oObjectType, GlobalVariable.oCompany.GetNewObjectKey().ToString());
                SystemFunction.transHandler(strProcess, "E-mail Sender", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }
        private static void updateBaseDoc()
        {
            
        }
    }
}
