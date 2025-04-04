using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using System.Net;
using Ionic.Zip;
using Renci.SshNet;



namespace SAPB1iService
{
    class TransferFile
    {
        private static string strSFTPHost, strSFTPUserName, strSFTPPassword, strSFTPExpPath, strSFTPImpPath, strFTPErrRetPath, strFTPSucRetPath;

        private static string strMsgSub = "Transfer File", strMsgBod;

        private static DateTime dteStart;

        private static int intSFTPPort;
        private static bool getSFTPCredentials()
        {
            string strSFTPSettings;
            string[] strLines;

            strSFTPSettings = GlobalVariable.strConPath + "\\FTP_ConnectSettings.ini";
            
            dteStart = DateTime.Now;

            try
            {
                strLines = File.ReadAllLines(strSFTPSettings);

                strSFTPHost = strLines[0].ToString().Substring(strLines[0].IndexOf("=") + 1);
                strSFTPUserName = strLines[1].ToString().Substring(strLines[1].IndexOf("=") + 1);
                strSFTPPassword = strLines[2].ToString().Substring(strLines[2].IndexOf("=") + 1);
                intSFTPPort = Convert.ToInt32(strLines[3].ToString().Substring(strLines[3].IndexOf("=") + 1));

                strSFTPImpPath = strLines[4].ToString().Substring(strLines[4].IndexOf("=") + 1);
                strSFTPExpPath = strLines[5].ToString().Substring(strLines[5].IndexOf("=") + 1);

                strFTPSucRetPath = strLines[6].ToString().Substring(strLines[6].IndexOf("=") + 1);
                strFTPErrRetPath = strLines[7].ToString().Substring(strLines[7].IndexOf("=") + 1);

                if (string.IsNullOrEmpty(strSFTPHost))
                    return false;

            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("System", "Transfer File", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
                return false;
            }

            return true;
        }
        public static void transferArcFiles(string strProcess, string strFile)
        {
            string strFileNewPath = "";

            try
            {

                SystemInitialization.initFolders();

                GlobalVariable.strArcImpPath = GlobalVariable.strArcImpPath + GlobalVariable.strCompany + @"\";
                if (!Directory.Exists(GlobalVariable.strArcImpPath))
                    Directory.CreateDirectory(GlobalVariable.strArcImpPath);

                GlobalVariable.strAttExpPath = GlobalVariable.strAttExpPath + GlobalVariable.strCompany + @"\";
                if (!Directory.Exists(GlobalVariable.strAttExpPath))
                    Directory.CreateDirectory(GlobalVariable.strAttExpPath);

                if (strProcess == "Import")
                {                   
                    strFileNewPath = GlobalVariable.strArcImpPath + Path.GetFileName(strFile);
                    if (File.Exists(strFileNewPath))
                        strFileNewPath = GlobalVariable.strArcImpPath + Path.GetFileNameWithoutExtension(strFile) + "_" + DateTime.Now.ToString("MMddyyyyHHmm") + Path.GetExtension(strFile);

                }
                else
                {
                    strFileNewPath = GlobalVariable.strAttExpPath + Path.GetFileName(strFile);
                    if (File.Exists(strFileNewPath))
                        strFileNewPath = GlobalVariable.strAttExpPath + Path.GetFileNameWithoutExtension(strFile) + "_" + DateTime.Now.ToString("MMddyyyyHHmm") + Path.GetExtension(strFile);
                }

                File.Move(strFile, strFileNewPath);
            }
            catch (Exception ex)
            {
                strMsgBod = "Error transfering file. Please check error log.";

                GlobalFunction.sendAlert("E", strMsgSub, strMsgBod, GlobalVariable.oObjectType, GlobalVariable.oCompany.GetNewObjectKey().ToString());
                SystemFunction.transHandler(strProcess, "Transfer File", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }
        public static void transferProcFiles(string strProcess, string strStatus, string strFileName)
        {
            string strFilePath, strFileNewPath = "";

            try
            {

                SystemInitialization.initFolders();

                GlobalVariable.strImpSucPath = GlobalVariable.strImpSucPath + GlobalVariable.strCompany + @"\";
                if (!Directory.Exists(GlobalVariable.strImpSucPath))
                    Directory.CreateDirectory(GlobalVariable.strImpSucPath);

                GlobalVariable.strImpErrPath = GlobalVariable.strImpErrPath + GlobalVariable.strCompany + @"\";
                if (!Directory.Exists(GlobalVariable.strImpErrPath))
                    Directory.CreateDirectory(GlobalVariable.strImpErrPath);
                
                GlobalVariable.strExpSucPath = GlobalVariable.strExpSucPath + GlobalVariable.strCompany + @"\";
                if (!Directory.Exists(GlobalVariable.strExpSucPath))
                    Directory.CreateDirectory(GlobalVariable.strExpSucPath);
                 
                GlobalVariable.strExpErrPath = GlobalVariable.strExpErrPath + GlobalVariable.strCompany + @"\";
                if (!Directory.Exists(GlobalVariable.strExpErrPath))
                    Directory.CreateDirectory(GlobalVariable.strExpErrPath);
                

                if (strProcess == "Import")
                {
                    strFilePath = GlobalVariable.strImpPath + "/" + strFileName;

                    if (strStatus == "E")
                    {
                        strFileNewPath = GlobalVariable.strImpErrPath + strFileName;
                        if (File.Exists(strFileNewPath))
                            strFileNewPath = GlobalVariable.strImpErrPath + Path.GetFileNameWithoutExtension(strFilePath) + "_" + DateTime.Now.ToString("MMddyyyyHHmm") + Path.GetExtension(strFilePath);

                    }
                    else
                    {
                        strFileNewPath = GlobalVariable.strImpSucPath + strFileName;
                        if (File.Exists(strFileNewPath))
                            strFileNewPath = GlobalVariable.strImpSucPath + Path.GetFileNameWithoutExtension(strFilePath) + "_" + DateTime.Now.ToString("MMddyyyyHHmm") + Path.GetExtension(strFilePath);

                    }
                }
                else
                {
                    strFilePath = GlobalVariable.strExpPath + "/" + strFileName;

                    if (strStatus == "E")
                    {
                        strFileNewPath = GlobalVariable.strExpErrPath + strFileName;
                        if (File.Exists(strFileNewPath))
                            strFileNewPath = GlobalVariable.strExpErrPath + Path.GetFileNameWithoutExtension(strFilePath) + "_" + DateTime.Now.ToString("MMddyyyyHHmm") + Path.GetExtension(strFilePath);
                    }
                    else
                    {
                        strFileNewPath = GlobalVariable.strExpSucPath + strFileName;
                        if (File.Exists(strFileNewPath))
                            strFileNewPath = GlobalVariable.strExpSucPath + Path.GetFileNameWithoutExtension(strFilePath) + "_" + DateTime.Now.ToString("MMddyyyyHHmm") + Path.GetExtension(strFilePath);
                    }
                }


                File.Move(strFilePath, strFileNewPath);
            }
            catch (Exception ex)
            {
                strMsgBod = "Error transfering file. Please check error log.";

                GlobalFunction.sendAlert("E", strMsgSub, strMsgBod, GlobalVariable.oObjectType, GlobalVariable.oCompany.GetNewObjectKey().ToString());
                SystemFunction.transHandler(strProcess, "Transfer File", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }
        public static bool exportFTPFiles(string strFile, string strFileName, string strDir)
        {
            try
            {
                if (getSFTPCredentials())
                {
                    FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create(strSFTPHost + strSFTPExpPath + strDir + strFileName);

                    ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;

                    ftpRequest.Credentials = new NetworkCredential(strSFTPUserName, strSFTPPassword);

                    byte[] bytes = System.IO.File.ReadAllBytes(strFile);

                    ftpRequest.ContentLength = bytes.Length;
                    using (Stream UploadStream = ftpRequest.GetRequestStream())
                    {
                        UploadStream.Write(bytes, 0, bytes.Length);
                        UploadStream.Close();
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("System", "Transfer Export FTP File", "", "", "", "", dteStart, "E", "-111", string.Format("Error exporting {0} to FTP. {1}.", strFileName, ex.Message.ToString()));
                return false;
            }
        }
        public static bool importSFTPFiles(string strFileType)
        {
            try
            {
                if (getSFTPCredentials())
                {
                    SftpClient sftp = new SftpClient(strSFTPHost, intSFTPPort, strSFTPUserName, strSFTPPassword);
                    sftp.Connect();

                    var strImportFiles = sftp.ListDirectory(strSFTPImpPath);
                    foreach (var file in strImportFiles)
                    {
                        if (file.Equals(".") || file.Equals(".."))
                            continue;

                        if (file.Name.Contains(strFileType))
                        {
                            string strFilePath = GlobalVariable.strImpPath + file.Name;
                            if (File.Exists(strFilePath))
                                File.Delete(strFilePath);

                            using (Stream strmFilePath = File.OpenWrite(strFilePath))
                            {
                                sftp.DownloadFile(strSFTPImpPath + file.Name, strmFilePath);
                                sftp.DeleteFile(strSFTPImpPath + file.Name);
                            }
                        }
                    }

                    sftp.Disconnect();
                }

                return true;
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("System", "Import SFTP File", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
                return false;
            }
        }
        public static bool exportSFTPFiles(string strFilePath)
        {
            try
            {
                if (getSFTPCredentials())
                {
                    SftpClient sftp = new SftpClient(strSFTPHost, intSFTPPort, strSFTPUserName, strSFTPPassword);
                    sftp.Connect();

                    using (Stream strmFilePath = File.OpenRead(strFilePath))
                    {
                        sftp.UploadFile(strmFilePath, strSFTPExpPath + Path.GetFileName(strFilePath));
                    }

                    sftp.Disconnect();
                }

                return true;
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("System", "Export SFTP File", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
                return false;
            }
        }
        public static bool exportSFTPReturnFiles(string strFilePath, string strStatus)
        {
            try
            {
                if (getSFTPCredentials())
                {
                    SftpClient sftp = new SftpClient(strSFTPHost, intSFTPPort, strSFTPUserName, strSFTPPassword);
                    sftp.Connect();

                    using (Stream strmFilePath = File.OpenRead(strFilePath))
                    {
                        if (strStatus == "S")
                            sftp.UploadFile(strmFilePath, strFTPSucRetPath + Path.GetFileName(strFilePath));
                        else
                                sftp.UploadFile(strmFilePath, strFTPErrRetPath + Path.GetFileName(strFilePath));
                    }

                    sftp.Disconnect();
                }

                return true;
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("System", "Import FTP File", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
                return false;
            }
        }
        public static bool importFTPFiles(string strFileType)
        {
            try
            {
                if (getSFTPCredentials())
                {
                    List<String> files = GetFileFTPFileList(strFileType);

                    foreach (String strFileName in files)
                    {

                        int bytesRead = 0;
                        byte[] buffer = new byte[2048];

                        FtpWebRequest request = CreateFtpWebRequest(string.Format(@"{0}{1}/{2}/{3}", strSFTPHost, strSFTPImpPath, GlobalVariable.strCompany, strFileName), strSFTPUserName, strSFTPPassword, true);
                        request.Method = WebRequestMethods.Ftp.DownloadFile;

                        Stream reader = request.GetResponse().GetResponseStream();
                        FileStream fileStream = new FileStream(GlobalVariable.strImpPath + strFileName, FileMode.Create);

                        while (true)
                        {
                            bytesRead = reader.Read(buffer, 0, buffer.Length);

                            if (bytesRead == 0)
                                break;

                            fileStream.Write(buffer, 0, bytesRead);
                        }

                        fileStream.Close();

                        DeleteFTPFile(string.Format(@"{0}{1}/{2}/{3}", strSFTPHost, strSFTPImpPath, GlobalVariable.strCompany, strFileName));

                    }

                }

                return true;
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("System", "Import FTP File", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
                return false;
            }

        }
        private static void DeleteFTPFile(string strURI)
        {

            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(strURI);

                request.Method = WebRequestMethods.Ftp.DeleteFile;
                request.Credentials = new NetworkCredential(strSFTPUserName, strSFTPPassword);

                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                response.Close();
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("System", "Delete FTP File", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }
        private static FtpWebRequest CreateFtpWebRequest(string ftpDirectoryPath, string userName, string password, bool keepAlive = false)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(new Uri(ftpDirectoryPath));

            //Set proxy to null. Under current configuration if this option is not set then the proxy that is used will get an html response from the web content gateway (firewall monitoring system)
            request.Proxy = null;
            request.UsePassive = true;
            request.UseBinary = true;
            request.KeepAlive = keepAlive;

            request.Credentials = new NetworkCredential(userName, password);

            return request;
        }
        private static List<string> GetFileFTPFileList(string strFileType)
        {

            List<string> filelist = new List<string>();

            FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create(strSFTPHost + strSFTPImpPath + GlobalVariable.strCompany + @"/");
            ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;

            ftpRequest.Credentials = new NetworkCredential(strSFTPUserName, strSFTPPassword);

            FtpWebResponse ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();

            Stream responseStream = ftpResponse.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);

            String line = String.Empty;

            while((line = reader.ReadLine()) != null)
            {
                if (line.Contains(strFileType)) 
                {
                    filelist.Add(line);
                }
            }

            reader.Close();
            ftpResponse.Close();

            return filelist;
        }
        public static bool DeleteErrFiles(string strFilePath, string strFileName)
        {
            dteStart = DateTime.Now;

            try
            {
                string[] strErrFiles = Directory.GetFiles(strFilePath, strFileName);

                foreach (string strFile in strErrFiles)
                {
                    File.Delete(strFile);
                }
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("System", "Transfer File", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
                return false;
            }

            return true;
        }
    }
}
