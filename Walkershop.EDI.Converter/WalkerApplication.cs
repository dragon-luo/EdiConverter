using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.IO;
using log4net;

namespace Walkershop.EDI.Converter
{
    public class WalkerApplication
    {   
        public static string CONNECTION_STRING_NAME;
        public static string BACKUP_PATH;
        public static string SOURCE_PATH;
        public static string DEST_PATH;
        public static string SOURCE_BACKUP_PATH;
        public static string DEST_BACKUP_PATH;        
        public static string POR_OUTPUT_FILE_EXTENSION;
        public static string DO_OUTPUT_FILE_EXTENSION;
        public static string PENDING_FILE_EXTENSION;
        public static string FILE_ROOT_DIR;
        public static string SOURCE_FILE_EXTENSION;
        
        //FTP
        public static bool AUTO_UPLOAD_FTP;
        public static string FTP_USER;
        public static string FTP_PASSWORD;
        public static Uri FTP_POR_PATH;
        public static Uri FTP_DO_PATH;
        public static int FTP_RETRY_TIMES = 10;

        static ILog logger = log4net.LogManager.GetLogger(typeof(WalkerApplication));

        static WalkerApplication()
        {
            try
            {
                CONNECTION_STRING_NAME = ConfigurationManager.AppSettings["connectionString"].Trim();
                POR_OUTPUT_FILE_EXTENSION = ConfigurationManager.AppSettings["POR_OUTPUT_FILE_EXTENSION"].Trim();
                DO_OUTPUT_FILE_EXTENSION = ConfigurationManager.AppSettings["DO_OUTPUT_FILE_EXTENSION"].Trim();
                FILE_ROOT_DIR = ConfigurationManager.AppSettings["FILE_ROOT_DIR"].Trim();
                PENDING_FILE_EXTENSION = ConfigurationManager.AppSettings["PENDING_FILE_EXTENSION"].Trim();
                SOURCE_FILE_EXTENSION = ConfigurationManager.AppSettings["SOURCE_FILE_EXTENSION"].Trim();
                try
                {
                    AUTO_UPLOAD_FTP = Boolean.Parse(ConfigurationManager.AppSettings["AUTO_UPLOAD_FTP"].Trim());
                }
                catch 
                {
                    AUTO_UPLOAD_FTP = false;
                }

                try
                {
                    FTP_RETRY_TIMES = int.Parse(ConfigurationManager.AppSettings["FTP_RETRY_TIMES"].Trim());
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message);
                    logger.Error(ex.StackTrace);
                }

                FTP_USER = ConfigurationManager.AppSettings["FTP_USER"].Trim();
                FTP_PASSWORD = ConfigurationManager.AppSettings["FTP_PASSWORD"].Trim();
                FTP_POR_PATH = new Uri(ConfigurationManager.AppSettings["FTP_POR_PATH"].Trim());
                FTP_DO_PATH = new Uri(ConfigurationManager.AppSettings["FTP_DO_PATH"].Trim());

                FileHelper.CreateSourceDir();
                FileHelper.CreateDestDir();
                FileHelper.CreateSourceBackupDir();
                FileHelper.CreateDestBackupDir();
            }
            catch (Exception ex) 
            {
                logger.Error(ex.Message);
                logger.Error(ex.StackTrace);
            }
        }

    }
}
