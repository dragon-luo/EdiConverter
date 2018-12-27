using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using log4net;

namespace Walkershop.EDI.Converter
{
    public class FileHelper
    {
        static ILog logger = log4net.LogManager.GetLogger(typeof(FileHelper));

        public static bool CreateDir(string dirName) 
        {
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(dirName);
                if (!dirInfo.Exists)
                {
                    logger.Info(string.Format("Directory \"{0}\" does not exist.", dirName));
                    dirInfo.Create();
                    logger.Info(string.Format("Create directory \"{0}\" done.", dirName));
                }
                return true;
            }
            catch (Exception ex) 
            {
                logger.Error(ex.Message);
                logger.Error(ex.StackTrace);
                return false;
            }
        }

        public static bool CreateSourceBackupDir() 
        {   
            string date = DateTime.Now.ToString(Constant.DATE_FORMAT);
            string sourceBackupDir = Path.Combine(WalkerApplication.FILE_ROOT_DIR, Constant.BACKUP_DIR);
            sourceBackupDir = Path.Combine(sourceBackupDir, Constant.SOURCE_DIR);
            sourceBackupDir = Path.Combine(sourceBackupDir, date);
            WalkerApplication.SOURCE_BACKUP_PATH = sourceBackupDir;
            return CreateDir(sourceBackupDir);
        }

        public static bool CreateDestBackupDir()
        {
            string date = DateTime.Now.ToString(Constant.DATE_FORMAT);
            string destBackupDir = Path.Combine(WalkerApplication.FILE_ROOT_DIR, Constant.BACKUP_DIR);
            destBackupDir = Path.Combine(destBackupDir, Constant.DEST_DIR);
            destBackupDir = Path.Combine(destBackupDir, date);
            WalkerApplication.DEST_BACKUP_PATH = destBackupDir;
            return CreateDir(destBackupDir);
        }

        public static bool CreateSourceDir() 
        {
            string sourceDir = Path.Combine(WalkerApplication.FILE_ROOT_DIR, Constant.SOURCE_DIR);
            WalkerApplication.SOURCE_PATH = sourceDir;
            return CreateDir(sourceDir);
        }

        public static bool CreateDestDir() 
        {
            string destDir = Path.Combine(WalkerApplication.FILE_ROOT_DIR, Constant.DEST_DIR);
            WalkerApplication.DEST_PATH = destDir;
            return CreateDir(destDir);
        }

        public static string GetDestinationFileName(string sourceFileName) 
        {
            FileInfo fileInfo = new FileInfo(sourceFileName);
            string fileName = fileInfo.Name;

            int pos = fileName.IndexOf(".xls");
            if (pos > 0)
            {
                string ts = DateTime.Now.ToString(Constant.BACKUP_FILE_SUFFIX_FORMAT);
                fileName = fileName.Insert(pos, string.Format(".{0}", ts));
            }

            string destinationFileName = Path.Combine(WalkerApplication.SOURCE_BACKUP_PATH, fileName);

            return destinationFileName;
        }

        public static void BackupSourceFile(string sourceFileName, string destinationFileName)
        {
            FileInfo fileInfo = new FileInfo(sourceFileName);           

            fileInfo.MoveTo(destinationFileName);

            logger.Info(string.Format("{0}  Backup file \"{1}\" to \"{2}\".", Constant.STEP_END_TAB, sourceFileName, destinationFileName));            
        }

        private static FileInfo OutputToFile(String fileContent, string fileExtension)
        {            
            string name = DateTime.Now.ToString(Constant.BACKUP_FILE_SUFFIX_FORMAT);
            string fullFileName = string.Format("{0}_{2}{1}{2}", name, ".", fileExtension);
            string destionationFileName = Path.Combine(WalkerApplication.DEST_PATH, fullFileName);             

            if (!File.Exists(destionationFileName))
            {
                FileStream fs = new FileStream(destionationFileName, FileMode.Create, FileAccess.Write);                
                StreamWriter sw = new StreamWriter(fs);
                sw.Write(fileContent);                
                sw.Close();
                fs.Close();
            }
            else
            {
                OutputToFile(fileContent, fileExtension);
            }

            string backupDestDir = WalkerApplication.DEST_BACKUP_PATH;
            if (!Directory.Exists(backupDestDir))
            {
                Directory.CreateDirectory(backupDestDir);
            }

            return new FileInfo(destionationFileName);
        }

        public static FileInfo PorOutputToFile(string fileContent) 
        {
            return OutputToFile(fileContent, WalkerApplication.POR_OUTPUT_FILE_EXTENSION);
        }

        public static FileInfo DoOutputToFile(string fileContent) 
        {
            return OutputToFile(fileContent, WalkerApplication.DO_OUTPUT_FILE_EXTENSION);
        }


    }
}
