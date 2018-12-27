using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using log4net;
using Walkershop.EDI.Converter.Data.BusinessObjects;
using Walkershop.EDI.Converter.Data.ManagerObjects;

namespace Walkershop.EDI.Converter
{    
    public class EDIConverter
    {        
        static readonly object locker = new object();
        static EDIConverter instance = null;
        static ILog logger = log4net.LogManager.GetLogger(typeof(EDIConverter));

        private EDIConverter()
        {
        }
         
        public static EDIConverter GetInstance()
        {
            if (instance == null)
            {
                lock (locker)
                {
                    if (instance == null)
                    {
                        instance = new EDIConverter();
                    }
                }
            }

            return instance;
        }
 
        public void Run(int functionCode) 
        {            
            FucnctionType functionType = (FucnctionType)Enum.ToObject(typeof(FucnctionType), functionCode);

            logger.Info("-----------------------------------------------------------------------------------------------");
            logger.Info(string.Format("EDI Data Converter Interface starting to run program [{0}]{1} .....", functionCode, functionType.ToString()));
            var st = new System.Diagnostics.Stopwatch();
            st.Start();
            
            switch (functionType)
            {
                case FucnctionType.POR_IMP:                   
                    PorImport(WalkerApplication.SOURCE_PATH, WalkerApplication.SOURCE_FILE_EXTENSION);
                    break;

                case FucnctionType.POR_EXP:
                    PorExport();
                    break;

                case FucnctionType.DO_EXP:
                    DoExport();
                    break;

                default:                
                    break;
            }

            logger.Info(string.Format("Program [{0}]{1} done in {2} seconds.", functionCode, functionType.ToString(), st.Elapsed.TotalSeconds));
            logger.Info("-----------------------------------------------------------------------------------------------");
        }

        private void PorImport(string sourcePath, string sourceFileExtension)
        {        
            int totalOfFile = 0;
           
            string[] files = null;
                               
            if (System.IO.Directory.Exists(sourcePath))
            {
                string searchPattern = string.Format("*.{0}", sourceFileExtension);
                files = Directory.GetFiles(sourcePath, searchPattern);

                foreach (string file in files)
                {
                    totalOfFile++;
                    string fileName = Path.GetFileName(file);
                    string sourceFile = Path.Combine(sourcePath, fileName);
                    if (totalOfFile > 1)
                        logger.Info("****************************************************************************************************");
                    logger.Info(string.Format("Importing EDI file {0} Of {1} .....", totalOfFile, files.Length));
                    var st = new System.Diagnostics.Stopwatch();
                    st.Start();
                    ImportData2DB(sourceFile);
                    st.Stop();
                    logger.Info(string.Format("Import EDI file {0} Of {1} done in {2} seconds.", totalOfFile, files.Length, st.Elapsed.TotalSeconds));
                }
                
                logger.Info(string.Format("{0} files was imported.", totalOfFile));

            }
        
        }

        private void ImportData2DB(string sourceFileName)
        {
            IEDIDataManager ediDataManager = new EDIDataManager();
            DataTable dtEDIDataHeader = ediDataManager.CreateEDIDataHeader();
            DataTable dtEDIDataDetail = ediDataManager.CreateEDIDataDetail();

            var st = new System.Diagnostics.Stopwatch();
            st.Start();
            logger.Info(string.Format("{0}Step 1 Of 4: Reading EDI file \"{1}\" .....", Constant.STEP_BEGIN_TAB, sourceFileName));
            NOPIHelper.ImportData(sourceFileName, dtEDIDataHeader, dtEDIDataDetail);
            st.Stop();
            logger.Info(string.Format("{0}Read EDI file done in {1} seconds.", Constant.STEP_END_TAB, st.Elapsed.TotalSeconds));
            
            logger.Info(string.Format("{0}Step 2 Of 4: Verifying data .....", Constant.STEP_BEGIN_TAB));
            st.Start();
            
            //Verifying data handle
            ediDataManager.VerifyData(dtEDIDataHeader.DataSet, dtEDIDataDetail.DataSet);
            st.Stop();
            logger.Info(string.Format("{0}Verifying data done in {1}.", Constant.STEP_END_TAB, st.Elapsed.TotalSeconds));

            string destinationFileName = FileHelper.GetDestinationFileName(sourceFileName);
            foreach (DataRow dataRow in dtEDIDataHeader.Rows)
            {
                dataRow["BackupPath"] = destinationFileName;
            }
            logger.Info(string.Format("{0}Step 3 Of 4: Writing data to database .....", Constant.STEP_BEGIN_TAB));
            st.Start();
            ediDataManager.SaveOrUpdate(dtEDIDataHeader.DataSet, dtEDIDataDetail.DataSet);
            st.Stop();
            logger.Info(string.Format("{0}Write data to database done in {1} seconds.", Constant.STEP_END_TAB, st.Elapsed.TotalSeconds));
            
            logger.Info(string.Format("{0}Step 4 Of 4: Backuping source file .....", Constant.STEP_BEGIN_TAB));
            st.Start();
            FileHelper.BackupSourceFile(sourceFileName, destinationFileName);
            st.Stop();
            logger.Info(string.Format("{0}Backup source file done in {1} seconds.", Constant.STEP_END_TAB, st.Elapsed.TotalSeconds));   
                      
        }      

        private void PorExport()
        {                         
            var st = new System.Diagnostics.Stopwatch();
            logger.Info("Starting to get POR data .....");
            st.Start();
            IEDIDataManager dataManager = new EDIDataManager();            
            DataTable dtPOROut = dataManager.GetPOROutPut();
            if (dtPOROut.Rows.Count == 0)
            {
                logger.Info("No POR data for output file.");
                st.Stop();
                logger.Info(string.Format("Get POR data done in {0} seconds.", st.Elapsed.TotalSeconds));
                return;
            }

            st.Stop();
            logger.Info(string.Format("Get POR data done in {0} seconds.", st.Elapsed.TotalSeconds));

            //st.Start();

            //logger.Info("正在获取EDI导出规则 ...");
            //DataTable dtExpRule = ediPOManager.GetEDIExpRule("PO");
            //st.Stop();
            //logger.Info(string.Format("获取EDI导出规则信息完成，用时 {0} 秒.", st.Elapsed.TotalSeconds));

            logger.Info("Starting to write POR data to file .....");
            st.Start();
            
            StringBuilder strBuilder = new StringBuilder();
            int row = 0;
            string prevDoNo = "";
            string prevPoNo = "";
            //DataView dv = dtExpRule.DefaultView;
            //dv.RowFilter = "POHD = 'H'";
            //dv.Sort = "ItemNo";
            //DataTable dt = dv.ToTable();

            foreach (DataRow dataRow in dtPOROut.Rows)
            {
                row++;
                
                string currentDoNo = dataRow["RefNo"].ToString();
                string currentPoNo = dataRow["PoNo"].ToString();                 

                if (!string.IsNullOrEmpty(currentDoNo) && !string.IsNullOrEmpty(currentPoNo))
                {
                    if (currentDoNo != prevDoNo || currentPoNo != prevPoNo)
                    {
                        if (!string.IsNullOrEmpty(prevDoNo))
                            strBuilder.AppendLine();

                        // 应用 EDI 格式规则表, 未测试 
                        //int rowCount = 0;

                        //foreach (DataRow dr in dt.Rows)                     
                        //{
                        //    rowCount++;
                        //    string fldValue = "";
                        //    string itemType = dr["ItemType"] == DBNull.Value ? "" : dr["ItemType"].ToString();
                        //    string itemValue = dr["ItemValue"] == DBNull.Value ? "" : dr["ItemValue"].ToString();
                        //    string dataType = dr["DataType"] == DBNull.Value ? "" : dr["DataType"].ToString();   
                        //    if (itemType == "H")
                        //    {
                        //        fldValue = itemValue;
                        //    }
                        //    else if (itemType == "F") 
                        //    {
                        //        fldValue = dataRow[itemValue] == DBNull.Value ? "" : dataRow[itemValue].ToString();
                        //    }

                        //    if (dataType.ToUpper() == "DATE") 
                        //    {
                        //        fldValue = DateTime.Parse(fldValue).ToString("yyyyMMdd");
                        //    }

                        //    if (rowCount == dt.Rows.Count)
                        //        strBuilder.Append(string.Format("{0}", fldValue));
                        //    else
                        //        strBuilder.Append(string.Format("{0}{1}", fldValue, Constant.SEPARATOR_CHAR));
                        //}

                        strBuilder.Append(string.Format("{0}{1}", "M3_Header", Constant.SEPARATOR_CHAR));
                        strBuilder.Append(string.Format("{0}{1}", "WALKER", Constant.SEPARATOR_CHAR));
                        strBuilder.Append(string.Format("{0}{1}", "LOC0001", Constant.SEPARATOR_CHAR));
                        strBuilder.Append(string.Format("{0}{1}", DateTime.Parse(dataRow["DeliveryDate"].ToString()).ToString("yyyyMMdd"), Constant.SEPARATOR_CHAR));
                        strBuilder.Append(string.Format("{0}{1}", "POR", Constant.SEPARATOR_CHAR));
                        strBuilder.Append(string.Format("{0}", dataRow["RefNo"].ToString()));
                      
                        prevDoNo = currentDoNo;
                        prevPoNo = currentPoNo;
                    }
                }

                strBuilder.AppendLine();
                strBuilder.Append(string.Format("{0}{1}", "M3_Line", Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}{1}", dataRow["M3PoNo"].ToString(), Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}{1}", dataRow["M3ItemCode"].ToString(), Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}{1}", dataRow["M3MoreSysCode"].ToString(), Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}{1}", dataRow["TotalOfPairQty"].ToString(), Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}{1}", dataRow["RefNo"].ToString(), Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}{1}", "CNC", Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}{1}{2}", "WLK_", dataRow["Ship2Code"].ToString(), Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}{1}", "", Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}{1}", dataRow["ShipFlag"].ToString(), Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}{1}", "", Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}{1}", "", Constant.SEPARATOR_CHAR));
                strBuilder.Append(string.Format("{0}", dataRow["PoNo"].ToString()));
            }
            
            FileInfo fileInfo = FileHelper.PorOutputToFile(strBuilder.ToString());
            st.Stop();
            logger.Info(string.Format("Writing POR data to file {0} done in {1} seconds.", fileInfo.FullName, st.Elapsed.TotalSeconds));

            if (WalkerApplication.AUTO_UPLOAD_FTP)
            {
                bool success = FtpHelper.Upload(fileInfo.FullName, Path.Combine(WalkerApplication.FTP_POR_PATH.AbsoluteUri, fileInfo.Name), WalkerApplication.FTP_USER, WalkerApplication.FTP_PASSWORD, WalkerApplication.FTP_RETRY_TIMES);
                if (success)
                {
                    logger.Info("Backup destionation file .....");
                    File.Move(fileInfo.FullName, Path.Combine(WalkerApplication.DEST_BACKUP_PATH, fileInfo.Name));
                    logger.Info("Backup destionation file done.");
                 
                }
                else
                {
                    string pendingFile = Path.ChangeExtension(fileInfo.FullName, WalkerApplication.PENDING_FILE_EXTENSION);
                    File.Move(fileInfo.FullName, pendingFile);
                }
            }

            logger.Info("Update POR data status .....");
            st.Start();
            dataManager.UpdatePOStatus(ExpType.PO, fileInfo.FullName, dtPOROut.DataSet);
            st.Stop();
            logger.Info(string.Format("Update POR data status done in {0} seconds.", st.Elapsed.TotalSeconds));
        }

        private void DoExport()
        {
            var st = new System.Diagnostics.Stopwatch();
            
            logger.Info("Get DO data  .....");
            st.Start();
            IEDIDataManager ediPOManager = new EDIDataManager();            
            DataTable dtDOOut = ediPOManager.GetDOOutPut();
            if (dtDOOut.Rows.Count == 0) 
            {
                logger.Info("No DO data for output file.");
                st.Stop();
                logger.Info(string.Format("Get DO data done in {0} seconds.", st.Elapsed.TotalSeconds));
                return;
            }
            st.Stop();
            logger.Info(string.Format("Get DO data done in {0} seconds.", st.Elapsed.TotalSeconds));

            logger.Info("Writing DO data to file .....");
            st.Start();            
            StringBuilder strBuilder = new StringBuilder();

            string prevDoNo = "";
            string prevPoNo = "";

            foreach (DataRow dataRow in dtDOOut.Rows)
            {
                string currentDoNo = dataRow["RefNo"].ToString();
                string currentPoNo = dataRow["PoNo"].ToString();

                if (!string.IsNullOrEmpty(currentDoNo) && !string.IsNullOrEmpty(currentPoNo))
                {
                    if (currentDoNo != prevDoNo || currentPoNo != prevPoNo)
                    {
                        if (!string.IsNullOrEmpty(prevDoNo))
                            strBuilder.AppendLine();

                        strBuilder.Append(string.Format("{0}{1}", "M3_Header", Constant.SEPARATOR_CHAR));
                        strBuilder.Append(string.Format("{0}{1}", DateTime.Parse(dataRow["DeliveryDate"].ToString()).ToString("yyyyMMdd"), Constant.SEPARATOR_CHAR));
                        strBuilder.Append(string.Format("{0}{1}", "SOC", Constant.SEPARATOR_CHAR));
                        strBuilder.Append(string.Format("{0}{1}", dataRow["RefNo"].ToString(), Constant.SEPARATOR_CHAR));
                        strBuilder.Append(string.Format("{0}", dataRow["NoticeNo"].ToString()));

                        prevDoNo = currentDoNo;
                        prevPoNo = currentPoNo;
                    }

                    strBuilder.AppendLine();
                    strBuilder.Append(string.Format("{0}{1}", "M3_Line", Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("{0}{1}", dataRow["M3ItemCode"].ToString(), Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("{0}{1}", dataRow["M3MoreSysCode"].ToString(), Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("{0}{1}", dataRow["TotalOfPairQty"].ToString(), Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("{0}{1}", "WALKER", Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("WLK_{0}{1}", dataRow["Ship2Code"].ToString(), Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("{0}{1}", "CNC", Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("WLK_{0}{1}", dataRow["Ship2Code"].ToString(), Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("{0}{1}", dataRow["M3PoNo"].ToString(), Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("{0}{1}", "", Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("{0}{1}", dataRow["ShipFlag"].ToString(), Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("{0}{1}", "", Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("{0}{1}", "", Constant.SEPARATOR_CHAR));
                    strBuilder.Append(string.Format("{0}{1}", "", Constant.SEPARATOR_CHAR));
                }
            }
           
            FileInfo fileInfo = FileHelper.DoOutputToFile(strBuilder.ToString());
            st.Stop();
            logger.Info(string.Format("Writing DO data to file {0} done in {1} seconds.", fileInfo.FullName, st.Elapsed.TotalSeconds));

            if (WalkerApplication.AUTO_UPLOAD_FTP)
            {
                bool success = FtpHelper.Upload(fileInfo.FullName, Path.Combine(WalkerApplication.FTP_POR_PATH.AbsoluteUri, fileInfo.Name), WalkerApplication.FTP_USER, WalkerApplication.FTP_PASSWORD, WalkerApplication.FTP_RETRY_TIMES);

                if (success)
                {
                    logger.Info("Backup destionation file .....");
                    File.Move(fileInfo.FullName, Path.Combine(WalkerApplication.DEST_BACKUP_PATH, fileInfo.Name));
                    logger.Info("Backup destionation file done.");
                  
                }
                else
                {
                    string pendingFile = Path.ChangeExtension(fileInfo.FullName, WalkerApplication.PENDING_FILE_EXTENSION);
                    File.Move(fileInfo.FullName, pendingFile);
                }
            }

            logger.Info("Update DO data status .....");
            st.Start();
            ediPOManager.UpdatePOStatus(ExpType.DO, fileInfo.FullName, dtDOOut.DataSet);
            st.Stop();
            logger.Info(string.Format("Update DO data status done in {0} seconds.", st.Elapsed.TotalSeconds));
        }        
       

    }
}
