using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data;
using log4net;

namespace Walkershop.EDI.Converter.Data.ManagerObjects          
{
    interface IEDIDataManager
    {
        DataTable CreateEDIDataHeader();
        DataTable CreateEDIDataDetail();
        bool SaveOrUpdate(DataSet dsEDIDataHeader, DataSet dsEDIDataDetail);

        DataTable GetPOROutPut();
        DataTable GetDOOutPut();
        void UpdatePOStatus(ExpType expType, string fileName, DataSet dsEDIPO);
        DataTable GetEDIExpRule(string expType);

        void VerifyData(DataSet dsEDIDataHeader, DataSet dsEDIDataDetail);
        
    }

    public class EDIDataManager : ManagerBase, IEDIDataManager
    {
        const string PROC_CREATE_EDI_DATA_HEADER = "proc_CreateEDIDataHeader";
        const string PROC_CREATE_EDI_DATA_DETAIL = "proc_CreateEDIDataDetail";
        const string PROC_UPDATE_EDI_DATA_HEADER = "proc_UpdtEDIDataHeader";
        const string PROC_UPDATE_EDI_DATA_DETAIL = "proc_UpdtEDIDataDetail";
        const string PROC_UPDATE_EDI_DATA_STATUS = "proc_UpdtEDIDataStatus";

        const string PROC_VERIFY_EDI_DATA_HEADER = "proc_VerifyEDIDataHeader";
        const string PROC_VERIFY_EDI_DATA_DETAIL = "proc_VerifyEDIDataDetail";

        const string PROC_GET_POR_OUT_PUT = "proc_GetPOROutPut";
        const string PROC_GET_DO_OUT_PUT = "proc_GetDOOutPut";
        const string EDI_DATA_HEADER_TABLE_NAME = "EDIDataHeader";
        const string EDI_DATA_DETAIL_TABLE_NAME = "EDIDataDetail";

        static ILog logger = log4net.LogManager.GetLogger(typeof(EDIDataManager));

        public DataTable CreateEDIDataHeader()
        {   
            using (IDbConnection connection = CurrentDB.CreateConnection())
            {
                connection.Open();

                try
                {
                    DbCommand selectCommand = CurrentDB.GetStoredProcCommand(PROC_CREATE_EDI_DATA_HEADER);
                    DataSet dsEDIDataHeader = CurrentDB.ExecuteDataSet(selectCommand);
                    DataTable dtEDIDataHeader = dsEDIDataHeader.Tables[0];
                    dtEDIDataHeader.TableName = EDI_DATA_HEADER_TABLE_NAME;
                    return dtEDIDataHeader;
                }
                catch 
                {
                    throw;
                }
                finally
                {
                    if (connection != null)
                    {
                        if (connection.State == ConnectionState.Open)
                            connection.Close();
                    }
                }            
            }
        }

        public DataTable CreateEDIDataDetail()
        {         
            using (IDbConnection connection = CurrentDB.CreateConnection())
            {
                connection.Open();

                try
                {
                    DbCommand selectCommand = CurrentDB.GetStoredProcCommand(PROC_CREATE_EDI_DATA_DETAIL);
                    DataSet dsEDIDataDetail = CurrentDB.ExecuteDataSet(selectCommand);
                    DataTable dtEDIDataDetail = dsEDIDataDetail.Tables[0];
                    dtEDIDataDetail.TableName = EDI_DATA_DETAIL_TABLE_NAME;
                    return dtEDIDataDetail;
                }
                catch
                {
                    throw;
                }
                finally
                {
                    if (connection != null)
                    {
                        if (connection.State == ConnectionState.Open)
                            connection.Close();
                    }
                }
            }
        }

        public bool SaveOrUpdate(DataSet dsEDIDataHeader, DataSet dsEDIDataDetail)
        {
            DbCommand ediDataHeaderInsertCommand = CurrentDB.GetStoredProcCommand(PROC_UPDATE_EDI_DATA_HEADER);
            CurrentDB.AddInParameter(ediDataHeaderInsertCommand, "@RefNo", DbType.String, "RefNo", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataHeaderInsertCommand, "@NoticeNo", DbType.String, "NoticeNo", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataHeaderInsertCommand, "@Supplier", DbType.String, "Supplier", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataHeaderInsertCommand, "@DeliveryDate", DbType.DateTime, "DeliveryDate", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataHeaderInsertCommand, "@Ship2Code", DbType.String, "Ship2Code", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataHeaderInsertCommand, "@Ship2Addr", DbType.String, "Ship2Addr", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataHeaderInsertCommand, "@SourceFile", DbType.String, "SourceFile", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataHeaderInsertCommand, "@SourceSheet", DbType.String, "SourceSheet", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataHeaderInsertCommand, "@BackupPath", DbType.String, "BackupPath", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataHeaderInsertCommand, "@Remark", DbType.String, "Remark", DataRowVersion.Current);  
            
            DbCommand ediDataDetailInsertCommand = CurrentDB.GetStoredProcCommand(PROC_UPDATE_EDI_DATA_DETAIL);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@RefNo", DbType.String, "RefNo", DataRowVersion.Current);            
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@PoNo", DbType.String, "PoNo", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@MoreSysCode", DbType.String, "MoreSysCode", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@ItemNo", DbType.Int32, "ItemNo", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@FactoryNo", DbType.String, "FactoryNo", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@Material", DbType.String, "Material", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SizeAssort", DbType.String, "SizeAssort", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA34", DbType.Int32, "SA34", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA35", DbType.Int32, "SA35", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA36", DbType.Int32, "SA36", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA37", DbType.Int32, "SA37", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA38", DbType.Int32, "SA38", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA39", DbType.Int32, "SA39", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA40", DbType.Int32, "SA40", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA41", DbType.Int32, "SA41", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA42", DbType.Int32, "SA42", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA43", DbType.Int32, "SA43", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA44", DbType.Int32, "SA44", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA45", DbType.Int32, "SA45", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA46", DbType.Int32, "SA46", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA47", DbType.Int32, "SA47", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA48", DbType.Int32, "SA48", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA49", DbType.Int32, "SA49", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@SA50", DbType.Int32, "SA50", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@RowType", DbType.String, "RowType", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@BoxQty", DbType.Int32, "BoxQty", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@PackageQty", DbType.Int32, "PackageQty", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@PairQty", DbType.Int32, "PairQty", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@BoxNo", DbType.String, "BoxNo", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediDataDetailInsertCommand, "@Remark", DbType.String, "Remark", DataRowVersion.Current);

            using (IDbConnection connection = CurrentDB.CreateConnection())
            {
                connection.Open();
                IDbTransaction transaction = connection.BeginTransaction();
                try
                {
                  
                    CurrentDB.UpdateDataSet(dsEDIDataHeader, dsEDIDataHeader.Tables[0].TableName, ediDataHeaderInsertCommand, null, null, (DbTransaction)transaction);
                    CurrentDB.UpdateDataSet(dsEDIDataDetail, dsEDIDataDetail.Tables[0].TableName, ediDataDetailInsertCommand, null, null, (DbTransaction)transaction);                    
                  
                    transaction.Commit();
                    
                    return true;
                }
                catch  
                {
                    transaction.Rollback();             
                    throw;
                }
                finally
                {
                    if (connection != null)
                    {
                        if (connection.State == ConnectionState.Open)
                            connection.Close();
                    }
                }

            }
        }

        public DataTable GetPOROutPut()
        {            
            DbCommand selectCommand = CurrentDB.GetStoredProcCommand(PROC_GET_POR_OUT_PUT);            
            DataSet dsPOROut = CurrentDB.ExecuteDataSet(selectCommand);
            dsPOROut.Tables[0].TableName = "POROutPut";
            return dsPOROut.Tables[0];
        }

        public DataTable GetDOOutPut()
        {            
            DbCommand selectCommand = CurrentDB.GetStoredProcCommand(PROC_GET_DO_OUT_PUT);            
            DataSet dsPOROut = CurrentDB.ExecuteDataSet(selectCommand);
            dsPOROut.Tables[0].TableName = "DOOutPut";
            return dsPOROut.Tables[0];
        }

        public DataTable GetEDIExpRule(string expType)
        {
            string storedProc = "proc_GetEDIExpRule";
            DbCommand selectCommand = CurrentDB.GetStoredProcCommand(storedProc);
            CurrentDB.AddInParameter(selectCommand, "@ExpType", DbType.String, expType);
            DataSet dsPOROut = CurrentDB.ExecuteDataSet(selectCommand);
            dsPOROut.Tables[0].TableName = "EXPRULE";
            return dsPOROut.Tables[0];
        }

        public void UpdatePOStatus(ExpType expType, string fileName, DataSet dsEDIPO)
        {
            int status = 10;
            string poDestFileName = "";
            string doDestFileName = "";

            if (expType == ExpType.PO)
            {
                status = 20;
                poDestFileName = fileName;
            }
            else
            {
                status = 30;
                doDestFileName = fileName;
            }

            foreach (DataRow dataRow in dsEDIPO.Tables[0].Rows)
            {
                dataRow["Status"] = status;
                
                if (!string.IsNullOrEmpty(poDestFileName))
                    dataRow["POutPutFile"] = poDestFileName;

                if (!string.IsNullOrEmpty(doDestFileName))                
                    dataRow["DOutPutFile"] = doDestFileName;
            }

            DbCommand ediPoUpdateCommand = CurrentDB.GetStoredProcCommand(PROC_UPDATE_EDI_DATA_STATUS);
            ediPoUpdateCommand.CommandTimeout = 100;            
            CurrentDB.AddInParameter(ediPoUpdateCommand, "@RefNo", DbType.String, "RefNo", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediPoUpdateCommand, "@POutPutFile", DbType.String, "POutPutFile", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediPoUpdateCommand, "@DOutPutFile", DbType.String, "DOutPutFile", DataRowVersion.Current);
            CurrentDB.AddInParameter(ediPoUpdateCommand, "@Status", DbType.Int32, "Status", DataRowVersion.Current);
            CurrentDB.UpdateDataSet(dsEDIPO, dsEDIPO.Tables[0].TableName, null, ediPoUpdateCommand, null, UpdateBehavior.Transactional);
        }

        public void VerifyData(DataSet dsEDIDataHeader, DataSet dsEDIDataDetail)
        {
            foreach (DataRow dr in dsEDIDataHeader.Tables[0].Rows)
            {
                string refNo = dr["RefNo"].ToString();
                DbCommand ediDataHeaderVerifyCommand = CurrentDB.GetStoredProcCommand(PROC_VERIFY_EDI_DATA_HEADER);
                CurrentDB.AddInParameter(ediDataHeaderVerifyCommand, "@RefNo", DbType.String, refNo);               
                CurrentDB.ExecuteNonQuery(ediDataHeaderVerifyCommand);                            
        
                DataRow[] rows = dsEDIDataDetail.Tables[0].Select(string.Format("RefNo='{0}'", refNo));
                foreach (DataRow drDetail in rows)
                {
                    string poNo = drDetail["PoNo"].ToString();
                    string moreSysCode = drDetail["MoreSysCode"].ToString();
                    DbCommand ediDataDetailVerifyCommand = CurrentDB.GetStoredProcCommand(PROC_VERIFY_EDI_DATA_DETAIL);
                    //CurrentDB.AddInParameter(ediDataDetailVerifyCommand, "@RefNo", DbType.String, refNo);    
                    CurrentDB.AddInParameter(ediDataDetailVerifyCommand, "@PoNo", DbType.String, poNo);
                    CurrentDB.AddInParameter(ediDataDetailVerifyCommand, "@MoreSysCode", DbType.String, moreSysCode);
                    CurrentDB.ExecuteNonQuery(ediDataDetailVerifyCommand);
                }

            }
        }

        string GetFieldValue(DataRow dataRow, string fieldName)
        {
            object obj = dataRow[fieldName];

            if (obj == null)
            {
                return string.Empty;
            }
            else 
            {
                return obj.ToString();
            }
        }

    }

}
