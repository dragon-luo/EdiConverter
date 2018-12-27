using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Configuration;
using System.Collections;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.POIFS.FileSystem;
using log4net;
using Walkershop.EDI.Converter.Data.BusinessObjects;

namespace Walkershop.EDI.Converter
{
    public class NOPIHelper
    {                    
        static string fileName;
        static HSSFWorkbook workbook;        
        static ILog logger = log4net.LogManager.GetLogger(typeof(NOPIHelper));
       
        //页码
        static int pageNumber = 0;

        //数据行计数器
        static int rowIndex = 0;

        //是否为数据行
        static bool isDetail = false;

        //数据行号
        static int itemNo = 0;

        //数据总行数
        static int totalOfDetail = 0;

        static EDIDataHeader ediDataHeader = null;  

        static NOPIHelper()
        {        
        }

        public static void WorkbookInitialize(string fullFileName)
        {
            using (FileStream fileStream = new FileStream(fullFileName, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(fileStream);
            }

            FileInfo fileInfo = new FileInfo(fullFileName);
            fileName = fileInfo.Name;

        }

        private static void WorkBookDispose()
        {            
            workbook = null;
        }

        private static IEnumerator GetRowCollection(int sheetIndex) 
        {
            ISheet sheet = workbook.GetSheetAt(sheetIndex);

            IEnumerator rows = sheet.GetRowEnumerator();

            return rows;
        }       

        private static int GetColumnCount(int sheetIndex) 
        {
            ISheet sheet = workbook.GetSheetAt(sheetIndex);
            IRow headerRow = sheet.GetRow(0);
            return headerRow.LastCellNum;
        }

        private static int GetRowCount(int sheetIndex) 
        {
            ISheet sheet = workbook.GetSheetAt(sheetIndex);            
            return sheet.LastRowNum;
        }

        public static void ImportData(string fullFileName, DataTable dtEDIDataHeader, DataTable dtEDIDataDetail)
        {
            int sheetIndex = 0;

            try
            {                
                WorkbookInitialize(fullFileName);
                for (sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
                {
                    string sheetName = workbook.GetSheetName(sheetIndex);
          
                    IEnumerator rows = GetRowCollection(sheetIndex);

                    logger.Info(string.Format("{0}Processing sheet {1} of {2}, sheet name [{3}] .....", Constant.STEP_DETAIL_TAB, sheetIndex + 1, workbook.NumberOfSheets, sheetName));
                    ImportDataFromSheet(sheetName, rows, dtEDIDataHeader, dtEDIDataDetail);
                    logger.Info(string.Format("{0}Processing sheet {1} of {2} done.", Constant.STEP_DETAIL_TAB, sheetIndex + 1, workbook.NumberOfSheets));
                }
            }
            catch
            {
                logger.Info(string.Format("{0}Processing sheet {1} of {2} failed.", Constant.STEP_DETAIL_TAB, sheetIndex + 1, workbook.NumberOfSheets));
                throw;
            }
            finally
            {
                WorkBookDispose();
            }
        }

        private static void InitializePage() 
        {
            pageNumber++;

            // 初始化数据行信息
            isDetail = false;
            totalOfDetail = 0;
            itemNo = 0;

            ediDataHeader = new EDIDataHeader();
            ediDataHeader.Remark = String.Empty;
        }

        private static DataTable SaveXLSColumnHeader2Tab(IRow row)
        {
            DataTable dtXlsColumnHeader = new DataTable();
            
            //XLS COLUMN HEADER
            for (int i = 0; i < row.LastCellNum; i++)
            {
                ICell cell = row.GetCell(i);

                object obj;
                if (cell == null)
                {
                    obj = string.Empty;
                }
                else
                {
                    obj = cell.ToString();
                }

                dtXlsColumnHeader.Columns.Add(string.Format("C{0}_{1}", i + 1, obj), typeof(string));
            }

            return dtXlsColumnHeader;
          
        }

        private static void ImportDataFromSheet(string sheetName, IEnumerator rows, DataTable dtEDIDataHeader, DataTable dtEDIDataDetail)
        {           
            string prevRowPoNo = "";            
            string prevRowFactoryNo = "";
            string prevRowMoreSysCode = "";
            string prevRowMaterial = "";
            bool addItem = false;
                      
            DataRow xlsDataRow = null;
            DataTable dtXLSColumnHeader = new DataTable();          
            string customerName = "";

            while (rows.MoveNext())
            {
                rowIndex++;

                int max_cell_num = 0;
                string firstCellValueOfRow = "";
                string lastCellValueOfRow = "";                

                IRow row = (HSSFRow)rows.Current;

                //当前行单元格数
                max_cell_num = row.LastCellNum;
                
                //if (firstCellValueOfRow == null) continue;

                //当前行首个单元格值
                firstCellValueOfRow = GetStringCellValue(row, row.FirstCellNum);              

                //当前行最后单元格值
                lastCellValueOfRow = GetStringCellValue(row, max_cell_num - 1);
                if (firstCellValueOfRow.Contains("申請出貨")) 
                {
                    customerName = GetStringCellValue(row, Constant.CUSTOMER_COL_INDEX).ToString(); 
                }                          
                else if (firstCellValueOfRow.Contains("供應商"))
                {
                    //新页初始化
                    InitializePage();

                    ediDataHeader.Supplier = GetStringCellValue(row, Constant.SUPPLIER_COL_INDEX).ToString();
                    if (string.IsNullOrEmpty(ediDataHeader.Supplier)) 
                    {
                        ediDataHeader.Supplier = GetStringCellValue(row, Constant.SUPPLIER_COL_INDEX - 1).ToString();
                    }

                    if (string.IsNullOrEmpty(ediDataHeader.Supplier))
                    {
                        ediDataHeader.Supplier = customerName;
                    }
                }
                else if (firstCellValueOfRow.Contains("出貨日"))
                {
                    ediDataHeader.DeliveryDate = row.GetCell(Constant.DELIVERY_DATE_COL_INDEX).DateCellValue;
                    ediDataHeader.Ship2Addr = GetStringCellValue(row, Constant.SHIP_TO_ADDR_COL_INDEX).ToString();
                }
                else if (firstCellValueOfRow.Contains("自編出貨編號"))
                {                    
                    ediDataHeader.RefNo = GetStringCellValue(row, Constant.REF_NO_ROW_COL_INDEX).ToString();
                    ediDataHeader.Ship2Code = GetStringCellValue(row, Constant.SHIP_TO_CODE_COL_INDEX).ToString();
                    ediDataHeader.NoticeNo = GetStringCellValue(row, Constant.NOTICE_NO_COL_INDEX).ToString();
                    ediDataHeader.SourceFile = fileName;
                    ediDataHeader.SourceSheet = sheetName;
                    ediDataHeader.BackupPath = WalkerApplication.BACKUP_PATH;              
                } 
                else if (lastCellValueOfRow == "箱号")
                {                                     
                    DataRow dataHeaderRow = dtEDIDataHeader.NewRow();
                   
                    dataHeaderRow["RefNo"] = ediDataHeader.RefNo;
                    dataHeaderRow["NoticeNo"] = ediDataHeader.NoticeNo;
                    dataHeaderRow["Supplier"] = ediDataHeader.Supplier;
                    dataHeaderRow["DeliveryDate"] = ediDataHeader.DeliveryDate;
                    dataHeaderRow["Ship2Code"] = ediDataHeader.Ship2Code;
                    dataHeaderRow["Ship2Addr"] = ediDataHeader.Ship2Addr;
                    dataHeaderRow["Status"] = 10;
                    dataHeaderRow["SourceFile"] = ediDataHeader.SourceFile;
                    dataHeaderRow["SourceSheet"] = ediDataHeader.SourceSheet;                    
                    dataHeaderRow["BackupPath"] = ediDataHeader.BackupPath;
                    if (!string.IsNullOrEmpty(ediDataHeader.RefNo) && !string.IsNullOrEmpty(ediDataHeader.NoticeNo))
                    {
                        dtEDIDataHeader.Rows.Add(dataHeaderRow);
                    }
                    else
                    {
                        logger.Error(string.Format("SheetName:{2} RefNo: {0} NoticeNo: {1}", ediDataHeader.RefNo, ediDataHeader.NoticeNo, sheetName));
                    }

                    //如果数据行是明细头，则保存栏位名
                    isDetail = true;

                    dtXLSColumnHeader = SaveXLSColumnHeader2Tab(row);

                    XlsColumn.Map2DB(dtXLSColumnHeader);
                }            
               
                //开始读取明细行数据
                if (isDetail)
                {
                    int columnIndex = -1;
                    columnIndex = dtXLSColumnHeader.Columns.IndexOf("RowType");
                    string rowType = GetStringCellValue(row, columnIndex);

                    columnIndex = dtXLSColumnHeader.Columns.IndexOf("PackageQty");
                    bool isSubTotal = GetStringCellValue(row, columnIndex) == "箱";
                    bool isColumnHeader = GetStringCellValue(row, columnIndex) == "箱數";
                    if (isSubTotal || isColumnHeader)
                        continue;

                    //读取 [出貨] 数据行
                    if (rowType == "出貨")
                    {
                        
                        totalOfDetail++;                                            
                        xlsDataRow = dtXLSColumnHeader.NewRow();                                          

                        for (int j = row.FirstCellNum; j < max_cell_num; j++)
                        {
                            object cellValue = (object)GetStringCellValue(row, j);
                            xlsDataRow[j] = cellValue == null ? "" : cellValue.ToString().Trim();
                        }
                        
                        string currRowPoNo = xlsDataRow["PoNo"] == DBNull.Value ? "" : xlsDataRow["PoNo"].ToString();
                        string currRowFactoryNo = xlsDataRow["FactoryNo"] == DBNull.Value ? "" : xlsDataRow["FactoryNo"].ToString();
                        string currRowMoreSysCode = xlsDataRow["MoreSysCode"] == DBNull.Value ? "" : xlsDataRow["MoreSysCode"].ToString();
                        string currRowMaterial = xlsDataRow["Material"] == DBNull.Value ? "" : xlsDataRow["Material"].ToString();
                        string currRowSizeAssort = xlsDataRow["SizeAssort"] == DBNull.Value ? "" : xlsDataRow["SizeAssort"].ToString();

                        if (string.IsNullOrEmpty(currRowPoNo) && !string.IsNullOrEmpty(prevRowPoNo))
                        {
                            currRowPoNo = prevRowPoNo;
                            currRowFactoryNo = prevRowFactoryNo;
                        }

                        if (string.IsNullOrEmpty(currRowMoreSysCode) && !string.IsNullOrEmpty(prevRowMoreSysCode))
                        {
                            currRowMoreSysCode = prevRowMoreSysCode;
                            currRowMaterial = prevRowMaterial;
                        }

                        if (!string.IsNullOrEmpty(currRowMoreSysCode) && !string.IsNullOrEmpty(currRowSizeAssort))
                        {
                            if (!currRowPoNo.Equals(prevRowPoNo))
                            {
                                if (string.IsNullOrEmpty(currRowFactoryNo))
                                {
                                    currRowFactoryNo = prevRowFactoryNo;
                                }

                                itemNo = 0;
                            }

                            itemNo++;

                            xlsDataRow["RefNo"] = ediDataHeader.RefNo;
                            xlsDataRow["PoNo"] = currRowPoNo;
                            xlsDataRow["ItemNo"] = itemNo;
                            xlsDataRow["FactoryNo"] = currRowFactoryNo;
                            xlsDataRow["MoreSysCode"] = currRowMoreSysCode;
                            xlsDataRow["Material"] = currRowMaterial;
                            xlsDataRow["PageNo"] = pageNumber;
                                                   
                            prevRowPoNo = currRowPoNo;
                            prevRowFactoryNo = currRowFactoryNo;
                            prevRowMoreSysCode = currRowMoreSysCode;
                            prevRowMaterial = currRowMaterial;
                                
                            addItem = true;
                        }
                    }
                    else 
                    {
                        if (addItem)
                        {
                            string moreSysCode = GetStringCellValue(row, Constant.MORE_SYS_CODE_COL_INDEX);
                            if (prevRowMoreSysCode.Contains("/") || !string.IsNullOrEmpty(moreSysCode))
                            {
                                xlsDataRow["MoreSysCode"] = moreSysCode;
                                prevRowMoreSysCode = moreSysCode;
                            }
                            
                            DataRow dataRow = dtEDIDataDetail.NewRow();
                            for (int i = 0; i < dtXLSColumnHeader.Columns.Count; i++) 
                            {
                                string xlsColumnName = dtXLSColumnHeader.Columns[i].ColumnName;
                                string value = xlsDataRow[xlsColumnName].ToString().Trim();
                                string fldName = string.Empty;

                                if (dtEDIDataDetail.Columns.Contains(xlsColumnName))
                                {
                                    fldName = xlsColumnName;                                  
                                }
                                else
                                {
                                    int pos = xlsColumnName.IndexOf("_");
                                    string name = xlsColumnName.Substring(pos + 1);
                                    string columnName = string.Format("SA{0}", name);
                                    if (dtEDIDataDetail.Columns.Contains(columnName))
                                    {
                                        fldName = columnName;
                                    }
                                }
                                
                                if (!string.IsNullOrEmpty(fldName))
                                {
                                    if (dtEDIDataDetail.Columns[fldName].DataType == typeof(System.Int32))
                                    {
                                        if (string.IsNullOrEmpty(value))
                                            dataRow[fldName] = DBNull.Value;
                                        else
                                            dataRow[fldName] = int.Parse(value);
                                    }
                                    else
                                    {
                                        dataRow[fldName] = value;
                                    }
                                }

                            }

                            dtEDIDataDetail.Rows.Add(dataRow);
                            addItem = false;
                        }
                    }
                    
                }
                                              
            }     

        }
 
        static string GetStringCellValue(IRow row, int cellIndex)
        {
            object cellValue = string.Empty;

            ICell cell = row.GetCell(cellIndex);
            if (cell == null)
            {                
                cellValue = string.Empty;
            }
            else
            {               
                switch (cell.CellType)
                {
                    case CellType.Boolean:
                        cellValue = cell.BooleanCellValue;
                        break;

                    case CellType.Numeric:
                        cellValue = cell.ToString();
                        break;
                 
                    case CellType.String:
                        cellValue = cell.StringCellValue;
                        break;

                    case CellType.Error:
                        cellValue = cell.ErrorCellValue;
                        break;
                    
                    case CellType.Blank:
                        cellValue = string.Empty;
                        break;

                    case CellType.Formula:                   
                        switch (cell.CachedFormulaResultType)
                        {
                            case CellType.Boolean:
                                cellValue = cell.BooleanCellValue;
                                break;

                            case CellType.Error:
                                cellValue = cell.ErrorCellValue;
                                break;

                            case CellType.Numeric:
                                if (DateUtil.IsCellDateFormatted(cell))
                                {
                                    cellValue = cell.DateCellValue.ToString(Constant.DATE_FORMAT);
                                }
                                else
                                {
                                    cellValue = cell.NumericCellValue;
                                }
                                break;

                            case CellType.String:                                
                                string str = cell.StringCellValue;
                                
                                if (!string.IsNullOrEmpty(str))
                                {
                                    cellValue = str.ToString();
                                }
                                else
                                {
                                    cellValue = string.Empty;
                                }
                                break;

                            case CellType.Unknown:
                            case CellType.Blank:
                            default:
                                cellValue = string.Empty;
                                break;
                        }
                        break;

                    default:
                        if (cell == null)
                            cellValue = string.Empty;
                        else                        
                            cellValue = cell.ToString();
                        break;
                }
            }

            return cellValue == null ? (string)null : cellValue.ToString();
             
        }

    }
 
}
