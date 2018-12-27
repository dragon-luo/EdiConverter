using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Walkershop.EDI.Converter.Data.BusinessObjects;

namespace Walkershop.EDI.Converter
{
    public class XlsColumn
    {
        const string MATERIAL_COLUMN_NAME = "材料";
        const string ROWTYPE_COLUMN_NAME = "箱數";

        //映射电子表格栏头到数据库表字段名
        public static void Map2DB(DataTable dtXLSColumnHeader)
        {            
            for (int j = 0; j < dtXLSColumnHeader.Columns.Count; j++)
            {
                string name = dtXLSColumnHeader.Columns[j].ColumnName;
                string columnName = "";

                if (name.IndexOf("訂單編號") > 0 || name.IndexOf("商標") > 0)
                {
                    columnName = "PoNo";
                }
                else if (name.IndexOf("廠號") > 0)
                {
                    columnName = "FactoryNo";
                }
                else if (name.IndexOf("旧货号") > 0 || name.IndexOf("貨品編號") > 0)
                {
                    columnName = "MoreSysCode";
                }
                else if (name.IndexOf("材料") > 0)
                {
                    columnName = "Material";
                }
                else if (name.IndexOf("配") > 0 && name.IndexOf("码") > 0)
                {
                    columnName = "SizeAssort";
                }
                else if (name.IndexOf("箱數") > 0)
                {
                    columnName = "RowType";
                }
                else if (name.IndexOf("包裝") > 0)
                {
                    columnName = "PackageQty";
                }
                else if (name.IndexOf("對數") > 0)
                {
                    columnName = "PairQty";
                }
                else if (name.IndexOf("箱号") > 0)
                {
                    columnName = "BoxNo";
                }
                else 
                {
                    columnName = name;
                }
                
                dtXLSColumnHeader.Columns[j].ColumnName = columnName;                
                
            }
          
            //修正配码栏位名
            int columnIndex = -1;
            if (!dtXLSColumnHeader.Columns.Contains("SizeAssort"))
            {                
                //获取材料栏位位置
                columnIndex = dtXLSColumnHeader.Columns.IndexOf("Material");

                //材料栏位后面即为配码栏位
                dtXLSColumnHeader.Columns[columnIndex + 1].ColumnName = "SizeAssort";
            }

            //修正行类型和箱数栏位名
            columnIndex = dtXLSColumnHeader.Columns.IndexOf("RowType");
            dtXLSColumnHeader.Columns[columnIndex + 1].ColumnName = "BoxQty";

            dtXLSColumnHeader.Columns.Add("RefNo", typeof(string));
            dtXLSColumnHeader.Columns.Add("PageNo", typeof(int));
            dtXLSColumnHeader.Columns.Add("ItemNo", typeof(int));
        }

    }
}
