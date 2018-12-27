using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Walkershop.EDI.Converter
{
    public class Constant
    {

        public const char SEPARATOR_CHAR = '|';
        public const string DATE_FORMAT = "yyyyMMdd";
        public const string BACKUP_FILE_SUFFIX_FORMAT = "yyyyMMddHHmmssf";
        public const int MORE_SYS_CODE_COL_INDEX = 2;        
        public const int SUPPLIER_COL_INDEX = 2;
        public const int CUSTOMER_COL_INDEX = 1;
        public const int DELIVERY_DATE_COL_INDEX = 1;
        public const int REF_NO_ROW_COL_INDEX = 1;
        public const int SHIP_TO_ADDR_COL_INDEX = 6;
        public const int SHIP_TO_CODE_COL_INDEX = 6;
        public const int NOTICE_NO_COL_INDEX = 15;

        public const string SOURCE_DIR = "SOURCE";
        public const string DEST_DIR = "DEST";
        public const string BACKUP_DIR = "BACKUP";

        public const string STEP_BEGIN_TAB = "    ";
        public const string STEP_END_TAB = "                 ";
        public const string STEP_DETAIL_TAB = "                     ";

    }

    public enum FucnctionType
    {
        POR_IMP = 1,
        POR_EXP = 2,
        DO_EXP = 3,
        EXIT = 4,
    }

    public enum ExpType
    {
        PO = 1,
        DO = 2,
    }

}
