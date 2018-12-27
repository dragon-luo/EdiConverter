using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Walkershop.EDI.Converter.Data.BusinessObjects
{
    public class EDIDataHeader
    {
        private string refNo;
        private string noticeNo;
        private string supplier;
        private DateTime deliveryDate;
        private string ship2Code;
        private string ship2Addr;
        private int status;
        private string sourceFile;
        private string sourceSheet;
        private string pOutputFile;
        private string dOutputFile;
        private string backupPath;
        private string remark;

        public string RefNo
        {
            get { return refNo; }
            set { refNo = value; }
        }

        public string Supplier
        {
            get { return supplier; }
            set { supplier = value; }
        }        

        public DateTime DeliveryDate
        {
            get { return deliveryDate; }
            set { deliveryDate = value; }
        }
        
        public string Ship2Code
        {
            get { return ship2Code; }
            set { ship2Code = value; }
        }        

        public string Ship2Addr
        {
            get { return ship2Addr; }
            set { ship2Addr = value; }
        }

        public string NoticeNo
        {
            get { return noticeNo; }
            set { noticeNo = value; }
        }

        public int Status
        {
            get { return status; }
            set { status = value; }
        }

        public string SourceFile
        {
            get { return sourceFile; }
            set { sourceFile = value; }
        }

        public string SourceSheet
        {
            get { return sourceSheet; }
            set { sourceSheet = value; }
        }

        public string POutputFile
        {
            get { return pOutputFile; }
            set { pOutputFile = value; }
        }

        public string DOutputFile
        {
            get { return dOutputFile; }
            set { dOutputFile = value; }
        }

        public string BackupPath
        {
            get { return backupPath; }
            set { backupPath = value; }
        }

        public string Remark
        {
            get { return remark; }
            set { remark = value; }
        }

    }
}
