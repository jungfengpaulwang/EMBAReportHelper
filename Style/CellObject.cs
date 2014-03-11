using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportHelper
{
    public class CellObject
    {
        public int RowIndex { set; get; }
        public int ColumnIndex { set; get; }
        public string TableName { set; get; }
        public string DataSetName { set; get; }
        public string WorkSheetName { set; get; }

        public CellObject(int RowIndex, int ColumnIndex, string TableName, string DataSetName, string WorkSheetName)
        {
            this.RowIndex = RowIndex;
            this.ColumnIndex = ColumnIndex;
            this.TableName = TableName;
            this.DataSetName = DataSetName;
            this.WorkSheetName = WorkSheetName;
        }
    }
}