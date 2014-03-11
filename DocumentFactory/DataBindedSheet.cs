using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.Data;

namespace ReportHelper
{
    public class DataBindedSheet
    {
        public Worksheet Worksheet { get; set; }
        public List<DataTable> DataTables { get; set; }
    }
}
