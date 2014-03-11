using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Aspose.Cells;
using System.IO;

namespace ReportHelper
{
    interface ExcelDocument
    {
        Workbook Produce(Dictionary<string, List<DataSet>> allData, MemoryStream template, bool AutoHPageBreak, Dictionary<CellObject, CellStyle> dicCellStyles);
    }
}