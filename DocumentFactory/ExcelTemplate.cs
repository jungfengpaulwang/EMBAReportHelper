using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace ReportHelper
{
    class ExcelTemplate : ExcelDocument
    {
        public Workbook Produce(Dictionary<string, List<DataSet>> allData, MemoryStream template)
        {
            return Produce(allData, template, true, null);
        }

        public Workbook Produce(Dictionary<string, List<DataSet>> allData, MemoryStream template, bool AutoHPageBreak, Dictionary<CellObject, CellStyle> dicCellStyles)
        {
            //  樣版檔
            Workbook workbook = new Workbook();
            workbook.Open(template);

            //  樣版
            List<string> templates = new List<string>();
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                templates.Add(workbook.Worksheets[i].Name);
            }

            bool hasReportHeader = templates.Contains("ReportHeader");
            bool hasPageHeader = templates.Contains("PageHeader");
            bool hasDataHeader = templates.Contains("DataHeader");
            bool hasDataSection = templates.Contains("DataSection");
            bool hasDataFooter = templates.Contains("DataFooter");
            bool hasPageFooter = templates.Contains("PageFooter");
            bool hasReportFooter = templates.Contains("ReportFooter");

            foreach (KeyValuePair<string, List<DataSet>> kv in allData)
            {
                Worksheet instanceSheet = null;
                int i = 0;
                // 複製樣版：ReportHeader        
                if (hasReportHeader)
                {
                    int instanceSheetIndex = workbook.Worksheets.AddCopy("ReportHeader");
                    instanceSheet = workbook.Worksheets[instanceSheetIndex];
                    instanceSheet.Name = kv.Key;

                    // 以資料來源替代變數：ReportHeader
                    foreach (DataSet dataSet in kv.Value)
                    {
                        if (dataSet.DataSetName.ToUpper() == "REPORTHEADER")
                            DocumentHelper.GenerateSheet(dataSet, instanceSheet, i, dicCellStyles);
                    }

                    i = instanceSheet.Cells.MaxRow + 1;
                }

                // 複製樣版：PageHeader        
                if (hasPageHeader)
                {
                    if (instanceSheet == null)
                    {
                        int instanceSheetIndex = workbook.Worksheets.AddCopy("PageHeader");
                        instanceSheet = workbook.Worksheets[instanceSheetIndex];
                        instanceSheet.Name = kv.Key;
                    }
                    else
                        DocumentHelper.CloneTemplate(instanceSheet, workbook.Worksheets["PageHeader"], i);

                    // 設定 PageHeader 重覆列印區塊，即「版面配置」的「列印標題」


                    instanceSheet.PageSetup.PrintTitleColumns = workbook.Worksheets["PageHeader"].PageSetup.PrintTitleColumns;
                    instanceSheet.PageSetup.PrintTitleRows = workbook.Worksheets["PageHeader"].PageSetup.PrintTitleRows;

                    // 以資料來源替代變數：PageHeader
                    foreach (DataSet dataSet in kv.Value)
                    {
                        if (dataSet.DataSetName.ToUpper() == "PAGEHEADER")
                            DocumentHelper.GenerateSheet(dataSet, instanceSheet, i, dicCellStyles);
                    }

                    i = instanceSheet.Cells.MaxRow + 1;
                }

                //  複製樣版：DataHeader
                if (instanceSheet == null && hasDataHeader)
                {
                    int instanceSheetIndex = workbook.Worksheets.AddCopy("DataHeader");
                    instanceSheet = workbook.Worksheets[instanceSheetIndex];
                    instanceSheet.Name = kv.Key;
                }

                //  複製樣版：DataSection
                if (instanceSheet == null && hasDataSection)
                {
                    int instanceSheetIndex = workbook.Worksheets.AddCopy("DataSection");
                    instanceSheet = workbook.Worksheets[instanceSheetIndex];
                    instanceSheet.Name = kv.Key;
                }

                foreach (DataSet dataSet in kv.Value)
                {
                    if (dataSet.DataSetName.ToUpper() == "DATASECTION")
                    {
                        //  複製樣版：DataHeader，暫不支援變數
                        if (hasDataHeader)
                        {
                            DocumentHelper.CloneTemplate(instanceSheet, workbook.Worksheets["DataHeader"], i);

                            i = instanceSheet.Cells.MaxRow + 1;
                        }

                        //  複製樣版：DataSection
                        if (hasDataSection)
                        {
                            if (i > 0 && !hasPageFooter && AutoHPageBreak == true)
                                instanceSheet.HPageBreaks.Add(i, instanceSheet.Cells.MaxColumn);

                            DocumentHelper.CloneTemplate(instanceSheet, workbook.Worksheets["DataSection"], i);
                            // 以資料來源替代變數：DataSection
                            DocumentHelper.GenerateSheet(dataSet, instanceSheet, i, dicCellStyles);
                            i = instanceSheet.Cells.MaxRow + 1;
                        }

                        //  複製樣版：DataFooter，暫不支援變數
                        if (hasDataFooter)
                        {
                            DocumentHelper.CloneTemplate(instanceSheet, workbook.Worksheets["DataFooter"], i);

                            i = instanceSheet.Cells.MaxRow + 1;
                            instanceSheet.HPageBreaks.Add(i, instanceSheet.Cells.MaxColumn);
                        }
                    }
                }

                // 複製樣版：PageFooter     
                if (hasPageFooter)
                {
                    DocumentHelper.CloneTemplate(instanceSheet, workbook.Worksheets["PageFooter"], i);

                    // 以資料來源替代變數：PageFooter
                    foreach (DataSet dataSet in kv.Value)
                    {
                        if (dataSet.DataSetName.ToUpper() == "PAGEFOOTER")
                            DocumentHelper.GenerateSheet(dataSet, instanceSheet, i, dicCellStyles);
                    }

                    i = instanceSheet.Cells.MaxRow + 1;
                }

                // 複製樣版：ReportFooter     
                if (hasReportFooter)
                {
                    DocumentHelper.CloneTemplate(instanceSheet, workbook.Worksheets["ReportFooter"], i);

                    // 以資料來源替代變數：ReportFooter
                    foreach (DataSet dataSet in kv.Value)
                    {
                        if (dataSet.DataSetName.ToUpper() == "REPORTFOOTER")
                            DocumentHelper.GenerateSheet(dataSet, instanceSheet, i, dicCellStyles);
                    }

                    i = instanceSheet.Cells.MaxRow + 1;
                }

                // 整份文件水平置中
                //instanceSheet.PageSetup.CenterHorizontally = true;
            }
            return workbook;
        }
    }
}
