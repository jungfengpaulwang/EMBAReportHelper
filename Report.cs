using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Text.RegularExpressions;
using System.Collections;
using Aspose.Cells;
using System.Linq;

namespace ReportHelper
{
    public class Report
    {
        public static Workbook Produce(Dictionary<string, List<DataSet>> allData, MemoryStream template)
        {
            return Produce(allData, template, true, null);
        }

        public static Workbook Produce(Dictionary<string, List<DataSet>> allData, MemoryStream template, bool AutoHPageBreak)
        {
            return Produce(allData, template, AutoHPageBreak, null);
        }

        /// <summary>
        /// 使用自訂樣版，產生單一檔案的自訂報表。
        /// 說明一：樣版的資料來源，請使用 DataSet 的 DataTable 儲存。
        /// 說明二：樣版中的所有變數，其名稱對應 DataSet 中的 DataTable Name。
        /// 說明三：以單一工作表的方式產生多張報表時，請將 DataSet 置入 List。
        /// 說明四：以多張工作表的方式產生報表時，請將 List<DataSet> 置入 Dictionary<string, List<DataSet>>，其中 Dictionary 的 Key 為工作表的名稱。
        /// 說明五：學號、姓名等簡單資料仍請使用 DataTable 儲存，唯僅有1欄、1列。
        /// 說明六：變數可置於含有固定文字的儲存格內，產生報表時，固定文字將保留，僅有變數被取代。
        /// 說明八：產生至報表的圖片，其格式請使用「Byte[]」，並設定儲存資料的 DataTable 之 DataColumn 的 DataType 為「System.Byte[]」。
        /// 說明九：若僅使用一個樣版不足以產生報表，則 DataSet 的 DataSetName 請以樣版名稱命名。(這樣才能辨識變數屬於哪個樣版)
        /// 注意事項一：若報表僅由一個樣版即足以呈現，請使用「DataSection」命名，反之請使用「ReportHeader、PageHeader、DataHeader、DataSection、DataFooter、PageFooter、ReportFooter」之結構。請參考：https://sites.google.com/a/ischool.com.tw/dev/k-reporthelper-yan-jiu-shi
        /// 注意事項二：請將程式與樣版所對應的變數登載於「樣版說明文件」，便於使用者自行套用，在報表中產生自訂的資料。
        /// 注意事項三：報表產生完畢，樣版即被刪除。
        /// </summary>
        /// <param name="allData">樣版的資料來源。</param>
        /// <param name="template">樣版檔。</param>
        /// <param name="AutoHPageBreak">DataSection 後自動換頁。</param>
        public static Workbook Produce(Dictionary<string, List<DataSet>> allData, MemoryStream template, bool AutoHPageBreak, Dictionary<CellObject, CellStyle> dicCellStyles)
        {
            if (allData == null)
                throw new Exception("未傳入資料，無法產生報表！");

            if (allData.Keys.Count == 0)
                throw new Exception("未傳入資料，無法產生報表！");

            //Workbook workbook = new Workbook();
            //workbook.Open(template);
            //  使用多形產生不同樣版型態之報表
            ExcelDocument reportBuilder;  
                        
            reportBuilder = ReportFactory.CreateExcelDocument<ExcelDocument>("ExcelTemplate");
            
            //  開始產生報表
            Workbook report = reportBuilder.Produce(allData, template, AutoHPageBreak, dicCellStyles);

            //  移除樣版檔：ReportHeader、PageHeader、DataHeader、DataSection、DataFooter、PageFooter、ReportFooter、樣版說明文件
            DocumentHelper.RemoveTemplateSheet(report);

            //  移除報表中的變數
            DocumentHelper.RemoveReportVariable(report);

            // 置換工作表名稱中的保留字
            report.Worksheets.Cast<Worksheet>().ToList().ForEach((x) =>
            {
                //  :\/?*[]
                //  System.Text.Encoding.UTF8.GetString((new byte[]{0xf9, 0x00, 0x00}).Select(s=>Convert.ToByte(s, 16)).ToArray())
                //byte[] bytes = { 0x00, 0x40, 0x00 };
                //Encoding enc = new UnicodeEncoding(false, true, true);

                //string value = enc.GetString(bytes);
                //List<string> utf8Bytes = new List<string>() { "F9", "00", "00" };
                //string value = System.Text.Encoding.UTF8.GetString(utf8Bytes.Select(s=>Convert.ToByte(s, 16)).ToArray());
                x.Name = x.Name.Replace("：", "꞉").Replace(":", "꞉").Replace("/", "⁄").Replace("／", "⁄").Replace(@"\", "∖").Replace("＼", "∖").Replace("?", "_").Replace("？", "_").Replace("*", "✻").Replace("＊", "✻").Replace("[", "〔").Replace("〔", "〔").Replace("]", "〕").Replace("〕", "〕");
            });
            return report;
        }
    }
}