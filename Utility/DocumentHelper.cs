using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace ReportHelper
{
    public class DocumentHelper
    {
        public enum PrintTitle
	    {
	        PrintTitleColumns, PrintTitleRows
	    }

        public static string ShiftPrintTitle(string title, int offset, PrintTitle printTitle)
        {
            if (string.IsNullOrEmpty(title))
            {
                if (offset > 0)
                    if (printTitle == PrintTitle.PrintTitleColumns)
                        return string.Empty;
                    else
                        return string.Empty;
                else
                    return string.Empty;
            }

            if (!title.Contains(":"))
                return string.Empty;

            string newPrintTitle = string.Empty;

            return string.Empty;
        }

        public static string NumberToAlpha(int number)
        {   
            //  A~Z：65~90
            if (65 <= number && 90 >= number)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] btNumber = new byte[] { (byte)number };
                return asciiEncoding.GetString(btNumber);
            }
            throw new Exception("傳入不正確的數字");
        }

        public static char Chr(int Num)
        {
            char C = Convert.ToChar(Num);

            return C;
        }

        public static int ASC(string S)
        {
            int N = Convert.ToInt32(S[0]);

            return N;
        }

        public static int ASC(char C)
        {
            int N = Convert.ToInt32(C);

            return N;
        }

        public static void RemoveReportVariable(Workbook wb)
        {
            foreach(Worksheet ws in wb.Worksheets)
            {
                foreach (Cell c in ws.Cells)
                {
                    if (c.Value == null)
                        continue;

                    Regex rx = new Regex(@"\[\[.*?\]\]");
                    MatchCollection m = rx.Matches(c.Value.ToString());

                    if (m.Count > 0)
                    {
                        for (int k = 0; k < m.Count; k++)
                        {
                            string newValue = c.Value.ToString().Replace(m[k].Value.ToString(), "");

                            c.PutValue(newValue);
                        }
                    }
                }
            }
        }

        //  Formatting Selected Characters in a Cell
        private static void RestoreSelectedCharactersFormatting(Font Font_Applied, dynamic Font_ApplyBy)
        {
            Font_ApplyBy = Font_ApplyBy as Font;

            if (Font_ApplyBy == null)
                return;

            Font_Applied.Color = Font_ApplyBy.Color;
            Font_Applied.IsBold = Font_ApplyBy.IsBold;
            Font_Applied.IsItalic = Font_ApplyBy.IsItalic;
            Font_Applied.IsStrikeout = Font_ApplyBy.IsStrikeout;
            Font_Applied.IsSubscript = Font_ApplyBy.IsSubscript;
            Font_Applied.IsSuperscript = Font_ApplyBy.IsSuperscript;
            Font_Applied.Name = Font_ApplyBy.Name;
            Font_Applied.Size = Font_ApplyBy.Size;
            Font_Applied.Underline = Font_ApplyBy.Underline;
        }
        
        //  Formatting Selected Characters in a Cell
        private static Dictionary<int, dynamic> BackupSelectedCharactersFormatting(Cell cell)
        {
            Dictionary<int, dynamic> dicCharactersFormatting = new Dictionary<int, dynamic>();
            if (cell == null || cell.Value == null)
                return dicCharactersFormatting;

            if (cell.Value.GetType().Name == "String")
            {
                for (int k = 0; k < (cell.Value + "").Length; k++)
                {
                    dynamic o = new ExpandoObject();

                    o.Color = cell.Characters(k, 1).Font.Color;
                    o.IsBold = cell.Characters(k, 1).Font.IsBold;
                    o.IsItalic = cell.Characters(k, 1).Font.IsItalic;
                    o.IsStrikeout = cell.Characters(k, 1).Font.IsStrikeout;
                    o.IsSubscript = cell.Characters(k, 1).Font.IsSubscript;
                    o.IsSuperscript = cell.Characters(k, 1).Font.IsSuperscript;
                    o.Name = cell.Characters(k, 1).Font.Name;
                    o.Size = cell.Characters(k, 1).Font.Size;
                    o.Underline = cell.Characters(k, 1).Font.Underline;

                    dicCharactersFormatting.Add(k, o);
                }
            }

            return dicCharactersFormatting;
        }

        /// <summary>
        /// Aspose 版複製樣版
        /// </summary>
        /// <param name="instanceSheet">報表</param>
        /// <param name="templateSheet">樣版</param>
        /// <param name="dataIndex">報表目前資料位置</param>
        public static void CloneTemplate(Aspose.Cells.Worksheet instanceSheet, Aspose.Cells.Worksheet templateSheet, int dataIndex)
        {            
            //步驟一：以 CopyRow 的方式複製樣版
            int index = dataIndex;
            for (int i = templateSheet.Cells.MinRow; i <= templateSheet.Cells.MaxRow; i++)
            {
                if (templateSheet.Cells.Rows[i] == null)
                    continue;

                //  複制 Data
                instanceSheet.Cells.CopyRow(templateSheet.Cells, i, index);
                //  複制 Style
                instanceSheet.Cells.Rows[index].Style.Copy(templateSheet.Cells.Rows[i].Style);
                //  Formatting Selected Characters in a Cell
                //for (int j = instanceSheet.Cells.MinColumn; j <= instanceSheet.Cells.MaxColumn; j++)
                //{
                //    if (instanceSheet.Cells[index, j] == null)
                //        continue;

                //    if (instanceSheet.Cells[index, j].Value == null)
                //        continue;

                //    Cell cell = templateSheet.Cells[i, j];
                //    if (cell == null || cell.Value == null)
                //        continue;
                //    if (instanceSheet.Cells[index, j].Value.GetType().Name == "String")
                //    {
                //        for (int k = 0; k < (instanceSheet.Cells[index, j].Value + "").Length; k++)
                //        {
                //            instanceSheet.Cells[index, j].Characters(k, 1).Font.Color = cell.Characters(k, 1).Font.Color;
                //            instanceSheet.Cells[index, j].Characters(k, 1).Font.IsBold = cell.Characters(k, 1).Font.IsBold;
                //            instanceSheet.Cells[index, j].Characters(k, 1).Font.IsItalic = cell.Characters(k, 1).Font.IsItalic;
                //            instanceSheet.Cells[index, j].Characters(k, 1).Font.IsStrikeout = cell.Characters(k, 1).Font.IsStrikeout;
                //            instanceSheet.Cells[index, j].Characters(k, 1).Font.IsSubscript = cell.Characters(k, 1).Font.IsSubscript;
                //            instanceSheet.Cells[index, j].Characters(k, 1).Font.IsSuperscript = cell.Characters(k, 1).Font.IsSuperscript;
                //            instanceSheet.Cells[index, j].Characters(k, 1).Font.Name = cell.Characters(k, 1).Font.Name;
                //            instanceSheet.Cells[index, j].Characters(k, 1).Font.Size = cell.Characters(k, 1).Font.Size;
                //            instanceSheet.Cells[index, j].Characters(k, 1).Font.Underline = cell.Characters(k, 1).Font.Underline;
                //        }
                //    }
                //}
                // 若不手動設定列高，則列高為預設值
                instanceSheet.Cells.SetRowHeight(index, templateSheet.Cells.GetRowHeight(i));
                index++;
            }

            //步驟二：合併儲存格並複制 Formatting Selected Characters in a Cell
            foreach (Aspose.Cells.CellArea cellArea in templateSheet.Cells.MergedCells)
            {                
                instanceSheet.Cells.Merge(cellArea.StartRow + dataIndex, cellArea.StartColumn, (cellArea.EndRow - cellArea.StartRow + 1), cellArea.EndColumn - cellArea.StartColumn + 1);
            }
            //index = dataIndex;
            //for (int i = instanceSheet.Cells.MinRow; i <= instanceSheet.Cells.MaxRow; i++)
            //{
            //    if (instanceSheet.Cells.Rows[i] == null)
            //        continue;
                
            //    for (int j = instanceSheet.Cells.MinColumn; j <= instanceSheet.Cells.MaxColumn; j++)
            //    {
            //        if (instanceSheet.Cells[index, j] == null)
            //            continue;

            //        if (instanceSheet.Cells[index, j].Value == null)
            //            continue;

            //        Cell cell = templateSheet.Cells[i, j];
            //        if (cell == null || cell.Value == null)
            //            continue;

            //        Range mergeRange = instanceSheet.Cells[index, j].GetMergedRange();
            //        if (mergeRange != null)
            //        {
            //            Cell mergeCell = instanceSheet.Cells[mergeRange.FirstRow, mergeRange.FirstColumn];
            //            for (int k = 0; k < (mergeCell.Value + "").Length; k++)
            //            {
            //                mergeCell.Characters(k, 1).Font.Color = cell.Characters(k, 1).Font.Color;
            //                mergeCell.Characters(k, 1).Font.IsBold = cell.Characters(k, 1).Font.IsBold;
            //                mergeCell.Characters(k, 1).Font.IsItalic = cell.Characters(k, 1).Font.IsItalic;
            //                mergeCell.Characters(k, 1).Font.IsStrikeout = cell.Characters(k, 1).Font.IsStrikeout;
            //                mergeCell.Characters(k, 1).Font.IsSubscript = cell.Characters(k, 1).Font.IsSubscript;
            //                mergeCell.Characters(k, 1).Font.IsSuperscript = cell.Characters(k, 1).Font.IsSuperscript;
            //                mergeCell.Characters(k, 1).Font.Name = cell.Characters(k, 1).Font.Name;
            //                mergeCell.Characters(k, 1).Font.Size = cell.Characters(k, 1).Font.Size;
            //                mergeCell.Characters(k, 1).Font.Underline = cell.Characters(k, 1).Font.Underline;
            //            }
            //        }
            //    }
            //    index++;
            //}
            //步驟三：複製圖片
            foreach (Aspose.Cells.Picture imgTemplate in templateSheet.Pictures)
            {
                try
                {
                    instanceSheet.Pictures.Add(imgTemplate.UpperLeftRow + dataIndex, imgTemplate.UpperLeftColumn, imgTemplate.LowerRightRow + dataIndex, imgTemplate.LowerRightColumn, new MemoryStream(GetPictureData(imgTemplate)));

                    instanceSheet.Pictures[instanceSheet.Pictures.Count - 1].UpperDeltaX = imgTemplate.UpperDeltaX;
                    instanceSheet.Pictures[instanceSheet.Pictures.Count - 1].UpperDeltaY = imgTemplate.UpperDeltaY;
                    instanceSheet.Pictures[instanceSheet.Pictures.Count - 1].LowerDeltaX = imgTemplate.LowerDeltaX;
                    instanceSheet.Pictures[instanceSheet.Pictures.Count - 1].LowerDeltaY = imgTemplate.LowerDeltaY;
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }
        }
        
        /// <summary>
        /// 取得 Aspose.Cells.Picture 的 Image Data
        /// </summary>
        /// <param name="picture">Aspose.Cells.Picture</param>
        /// <returns>Aspose.Cells.Picture 的 Image Data(格式為 byte[])</returns>
        public static byte[] GetPictureData(Aspose.Cells.Picture picture)
        {
            try
            {
                PropertyInfo p = picture.GetType().GetProperty("x90c6e45466e5b849", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static);

                object obj = p.GetValue(picture, null);

                p = obj.GetType().GetProperty("xe36c96d8c564b382", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static);

                return (p.GetValue(obj, null) as byte[]);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        
        /// <summary>
        /// 移除樣版檔的所有樣版工作表
        /// </summary>
        /// <param name="wb">樣版檔</param>
        public static void RemoveTemplateSheet(Aspose.Cells.Workbook wb)
        {
            List<Worksheet> Worksheets = wb.Worksheets.Cast<Worksheet>().ToList();
            foreach (Worksheet sheet in Worksheets)
            {
                if (sheet.Name.ToUpper() == "REPORTHEADER")
                    wb.Worksheets.RemoveAt(sheet.Name);
                if (sheet.Name.ToUpper() == "PAGEHEADER")
                    wb.Worksheets.RemoveAt(sheet.Name);
                if (sheet.Name.ToUpper() == "DATAHEADER")
                    wb.Worksheets.RemoveAt(sheet.Name);
                if (sheet.Name.ToUpper() == "DATASECTION")
                    wb.Worksheets.RemoveAt(sheet.Name);
                //if (sheet.Name.ToUpper() == "DATASECTIONNOPAGEBREAK")
                //    wb.Worksheets.RemoveAt(sheet.Name);
                if (sheet.Name.ToUpper() == "DATAFOOTER")
                    wb.Worksheets.RemoveAt(sheet.Name);
                if (sheet.Name.ToUpper() == "PAGEFOOTER")
                    wb.Worksheets.RemoveAt(sheet.Name);
                if (sheet.Name.ToUpper() == "REPORTFOOTER")
                    wb.Worksheets.RemoveAt(sheet.Name);
                if (sheet.Name.ToUpper() == "樣版說明文件")
                    wb.Worksheets.RemoveAt(sheet.Name);
            }
        }

        /// <summary>
        /// 目前僅支援 Aspose.Cells 報表。
        /// </summary>
        /// <param name="document">要儲存的報表物件。</param>
        /// <param name="fileFullName">儲存的檔案名稱，請記得加入副檔名(.xls)。</param>
        /// <param name="openAfterSave">產生完報表是否由程式開啟。批次產生報表時，務必傳入 false。</param>
        /// <param name="type">目前僅支援 DocumentType.HSSFWorkbook，未來將支援 DocumentType.PDF。</param>
        public static void Save(Aspose.Cells.Workbook document, string fileFullName, bool openAfterSave, DocumentType type)
        {
            try
            {
                document.Save(fileFullName);

                if (openAfterSave)
                    Open(fileFullName);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// 探索工作表包含變數的儲存格。
        /// </summary>
        /// <param name="sheet">包含變數的工作表，通常為「樣版」。</param>
        public static Dictionary<string, List<Aspose.Cells.Cell>> DiscoverVariableCells(Aspose.Cells.Worksheet sheet, int dataIndex)
        {
            Dictionary<string, List<Aspose.Cells.Cell>> variableCollection = new Dictionary<string, List<Aspose.Cells.Cell>>();

            for (int i = dataIndex; i <= sheet.Cells.MaxRow; i++)
            {
                if (sheet.Cells.Rows[i] == null)
                    continue;

                for (int j = sheet.Cells.MinColumn; j <= sheet.Cells.MaxColumn; j++)
                {
                    if (sheet.Cells[i, j] == null)
                        continue; 
                    
                    if (sheet.Cells[i, j].Value == null)
                        continue;

                    try
                    {
                        Aspose.Cells.Cell cell = sheet.Cells[i, j];

                        Regex rx = new Regex(@"\[\[.*?\]\]");
                        MatchCollection m = rx.Matches(cell.Value.ToString());

                        if (m.Count > 0)
                        {
                            for (int k = 0; k < m.Count; k++)
                            {
                                string keyWord = m[k].Value.Replace("[[", "").Replace("]]", "");

                                if (!variableCollection.ContainsKey(keyWord))
                                    variableCollection.Add(keyWord, new List<Aspose.Cells.Cell>());

                                variableCollection[keyWord].Add(cell);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }

            return variableCollection;
        }

        public static void PutValue(Cell cell, string value, Type type)
        {
            if (Type.GetType("System.Int32") == type)
            {
                int bValue;
                if (int.TryParse(value, out bValue))
                    cell.PutValue(bValue);
                else
                    cell.PutValue(value);

                return;
            }

            if ((DateTime.Today).GetType() == type)
            {
                DateTime bValue;
                if (DateTime.TryParse(value, out bValue))
                    cell.PutValue(bValue);
                else
                    cell.PutValue(value);

                return;
            }

            if (Type.GetType("System.Double") == type)
            {
                double bValue;
                if (double.TryParse(value, out bValue))
                    cell.PutValue(bValue);
                else
                    cell.PutValue(value);

                return;
            }

            if (true.GetType() == type)
            {
                bool bValue;
                if (bool.TryParse(value, out bValue))
                    cell.PutValue(bValue);
                else
                    cell.PutValue(value);

                return;
            }

            cell.PutValue(value);
        }

        private static void SetCellStyle(Aspose.Cells.Worksheet instanceSheet, Aspose.Cells.Cell cell, CellStyle cell_style)
        {
            //  粗體
            if (cell_style.Bold.HasValue)
                instanceSheet.Cells[cell.Row, cell.Column].Style.Font.IsBold = cell_style.Bold.Value;
            //  底線
            if (cell_style.Underline.HasValue)
            {
                if (cell_style.Underline.Value)
                    instanceSheet.Cells[cell.Row, cell.Column].Style.Font.Underline = FontUnderlineType.Single;
                else
                    instanceSheet.Cells[cell.Row, cell.Column].Style.Font.Underline = FontUnderlineType.None;
            }
            //  字體名稱
            if (!string.IsNullOrEmpty(cell_style.FontName))
                instanceSheet.Cells[cell.Row, cell.Column].Style.Font.Name = cell_style.FontName;
            //  字體大小
            if (cell_style.FontSize.HasValue)
                instanceSheet.Cells[cell.Row, cell.Column].Style.Font.Size = cell_style.FontSize.Value;
            //  水平位置
            if (cell_style.HAlignment.HasValue)
                instanceSheet.Cells[cell.Row, cell.Column].Style.HorizontalAlignment = cell_style.HAlignment.Value;
            //  垂直位置
            if (cell_style.VAlignment.HasValue)
                instanceSheet.Cells[cell.Row, cell.Column].Style.VerticalAlignment = cell_style.VAlignment.Value;
            //  列高
            if (cell_style.RowHeight.HasValue)
                instanceSheet.Cells.SetRowHeight(cell.Row, cell_style.RowHeight.Value);
            //  合併儲存格
            if (cell_style.MergeObject != null)
            {
                instanceSheet.Cells.Merge(cell.Row, cell.Column, cell_style.MergeObject.row_length, cell_style.MergeObject.column_length);
            }
            //  自動調整列高
            if (cell_style.AutoFitRow.HasValue && cell_style.AutoFitRow.Value)
            {
                if (cell_style.MergeObject != null)
                {
                    Range merged_Range = cell.GetMergedRange();
                    if (merged_Range != null)
                    {
                        double column_width = 0.0f;
                        string content = string.Empty;
                        for (int c = 0; c < merged_Range.ColumnCount; c++)
                        {
                            column_width += merged_Range.Worksheet.Cells.GetColumnWidth(merged_Range.FirstColumn + c);
                        }
                        content = (cell.Value + "").Replace(" ", "_");
                        double row_height_after = SandBox.Instance.GetFitedRowHeight(content, column_width) * 10 / 7.5;
                        instanceSheet.Cells.SetRowHeight(cell.Row, row_height_after > 409 ? 409 : row_height_after);

                        Style style = cell.Style;
                        style.IsTextWrapped = true;
                        StyleFlag sf = new StyleFlag();
                        sf.All = true;
                        merged_Range.ApplyStyle(style, sf);
                    }
                }
                else
                {
                    instanceSheet.Cells[cell.Row, cell.Column].Style.IsTextWrapped = true;
                    double column_width = instanceSheet.Cells.GetColumnWidth(instanceSheet.Cells[cell.Row, cell.Column].Column);
                    string content = (instanceSheet.Cells[cell.Row, cell.Column].Value + "").Replace(" ", "_");
                    double row_height_after = SandBox.Instance.GetFitedRowHeight(content, column_width) * 10 / 7.5;
                    instanceSheet.Cells.SetRowHeight(cell.Row, row_height_after > 409 ? 409 : row_height_after);
                }
            }
            //  背景色
            if (cell_style.BackGroundColor.HasValue)
            {
                Cells celice = instanceSheet.Cells;
                Style celicaStil = null;

                celicaStil = celice[cell.Row, cell.Column].Style;
                celicaStil.ForegroundColor = cell_style.BackGroundColor.Value;
                celicaStil.Pattern = BackgroundType.Solid;
                celice[cell.Row, cell.Column].Style = celicaStil;
            }
            //還原 Formatting Selected Characters in a Cell
            //if (newValue.Length > 0)
            //{
            //    int x = ("[[" + dataTable.TableName + "]]").Length;
            //    int y = dataTable.Rows[0][0].ToString().Length;
            //    int z = 0;
            //    for (int k = 0; k < ((cell.Value + "")).Length; k++)
            //    {
            //        if (newValue.IndexOf(dataTable.Rows[0][0].ToString(), k) == 0)
            //        {
            //            for(int j=k; j<(k+y); j++)
            //                RestoreSelectedCharactersFormatting(cell.Characters(j, 1).Font, dicCharactersFormatting[vIndexs[z]]);

            //            k = k+y;
            //            z ++;
            //        }
            //        //else
            //        //    RestoreSelectedCharactersFormatting(cell.Characters(k, 1).Font, dicCharactersFormatting[k + z * x]);
            //    }
            //}

        }

        /// <summary>
        /// 以實值取代變數
        /// </summary>
        /// <param name="dataSet">資料來源</param>
        /// <param name="instanceSheet">報表</param>
        /// <param name="dataIndex">報表列</param>
        public static void GenerateSheet(DataSet dataSet, Aspose.Cells.Worksheet instanceSheet, int dataIndex, Dictionary<CellObject, CellStyle> dicCellStyles)
        {
            //  所有變數
            Dictionary<string, List<Aspose.Cells.Cell>> variableCollection = DocumentHelper.DiscoverVariableCells(instanceSheet, dataIndex);

            // 資料來源中的所有資料
            foreach (KeyValuePair<string, List<Aspose.Cells.Cell>> kv in variableCollection)
            {
                if (!dataSet.Tables.Contains(kv.Key))
                    continue;

                DataTable dataTable = dataSet.Tables[kv.Key];
                // 取得包含變數的儲存格
                foreach (Aspose.Cells.Cell cell in variableCollection[kv.Key])                
                {
                    int n = 0;
                    int m = 0;

                    //Dictionary<int, dynamic> dicCharactersFormatting = BackupSelectedCharactersFormatting(cell);
                    Aspose.Cells.Range mergedRange = null;
                    // 若變數內容為影像，如：畢業照
                    if (dataTable.Columns[0].DataType == Type.GetType("System.Byte[]"))
                    {
                        Aspose.Cells.Range imgRange = instanceSheet.Cells[cell.Row, cell.Column].GetMergedRange();
                        Byte[] imageData = dataTable.Rows[0][0] as Byte[];
                        MemoryStream ms = new MemoryStream(imageData);
                        int pictureIndex = 0;

                        if (ms.Length == 0)
                            continue;

                        if (imgRange != null)
                            pictureIndex = instanceSheet.Pictures.Add(imgRange.FirstRow, imgRange.FirstColumn, (imgRange.FirstRow + imgRange.RowCount), (imgRange.FirstColumn + imgRange.ColumnCount), ms);
                        else
                            pictureIndex = instanceSheet.Pictures.Add(cell.Row, cell.Column, cell.Row + 1, cell.Column + 1, ms);

                        instanceSheet.Pictures[pictureIndex].UpperDeltaX = 10;
                        instanceSheet.Pictures[pictureIndex].UpperDeltaY = 10;
                        instanceSheet.Pictures[pictureIndex].LowerDeltaX = -10;
                        instanceSheet.Pictures[pictureIndex].LowerDeltaY = -10;

                        cell.PutValue("");

                        continue;
                    }
                    //若變數內容為單一值，如：學號、姓名
                    if (dataTable.Rows.Count == 1 && dataTable.Columns.Count == 1)
                    {
                        string oldValue = cell.Value + "";
                        //List<int> vIndexs = new List<int>();
                        //if (oldValue.Length >0)
                        //{
                        //    int i=0;
                        //    do
                        //    {
                        //        i = oldValue.IndexOf("[[" + dataTable.TableName + "]]", i);
                        //        if (i >= 0)
                        //        {
                        //            vIndexs.Add(i);
                        //            i++;
                        //        }
                        //    }
                        //    while (oldValue.IndexOf("[[" + dataTable.TableName + "]]", i) >= 0);
                        //}
                        //  以值取代變數
                        string newValue = oldValue.Replace("[[" + dataTable.TableName + "]]", dataTable.Rows[0][0].ToString());
                        //if (newValue == dataTable.Rows[0][0].ToString())
                        //{
                        //    DateTime date_time_data;
                        //    decimal decimal_data;

                        //    if (DateTime.TryParse(newValue, out date_time_data))
                        //        cell.PutValue(date_time_data.ToString(), true);
                        //    else if (decimal.TryParse(newValue, out decimal_data))
                        //        cell.PutValue(decimal_data + "", true);
                        //    else
                        //        cell.PutValue(newValue);
                        //}
                        if (!string.IsNullOrEmpty(newValue))
                        {
                            PutValue(cell, newValue, dataTable.Columns[0].DataType);
                            //cell.PutValue(newValue);
                            //if ((cell.Value + "").Length != newValue.Length || (!(cell.Value + "").Equals(newValue)))
                            //    cell.PutValue(newValue);
                        }

                        CellStyle cell_style = null;
                        CellObject cell_object = new CellObject(0, 0, dataTable.TableName, dataTable.DataSet.DataSetName, instanceSheet.Name);

                        if (dicCellStyles != null)
                        {
                            foreach (CellObject c in dicCellStyles.Keys)
                            {
                                if (c.RowIndex == cell_object.RowIndex && c.ColumnIndex == cell_object.ColumnIndex && c.TableName == cell_object.TableName && c.DataSetName == cell_object.DataSetName && c.WorkSheetName == cell_object.WorkSheetName)
                                    cell_style = dicCellStyles[c];
                            }
                        }
                        if (cell_style != null)
                        {
                            SetCellStyle(instanceSheet, cell, cell_style);
                            continue;
                        }
                    }
                    // 若變數內容為矩陣結構，如：異動資料、學期成績資料
                    if (dataTable.Rows.Count > 1 || dataTable.Columns.Count > 1)
                    {
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            m = 0;
                            for (int j = 0; j < dataTable.Columns.Count; j++)
                            {
                                CellStyle cell_style = null;
                                CellObject cell_object = new CellObject(i, j, dataTable.TableName, dataTable.DataSet.DataSetName, instanceSheet.Name);
                                if (dicCellStyles != null)
                                {
                                    foreach (CellObject c in dicCellStyles.Keys)
                                    {
                                        if (c.RowIndex==cell_object.RowIndex && c.ColumnIndex == cell_object.ColumnIndex && c.TableName == cell_object.TableName && c.DataSetName == cell_object.DataSetName && c.WorkSheetName == cell_object.WorkSheetName)
                                            cell_style = dicCellStyles[c];
                                    }
                                }
                                // 以值取代變數
                                string oldValue = (instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Value == null) ? string.Empty : instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Value.ToString();
                                string newValue = string.Empty;

                                if (oldValue.Contains("[[" + kv.Key + "]]"))
                                    newValue = oldValue.Replace("[[" + kv.Key + "]]", dataTable.Rows[i][j].ToString());
                                else if (oldValue.Contains("[[") && oldValue.Contains("]]"))
                                    newValue = dataTable.Rows[i][j].ToString() + oldValue;
                                else
                                    newValue = oldValue + dataTable.Rows[i][j].ToString();

                                if (!string.IsNullOrEmpty(newValue))
                                {
                                    PutValue(instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m], newValue, dataTable.Columns[j].DataType);
                                    //instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].PutValue(newValue);
                                    //if ((instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Value + "").Length != newValue.Length || (!(instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Value + "").Equals(newValue)))
                                    //    instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].PutValue(newValue);
                                }
                                if (cell_style != null)
                                {
                                    SetCellStyle(instanceSheet, instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m], cell_style);
                                    ////  粗體
                                    //instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Style.Font.IsBold = cell_style.Bold;
                                    ////  底線
                                    //if (cell_style.Underline)
                                    //    instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Style.Font.Underline = FontUnderlineType.Single;
                                    ////  字體名稱
                                    //if (!string.IsNullOrEmpty(cell_style.FontName))
                                    //    instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Style.Font.Name = cell_style.FontName;
                                    ////  字體大小
                                    //if (cell_style.FontSize > 0)
                                    //    instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Style.Font.Size = cell_style.FontSize;
                                    ////  水平位置
                                    //if (cell_style.HAlignment != null)
                                    //    instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Style.HorizontalAlignment = cell_style.HAlignment.Value;
                                    ////  垂直位置
                                    //if (cell_style.VAlignment != null)
                                    //    instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Style.VerticalAlignment = cell_style.VAlignment.Value;
                                    ////  列高
                                    //if (cell_style.RowHeight != null)
                                    //    instanceSheet.Cells.SetRowHeight(cell.Row + i + n, cell_style.RowHeight.Value);
                                    ////  合併儲存格
                                    //if (cell_style.MergeObject != null)
                                    //{
                                    //    instanceSheet.Cells.Merge(cell.Row + i + n, cell.Column + j + m, cell_style.MergeObject.row_length, cell_style.MergeObject.column_length);
                                    //}
                                    //  自動調整列高
                                    //if (cell_style.AutoFitRow)
                                    //{
                                    //    if (cell_style.MergeObject != null)
                                    //    {
                                    //        Range merged_Range = instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].GetMergedRange();
                                    //        if (merged_Range != null)
                                    //        {
                                    //            double column_width = 0.0f;
                                    //            for (int c = 0; c < merged_Range.ColumnCount; c++)
                                    //            {
                                    //                column_width += merged_Range.Worksheet.Cells.GetColumnWidth(merged_Range.FirstColumn + c);
                                    //            }
                                                
                                    //            instanceSheet.Cells.SetRowHeight(cell.Row + i + n, SandBox.Instance.GetFitedRowHeight(instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Value + "", column_width) * 10 / 7.5);

                                    //            //AutoFitRowByNPOI();

                                    //            Style style = instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Style;
                                    //            style.IsTextWrapped = true;
                                    //            StyleFlag sf = new StyleFlag();
                                    //            sf.All = true;
                                    //            merged_Range.ApplyStyle(style, sf);
                                    //        }
                                    //    }
                                    //    else
                                    //    {
                                    //        //instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Style.IsTextWrapped = true;
                                    //        //instanceSheet.AutoFitRow(cell.Row + i + n, 0, instanceSheet.Cells.MaxDataColumn);
                                    //        //double row_height_after = instanceSheet.Cells.GetRowHeight(cell.Row + i + n);
                                    //        //instanceSheet.Cells.SetRowHeight(cell.Row + i + n, row_height_after * 10 / 9.0);

                                    //        instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Style.IsTextWrapped = true;
                                    //        double column_width = instanceSheet.Cells.GetColumnWidth(instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Column);
                                    //        string content = (instanceSheet.Cells[cell.Row + i + n, cell.Column + j + m].Value + "").Replace(" ", "_");
                                    //        double row_height_after = SandBox.Instance.GetFitedRowHeight(content, column_width);
                                    //        instanceSheet.Cells.SetRowHeight(cell.Row + i + n, row_height_after * 10 / 7.5);
                                    //    }
                                    //}
                                    ////  背景色
                                    //if (cell_style.BackGroundColor != null)
                                    //{
                                    //    Cells celice = instanceSheet.Cells;
                                    //    Style celicaStil = null;

                                    //    celicaStil = celice[cell.Row + i + n, cell.Column + j + m].Style;
                                    //    celicaStil.ForegroundColor = cell_style.BackGroundColor.Value;                                         
                                    //    celicaStil.Pattern = BackgroundType.Solid;
                                    //    celice[cell.Row + i + n, cell.Column + j + m].Style = celicaStil;
                                    //}
                                }

                                // 若變數所在儲存格為合併儲存格，則 ColumnIndex 要增加
                                mergedRange = instanceSheet.Cells[cell.Row, cell.Column + j + m].GetMergedRange();
                                if (mergedRange != null)
                                    // 合併儲存格 ColumnIndex 的增加量
                                    m += mergedRange.ColumnCount - 1;
                            }
                            // 若變數所在儲存格為合併儲存格，則 RowIndex 要增加
                            mergedRange = instanceSheet.Cells[cell.Row + i + n, cell.Column].GetMergedRange();
                            if (mergedRange != null)
                                // 合併儲存格 RowIndex 的增加量
                                n += mergedRange.RowCount - 1;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 開啟檔案
        /// </summary>
        /// <param name="fileFullName">檔案名稱=路徑+檔名(含副檔名)。</param>
        public static void Open(string fileFullName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileFullName);
            }
            catch (Exception)
            {
                throw new Exception("開啟檔案發生錯誤，您可能沒有相關的應用程式可以開啟此類型檔案，或權限不足！");
            }
        }

        /// <summary>
        /// 開啟檔案
        /// </summary>
        /// <param name="path"></param>
        /// <param name="filename">檔案名稱=檔名(含副檔名)，程式會將路徑與檔名結合。</param>
        public static void Open(string path, string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(Path.Combine(path, filename));
            }
            catch (Exception)
            {
                throw new Exception("開啟檔案發生錯誤，您可能沒有相關的應用程式可以開啟此類型檔案，或權限不足！");
            }
        }
    }
}
