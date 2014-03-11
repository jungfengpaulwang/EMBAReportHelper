using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReportHelper
{
    public class MergeObject
    {
        public int row_length { private set; get; }
        public int column_length { private set; get; }

        public MergeObject(int RowLength, int ColumnLength) 
        {
            this.row_length = RowLength;
            this.column_length = ColumnLength;
        }
    }

    public class CellStyle
    {
        private bool? bold;
        private bool? underline;
        private string font_name;
        private int? font_size;
        private bool? auto_fit_row;
        private MergeObject merge_object;

        public MergeObject MergeObject
        {
            private set { this.merge_object = value; }
            get { return this.merge_object; }
        }

        public bool Bold 
        {
            private set { this.bold = value; } 
            get { return (this.bold.HasValue) ? this.bold.Value : false; }
        }

        public bool Underline
        {
            private set { this.underline = value; }
            get { return (this.underline.HasValue) ? this.underline.Value : false; }
        }

        public bool AutoFitRow
        {
            private set { this.auto_fit_row = value; }
            get { return (this.auto_fit_row.HasValue) ? this.auto_fit_row.Value : false; }
        }

        public string FontName
        {
            private set { this.font_name = value; }
            get 
            {
                if (!string.IsNullOrWhiteSpace(this.font_name))
                    return this.font_name;

                return "新細明體"; 
            }
        }

        public int FontSize
        {
            private set { this.font_size = value; }
            get { return (this.font_size.HasValue) ? this.font_size.Value : 0; }
        }

        public enum HorizontalAlignment { Left, Center, Right };
        public enum VerticalAlignment { Top, Center, Bottom };

        public Aspose.Cells.TextAlignmentType? HAlignment { private set; get; }
        public Aspose.Cells.TextAlignmentType? VAlignment { private set; get; }
        public double? RowHeight { private set; get; }
        public System.Drawing.Color? BackGroundColor { private set; get; }

        public CellStyle() { }

        public CellStyle SetFontBold(bool Bold)
        {
            this.Bold = Bold;
            return this;
        }

        public CellStyle SetFontUnderline(bool Underline)
        {
            this.Underline = Underline;
            return this;
        }

        public CellStyle SetAutoFitRow(bool AutoFitRow)
        {
            this.AutoFitRow = AutoFitRow;
            return this;
        }

        public CellStyle SetFontName(string FontName)
        {
            this.FontName = FontName;
            return this;
        }

        public CellStyle SetFontSize(int FontSize)
        {
            this.FontSize = FontSize;
            return this;
        }

        public CellStyle SetFontBackGroundColor(System.Drawing.Color Color)
        {
            this.BackGroundColor = Color;
            return this;
        }

        public CellStyle SetRowHeight(double? RowHeight)
        {
            this.RowHeight = RowHeight;
            return this;
        }

        public CellStyle Merge(int RowLength, int ColumnLength)
        {
            this.merge_object = new MergeObject(RowLength, ColumnLength);
            return this;
        }

        public CellStyle SetFontHorizontalAlignment(HorizontalAlignment Alignment)
        {
            switch (Alignment)
            {
                case HorizontalAlignment.Center:
                    this.HAlignment = Aspose.Cells.TextAlignmentType.Center;
                    break;
                case HorizontalAlignment.Left:
                    this.HAlignment = Aspose.Cells.TextAlignmentType.Left;
                    break;
                case HorizontalAlignment.Right:
                    this.HAlignment = Aspose.Cells.TextAlignmentType.Right;
                    break;
                default:
                    this.HAlignment = null;
                    break;
            }
            return this;
        }

        public CellStyle SetFontVerticalAlignment(VerticalAlignment Alignment)
        {
            switch (Alignment)
            {
                case VerticalAlignment.Center:
                    this.VAlignment = Aspose.Cells.TextAlignmentType.Center;
                    break;
                case VerticalAlignment.Top:
                    this.VAlignment = Aspose.Cells.TextAlignmentType.Top;
                    break;
                case VerticalAlignment.Bottom:
                    this.VAlignment = Aspose.Cells.TextAlignmentType.Bottom;
                    break;
                default:
                    this.VAlignment = null;
                    break;
            }
            return this;
        }    
    }
}
