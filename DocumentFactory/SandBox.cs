using System;
using System.IO;

public class SandBox
{
    private static readonly Lazy<SandBox> LazyInstance = new Lazy<SandBox>(() => new SandBox());
    public static SandBox Instance { get { return LazyInstance.Value; } }
    private  Aspose.Cells.Workbook workbook;

    private SandBox() 
    {
        this.workbook = new Aspose.Cells.Workbook();
        this.workbook.Worksheets.Add();
    }

    public double GetFitedRowHeight(string content, double width)
    {
        SandBox.Instance.workbook.Worksheets[0].Cells[0, 0].PutValue(content);
        SandBox.Instance.workbook.Worksheets[0].Cells.SetColumnWidth(0, width);
        SandBox.Instance.workbook.Worksheets[0].Cells[0, 0].Style.IsTextWrapped = true;
        SandBox.Instance.workbook.Worksheets[0].AutoFitRow(0);

        return SandBox.Instance.workbook.Worksheets[0].Cells.GetRowHeight(0);
    }
}