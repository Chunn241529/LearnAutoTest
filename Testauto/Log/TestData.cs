using System;
using ClosedXML.Excel;
using Testauto.Ultils;

namespace Testauto.Log
{
    public class TestData
    {
        public string Action { get; set; }
        public DateTime LogTime { get; set; }
        public string TestMethod { get; set; }
        public string Expected { get; set; }
        public string Actual { get; set; }
        public string Status { get; set; }
        public string Exception { get; set; }
        public string Image { get; set; }

        public void WriteTestData(int index, IXLRow row, IXLWorksheet sheet)
        {
            var globalStyle = row.Style;
            IXLCell cell; // Tạo đối tượng Cell để ghi Data

            cell = row.Cell(index); // Vị trí Column bắt đầu
            cell.Value = Action;
            cell.Style = globalStyle;

            cell = row.Cell(index + 1);
            cell.Value = LogTime;
            var datetimeStyle = globalStyle;
            datetimeStyle.DateFormat.Format = "hh:mm:ss dd-MM-yyyy";
            cell.Style = datetimeStyle;

            cell = row.Cell(index + 2);
            cell.Value = TestMethod;
            cell.Style = globalStyle;

            cell = row.Cell(index + 3);
            cell.Value = Expected;
            cell.Style = globalStyle;

            cell = row.Cell(index + 4);
            cell.Value = Actual;
            cell.Style = globalStyle;

            cell = row.Cell(index + 5);
            cell.Value = Status;
            cell.Style = globalStyle;

            if (Exception != null)
            {
                cell = row.Cell(index + 6);
                cell.Value = Exception;
                cell.Style = globalStyle;
            }
            if (Image != null)
            {
                cell = row.Cell(index + 7);
                cell.Style = globalStyle;
                ExcelUltils.WriteImage(Image, row, cell, sheet);

                cell = row.Cell(index + 8);
                cell.Value = "Link Screenshot";
                cell.Style = globalStyle;

                cell.Style.Font.FontColor = XLColor.Blue;
                cell.Style.Font.Underline = XLFontUnderlineValues.Single;

                var hyperlinkFormula = $"HYPERLINK(\"{Image.Replace("\\", "/")}\", \"Link Screenshot\")";
                cell.FormulaA1 = hyperlinkFormula;
            }
        }
    }
}
