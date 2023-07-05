using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using Testauto.Ultils;
using Testauto.Log;

namespace Testauto.Log
{
    public class LoginData : TestData, ILog<LoginData>
    {
        public string Username { get; set; }
        public string Password { get; set; }

        public void WriteLog(string src, string sheetName, HashSet<LoginData> logs)
        {
            var workbook = ExcelUltils.GetWorkbook(src);
            var sheet = ExcelUltils.GetSheet(workbook, sheetName);

            int startRow = 0;
            int lastRow = sheet.RowCount();
            if (lastRow < startRow)
                lastRow = startRow;

            var rowStyle = ExcelUltils.GetRowStyle(workbook);

            foreach (var log in logs)
            {
                var row = sheet.Row(lastRow + 1); // Create a new row using Row() method

                row.Height = 60;
                row.Style = rowStyle;
                log.WriteDataRow(log, row, sheet);
                lastRow++;
            }

            ExcelUltils.Export(src, workbook);
        }

        public void WriteDataRow(LoginData log, IXLRow row, IXLWorksheet sheet)
        {
            var globalStyle = row.Style;
            IXLCell cell;

            cell = row.Cell(1); // Cell index 0
            cell.Value = log.Username;
            cell.Style = globalStyle;

            cell = row.Cell(2);
            cell.Value = log.Password;
            cell.Style = globalStyle;

            WriteTestData(3, row, sheet); // Assuming WriteTestData starts from index 3
        }
    }
}
