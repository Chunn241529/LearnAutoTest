using System;
using ClosedXML.Excel;
using System.IO;
using OpenQA.Selenium;

namespace Testauto.Ultils
{
    public class ExcelUltils
    {
        public static string CHROME_DRIVER_SRC = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
        public static string DATA_SRC = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test-resources", "Data");
        public static string IMG_SRC = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test-resources", "Img");


        public static XLWorkbook GetWorkbook(string filePath)
        {
            var src = new FileInfo(filePath);
            if (!src.Exists)
            {
                throw new IOException("Không tồn tại file với đường dẫn: " + filePath);
            }
            var workbook = new XLWorkbook(src.FullName);
            return workbook;
        }

        public static IXLWorksheet GetSheet(XLWorkbook workbook, string sheetName)
        {
            var sheet = workbook.Worksheet(sheetName);
            if (sheet == null)
            {
                throw new NullReferenceException("Sheet " + sheetName + " không tồn tại, thêm dữ liệu thất bại!");
            }
            return sheet;
        }

        public static IXLStyle GetRowStyle(XLWorkbook workbook)
        {
            var rowStyle = workbook.Style;
            rowStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rowStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            rowStyle.Alignment.WrapText = true;
            return rowStyle;
        }

        public static string GetCellValue(IXLWorksheet sheet, int row, int column)
        {
            var cell = sheet.Cell(row, column);
            if (cell.IsEmpty())
            {
                return "";
            }
            if (cell.DataType == XLDataType.Text)
            {
                return cell.GetString();
            }
            if (cell.DataType == XLDataType.Number)
            {
                return cell.GetFormattedString();
            }
            return "";
        }

        // Chụp ảnh
        public static void TakeScreenshot(IWebDriver driver, String outputSrc)
        {
            Screenshot screenshot = ((ITakesScreenshot)driver).GetScreenshot();

            // Save the screenshot to a file (replace "ScreenshotPath" with the desired path and file name)
            screenshot.SaveAsFile(outputSrc, ScreenshotImageFormat.Png);
        }

        public static object[,] ReadSheetData(IXLWorksheet sheet)
        {
            var range = sheet.RangeUsed();
            int rows = range.RowCount();
            int columns = range.ColumnCount();

            object[,] data = new object[rows - 1, columns];

            for (int row = 2; row <= rows; row++)
            {
                for (int col = 1; col <= columns; col++)
                {
                    data[row - 2, col - 1] = GetCellValue(sheet, row, col);
                }
            }
            return data;
        }

        // Vẽ ảnh
        public static void WriteImage(string image, IXLRow row, IXLCell cell, IXLWorksheet sheet)
        {
            var picture = sheet.Pictures.Add(image);
            picture.MoveTo(sheet.Cell(row.RowNumber(), cell.Address.ColumnNumber));
        }

        public static void Export(string outputSrc, XLWorkbook workbook)
        {
            workbook.SaveAs(outputSrc);
        }
    }
}
