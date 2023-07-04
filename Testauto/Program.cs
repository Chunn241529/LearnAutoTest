using System;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using OpenQA.Selenium;

namespace Testauto
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var wbook = new XLWorkbook();

            var ws = wbook.Worksheets.Add("Sheet1");

            ws.Cell("A1").Value = "skaaaaaa";
            ws.Cell("A2").Value = "cloud";
            ws.Cell("A3").Value = "book";
            ws.Cell("A4").Value = "cup";
            ws.Cell("A5").Value = "snake";
            ws.Cell("A6").Value = "falcon";
            ws.Cell("B1").Value = "in";
            ws.Cell("B2").Value = "tool";
            ws.Cell("B3").Value = "war";
            ws.Cell("B4").Value = "snow";
            ws.Cell("B5").Value = "tree";
            ws.Cell("B6").Value = "ten";

            var n = ws.Range("A1:C10").CellsUsed().Count();
            Console.WriteLine($"There are {n} words in the range");

            Console.WriteLine("The following words have three latin letters:");

            var words = ws.Range("A1:C10")
                .CellsUsed()
                .Select(c => c.Value.ToString())
                .Where(c => c?.Length == 3)
                .ToList();

            words.ForEach(Console.WriteLine);

            wbook.SaveAs(@"D:\trung\usedcells.xlsx");
        }



      
    }
}
