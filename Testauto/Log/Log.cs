using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Testauto.Log
{
    public interface ILog<T> where T : TestData
    {
        void WriteLog(string src, string sheetName, HashSet<T> logs);
        void WriteDataRow(T log, IXLRow row, IXLWorksheet sheet);
    }
}
