using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchema
{
    public class Excel
    {
        public string FilePath { get; private set; }
        public string SheetName { get; private set; }
        public ExcelPackage workbook { get; private set; }
        public ExcelWorksheet sheet { get; private set; }

        public Excel(string filePath, string sheetName)
        {
            FilePath = filePath;
            SheetName = sheetName;
            ini();
        }

        public ExcelWorksheet SetSheet(string sheetName)
        {
            sheet = workbook.Workbook.Worksheets[sheetName]; // 可以使用頁籤名稱
            return sheet;
        }
        public void ini()
        {
            workbook = new ExcelPackage(new FileInfo(FilePath));
            sheet = workbook.Workbook.Worksheets[SheetName];
        }
        public void Dispose()
        {
            workbook.Dispose();
            sheet.Dispose();
        }
    }
}
