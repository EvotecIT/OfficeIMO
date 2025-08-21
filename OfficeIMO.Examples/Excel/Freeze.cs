using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates freezing top rows and left columns.
    /// </summary>
    public static class Freeze {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Freeze panes");
            string filePath = System.IO.Path.Combine(folderPath, "Freeze.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Header1");
                sheet.CellValue(1, 2, "Header2");
                sheet.CellValue(2, 1, "Value1");
                sheet.CellValue(2, 2, "Value2");
                sheet.Freeze(topRows: 1, leftCols: 1);
                document.Save(openExcel);
            }
        }
    }
}

