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
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Header1");
                sheet.Cell(1, 2).SetValue("Header2");
                sheet.Cell(2, 1).SetValue("Value1");
                sheet.Cell(2, 2).SetValue("Value2");
                sheet.View.Freeze(topRows: 1, leftColumns: 1);
                document.Save();
                if (openExcel) document.OpenInApplication();
            }
        }
    }
}

