using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates creating a table with style.
    /// </summary>
    public static class AddTable {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Add table");
            string filePath = System.IO.Path.Combine(folderPath, "Add Table.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Name");
                sheet.Cell(1, 2).SetValue("Value");
                sheet.Cell(2, 1).SetValue("A");
                sheet.Cell(2, 2).SetValue(10d);
                sheet.AddTable("A1:B2", true, "MyTable", ExcelTableStyle.TableStyleMedium9);
                document.Save();
                if (openExcel) document.OpenInApplication();
            }
        }
    }
}
