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
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, "Name");
                sheet.SetCellValue(1, 2, "Value");
                sheet.SetCellValue(2, 1, "A");
                sheet.SetCellValue(2, 2, 10d);
                sheet.AddTable("A1:B2", true, "MyTable", TableStyle.TableStyleMedium9);
                document.Save(openExcel);
            }
        }
    }
}
