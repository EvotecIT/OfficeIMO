using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates adding a table when some cells in the range are missing.
    /// </summary>
    public static class AddTableMissingCells {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Add table with missing cells");
            string filePath = System.IO.Path.Combine(folderPath, "Add Table Missing Cells.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                // intentionally leave row 2 cells empty
                sheet.AddTable("A1:B2", true, "MyTable", TableStyle.TableStyleMedium9);
                document.Save(openExcel);
            }
        }
    }
}
