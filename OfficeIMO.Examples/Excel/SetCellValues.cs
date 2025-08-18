using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates setting various cell values.
    /// </summary>
    public static class SetCellValues {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Setting cell values");
            string filePath = System.IO.Path.Combine(folderPath, "Set Cell Values.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                sheet.SetCellValue(1, 1, "Name");
                sheet.SetCellValue(1, 2, "Value");
                sheet.SetCellValue(2, 1, "Pi");
                sheet.SetCellValue(2, 2, Math.PI);
                sheet.SetCellValue(3, 1, "Today");
                sheet.SetCellValue(3, 2, DateTime.Today);
                sheet.SetCellValue(4, 1, "IsCool");
                sheet.SetCellValue(4, 2, true);
                sheet.SetCellFormula(5, 2, "SUM(B2)");

                document.Save(openExcel);
            }
        }
    }
}
