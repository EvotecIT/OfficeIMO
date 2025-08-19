using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    public static class ConditionalFormatting {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Conditional formatting");
            string filePath = System.IO.Path.Combine(folderPath, "Conditional Formatting.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, 10d);
                sheet.SetCellValue(2, 1, 20d);
                sheet.SetCellValue(3, 1, 30d);
                sheet.AddConditionalColorScale("A1:A3", "FFFF0000", "FF00FF00");
                document.Save(openExcel);
            }
        }
    }
}
