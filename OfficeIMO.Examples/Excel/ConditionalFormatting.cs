using System;
using OfficeIMO.Excel;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Excel {
    public static class ConditionalFormatting {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Conditional formatting");
            string filePath = System.IO.Path.Combine(folderPath, "Conditional Formatting.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 10d);
                sheet.CellValue(2, 1, 20d);
                sheet.CellValue(3, 1, 30d);
                sheet.AddConditionalColorScale("A1:A3", Color.Red, Color.Green);
                document.Save(openExcel);
            }
        }
    }
}
