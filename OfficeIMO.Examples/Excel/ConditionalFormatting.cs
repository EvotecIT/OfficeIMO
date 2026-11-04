using System;
using OfficeIMO.Excel;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Excel {
    public static class ConditionalFormatting {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Conditional formatting");
            string filePath = System.IO.Path.Combine(folderPath, "Conditional Formatting.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue(10d);
                sheet.Cell(2, 1).SetValue(20d);
                sheet.Cell(3, 1).SetValue(30d);
                // Light status tints keep the values readable: a pure red-to-green scale passes through brown.
                sheet.Range("A1:A3").ConditionalFormat().ColorScale(Color.ParseHex("F8D7DA"), Color.ParseHex("C9ECD3"));
                document.Save();
                if (openExcel) document.OpenInApplication();
            }
        }
    }
}
