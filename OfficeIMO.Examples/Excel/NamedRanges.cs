using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    public static class NamedRanges {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Named ranges");
            string filePath = System.IO.Path.Combine(folderPath, "NamedRanges.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Value1");
                sheet.CellValue(2, 1, "Value2");

                document.SetNamedRange("GlobalRange", "'Data'!A1:A2", save: false);
                sheet.SetNamedRange("LocalRange", "A1", save: false);

                foreach (var pair in document.GetAllNamedRanges()) {
                    Console.WriteLine($"Workbook named range {pair.Key}: {pair.Value}");
                }
                foreach (var pair in sheet.GetAllNamedRanges()) {
                    Console.WriteLine($"Worksheet named range {pair.Key}: {pair.Value}");
                }

                document.Save(openExcel);
            }
        }
    }
}

