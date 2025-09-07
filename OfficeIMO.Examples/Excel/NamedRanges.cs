using System;
using System.IO;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    internal static class NamedRanges {
        public static void Example() {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "NamedRanges.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("My Sheet");
                sheet.CreateNamedRange("LocalRange", "A1:B2");
                document.CreateNamedRange("GlobalRange", sheet, "C1:D2", workbookScope: true);

                Console.WriteLine(sheet.GetNamedRange("LocalRange"));
                Console.WriteLine(document.GetNamedRange("GlobalRange"));
            }

            File.Delete(filePath);
        }
    }
}
