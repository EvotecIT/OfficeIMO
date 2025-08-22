using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates validating a spreadsheet document.
    /// </summary>
    public static class ValidateDocument {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Validate document");
            string filePath = System.IO.Path.Combine(folderPath, "ValidateDocument.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Test");

                Console.WriteLine(document.DocumentIsValid);
                Console.WriteLine(document.DocumentValidationErrors);
                document.Save(openExcel);
            }
        }
    }
}

