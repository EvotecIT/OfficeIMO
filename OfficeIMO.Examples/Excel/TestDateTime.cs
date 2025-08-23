using System;
using System.IO;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    internal static class TestDateTime {
        public static void Example_TestDateTime(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Testing DateTime formatting");
            string filePath = Path.Combine(folderPath, "TestDateTime.xlsx");
            
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("DateTest");
                
                // Test different ways of setting DateTime
                DateTime testDate = new DateTime(2024, 12, 25, 14, 30, 0);
                
                // Direct DateTime call
                sheet.CellValue(1, 1, "Direct DateTime:");
                sheet.CellValue(1, 2, testDate);
                
                // Through object
                sheet.CellValue(2, 1, "Through object:");
                object objDate = testDate;
                sheet.CellValue(2, 2, objDate);
                
                // Through Cell method
                sheet.CellValue(3, 1, "Through Cell:");
                sheet.Cell(3, 2, testDate);
                
                // With explicit format
                sheet.CellValue(4, 1, "With format:");
                sheet.Cell(4, 2, testDate, null, "yyyy-MM-dd");
                
                document.Save();
                
                // Check validation
                var errors = document.DocumentValidationErrors;
                Console.WriteLine($"Validation errors: {errors.Count}");
                
                if (openExcel) {
                    document.Open();
                }
            }
        }
    }
}