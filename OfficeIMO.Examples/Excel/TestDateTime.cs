using System;
using System.IO;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    internal static class TestDateTime {
        public static void Example_TestDateTime(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Testing DateTime formatting");
            string filePath = Path.Combine(folderPath, "TestDateTime.xlsx");
            
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddSheet("DateTest");
                
                // Test different ways of setting DateTime
                DateTime testDate = new DateTime(2024, 12, 25, 14, 30, 0);
                
                // Direct DateTime call
                sheet.Cell(1, 1).SetValue("Direct DateTime:");
                sheet.Cell(1, 2).SetValue(testDate);
                
                // Through object
                sheet.Cell(2, 1).SetValue("Through object:");
                object objDate = testDate;
                sheet.Cell(2, 2).SetValue(objDate);
                
                // Through Cell method
                sheet.Cell(3, 1).SetValue("Through Cell:");
                sheet.Cell(3, 2).SetValue(testDate);
                
                // With explicit format
                sheet.Cell(4, 1).SetValue("With format:");
                sheet.Cell(4, 2).SetValue(testDate).SetNumberFormat("yyyy-MM-dd");
                
                document.Save();
                
                // Check validation
                var errors = document.DocumentValidationErrors;
                Console.WriteLine($"Validation errors: {errors.Count}");
                
                if (openExcel) {
                    document.OpenInApplication();
                }
            }
        }
    }
}