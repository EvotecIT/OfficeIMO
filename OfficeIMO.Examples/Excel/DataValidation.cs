using System;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates various data validations on a worksheet.
    /// </summary>
    public static class DataValidation {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Data validation example");
            string filePath = Path.Combine(folderPath, "DataValidation.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");

                sheet.ValidationWholeNumber("A1:A10", DataValidationOperatorValues.Between, 1, 10);
                sheet.ValidationDecimal("B1:B10", DataValidationOperatorValues.GreaterThan, 5.5);
                sheet.ValidationDate("C1:C10", DataValidationOperatorValues.LessThan, new DateTime(2024, 1, 1));
                sheet.ValidationTime("D1:D10", DataValidationOperatorValues.Equal, TimeSpan.FromHours(12));
                sheet.ValidationTextLength("E1:E10", DataValidationOperatorValues.LessThanOrEqual, 20);
                sheet.ValidationCustomFormula("F1:F10", "SUM(A1:B1)>10");

                document.Save(openExcel);
            }
        }
    }
}

