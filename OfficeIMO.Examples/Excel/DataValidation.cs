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
                ExcelSheet sheet = document.AddSheet("Data");

                sheet.Range("A1:A10").Validation().WholeNumber(ExcelDataValidationOperator.Between, 1, 10);
                sheet.Range("B1:B10").Validation().Decimal(ExcelDataValidationOperator.GreaterThan, 5.5);
                sheet.Range("C1:C10").Validation().Date(ExcelDataValidationOperator.LessThan, new DateTime(2024, 1, 1));
                sheet.Range("D1:D10").Validation().Time(ExcelDataValidationOperator.Equal, TimeSpan.FromHours(12));
                sheet.Range("E1:E10").Validation().TextLength(ExcelDataValidationOperator.LessThanOrEqual, 20);
                sheet.Range("F1:F10").Validation().CustomFormula("SUM(A1:B1)>10");

                document.Save();
                if (openExcel) document.OpenInApplication();
            }
        }
    }
}

