using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates autofitting columns and rows when writing cell values.
    /// </summary>
    public static class AutoFit {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - AutoFit columns and rows");
            string filePath = System.IO.Path.Combine(folderPath, "AutoFit.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, "This is a very long piece of text");
                sheet.SetCellValue(1, 2, "Short");
                sheet.SetCellValue(2, 1, "Second line\nwith newline");
                sheet.SetCellValue(3, 1, "Line1\nLine2\nLine3");
                sheet.SetCellValue(4, 1, "Temporary");
                sheet.SetCellValue(4, 1, string.Empty);

                // Autofit a single column and a single row
                sheet.AutoFitColumn(1);
                sheet.AutoFitRow(3);

                // Autofit the entire worksheet
                sheet.AutoFitColumns();
                sheet.AutoFitRows();

                document.Save(openExcel);
            }
        }
    }
}
