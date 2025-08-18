using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates setting various cell values.
    /// </summary>
    public static class SetCellValues {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Setting cell values");
            string filePath = System.IO.Path.Combine(folderPath, "Set Cell Values.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                sheet.SetCellValue(1, 1, "Name");
                sheet.SetCellValue(1, 2, "Value");
                sheet.SetCellValue(2, 1, "Pi");
                sheet.SetCellValue(2, 2, Math.PI);
                sheet.SetCellValue(3, 1, "Today");
                sheet.SetCellValue(3, 2, DateTime.Today);
                sheet.SetCellFormat(3, 2, "yyyy-mm-dd");
                sheet.SetCellValue(4, 1, "IsCool");
                sheet.SetCellValue(4, 2, true);
                sheet.SetCellValue(5, 1, "Offset");
                sheet.SetCellValue(5, 2, DateTimeOffset.Now);
                sheet.SetCellFormat(5, 2, "yyyy-mm-dd hh:mm");
                sheet.SetCellValue(6, 1, "Duration");
                sheet.SetCellValue(6, 2, TimeSpan.FromHours(1.5));
                sheet.SetCellFormat(6, 2, "hh:mm:ss");
                sheet.SetCellValue(7, 1, "Unsigned");
                sheet.SetCellValue(7, 2, (uint)123);
                sheet.SetCellFormat(7, 2, "000000");
                sheet.SetCellValue(8, 1, "Nullable");
                int? optional = 42;
                sheet.SetCellValue(8, 2, optional);
                sheet.SetCellFormat(8, 2, "0");
                sheet.SetCellFormula(9, 2, "SUM(B2)");

                document.Save(openExcel);
            }
        }
    }
}
