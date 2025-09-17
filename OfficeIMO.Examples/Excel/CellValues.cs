using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates setting various cell values.
    /// </summary>
    public static class CellValues {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Setting cell values");
            string filePath = System.IO.Path.Combine(folderPath, "Cell Values.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "Pi");
                sheet.CellValue(2, 2, Math.PI);
                sheet.CellValue(3, 1, "Today");
                sheet.CellValue(3, 2, DateTime.Today);
                sheet.FormatCell(3, 2, "yyyy-mm-dd");
                sheet.CellValue(4, 1, "IsCool");
                sheet.CellValue(4, 2, true);
                sheet.CellValue(5, 1, "Offset");
                sheet.CellValue(5, 2, DateTimeOffset.Now);
                sheet.FormatCell(5, 2, "yyyy-mm-dd hh:mm");
                sheet.CellValue(6, 1, "Duration");
                sheet.CellValue(6, 2, TimeSpan.FromHours(1.5));
                sheet.FormatCell(6, 2, "hh:mm:ss");
                sheet.CellValue(7, 1, "Unsigned");
                sheet.CellValue(7, 2, (uint)123);
                sheet.FormatCell(7, 2, "000000");
                sheet.CellValue(8, 1, "Nullable");
                int? optional = 42;
                sheet.CellValue(8, 2, optional);
                sheet.FormatCell(8, 2, "0");
                sheet.CellFormula(9, 2, "SUM(B2)");
                sheet.CellValue(10, 1, "Combined");
                sheet.Cell(10, 2, 1.23, "B2+1", "0.00");
#if NET6_0_OR_GREATER
                sheet.CellValue(11, 1, "DateOnly");
                sheet.CellValue(11, 2, DateOnly.FromDateTime(DateTime.Today));
                sheet.FormatCell(11, 2, "yyyy-mm-dd");
                sheet.CellValue(12, 1, "TimeOnly");
                sheet.CellValue(12, 2, new TimeOnly(1, 2, 3));
                sheet.FormatCell(12, 2, "hh:mm:ss");
#endif

                document.Save(openExcel);
            }
        }
    }
}
