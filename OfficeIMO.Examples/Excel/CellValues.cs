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
                var sheet = document.AddSheet("Data");

                sheet.Cell(1, 1).SetValue("Name");
                sheet.Cell(1, 2).SetValue("Value");
                sheet.Cell(2, 1).SetValue("Pi");
                sheet.Cell(2, 2).SetValue(Math.PI);
                sheet.Cell(3, 1).SetValue("Today");
                sheet.Cell(3, 2).SetValue(DateTime.Today);
                sheet.Cell(3, 2).SetNumberFormat("yyyy-mm-dd");
                sheet.Cell(4, 1).SetValue("IsCool");
                sheet.Cell(4, 2).SetValue(true);
                sheet.Cell(5, 1).SetValue("Offset");
                sheet.Cell(5, 2).SetValue(DateTimeOffset.Now);
                sheet.Cell(5, 2).SetNumberFormat("yyyy-mm-dd hh:mm");
                sheet.Cell(6, 1).SetValue("Duration");
                sheet.Cell(6, 2).SetValue(TimeSpan.FromHours(1.5));
                sheet.Cell(6, 2).SetNumberFormat("hh:mm:ss");
                sheet.Cell(7, 1).SetValue("Unsigned");
                sheet.Cell(7, 2).SetValue((uint)123);
                sheet.Cell(7, 2).SetNumberFormat("000000");
                sheet.Cell(8, 1).SetValue("Nullable");
                int? optional = 42;
                sheet.Cell(8, 2).SetValue(optional);
                sheet.Cell(8, 2).SetNumberFormat("0");
                sheet.Cell(9, 2).SetFormula("SUM(B2)");
                sheet.Cell(10, 1).SetValue("Combined");
                sheet.Cell(10, 2).SetValue(1.23).SetFormula("B2+1").SetNumberFormat("0.00");
#if NET6_0_OR_GREATER
                sheet.Cell(11, 1).SetValue("DateOnly");
                sheet.Cell(11, 2).SetValue(DateOnly.FromDateTime(DateTime.Today));
                sheet.Cell(11, 2).SetNumberFormat("yyyy-mm-dd");
                sheet.Cell(12, 1).SetValue("TimeOnly");
                sheet.Cell(12, 2).SetValue(new TimeOnly(1, 2, 3));
                sheet.Cell(12, 2).SetNumberFormat("hh:mm:ss");
#endif

                document.Save();
                if (openExcel) document.OpenInApplication();
            }
        }
    }
}
