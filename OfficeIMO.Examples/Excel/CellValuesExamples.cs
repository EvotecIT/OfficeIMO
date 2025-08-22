using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates setting cell values sequentially and in parallel.
    /// </summary>
    public static class CellValuesExamples {
        /// <summary>
        /// Sets various cell values sequentially.
        /// </summary>
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

                document.Save(openExcel);
            }
        }

        /// <summary>
        /// Demonstrates bulk cell updates using CellValuesParallel.
        /// </summary>
        public static void ExampleParallel(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Set cell values in parallel");
            string filePath = System.IO.Path.Combine(folderPath, "CellValuesParallel.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                var column1 = new List<(int Row, int Column, object Value)>();
                var column2 = new List<(int Row, int Column, object Value)>();

                for (int i = 1; i <= 50; i++) {
                    column1.Add((i, 1, $"A{i}"));
                    column2.Add((i, 2, $"B{i}"));
                }

                Task.WaitAll(
                    Task.Run(() => sheet.CellValuesParallel(column1)),
                    Task.Run(() => sheet.CellValuesParallel(column2))
                );

                document.Save(openExcel);
            }
        }
    }
}

