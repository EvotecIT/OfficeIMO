using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates bulk cell updates using CellValuesParallel.
    /// </summary>
    public static class CellValuesParallel {
        public static void Example(string folderPath, bool openExcel) {
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

