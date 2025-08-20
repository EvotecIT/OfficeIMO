using System;
using System.Threading.Tasks;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates writing to a worksheet from multiple threads.
    /// </summary>
    public static class ConcurrentWrites {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Concurrent writes");
            string filePath = System.IO.Path.Combine(folderPath, "ConcurrentWrites.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                Parallel.For(1, 101, i => {
                    sheet.CellValue(i, 1, $"Value {i}");
                });
                document.Save(openExcel);
            }
        }
    }
}
