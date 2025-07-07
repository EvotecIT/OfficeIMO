using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates asynchronous operations with <see cref="ExcelDocument"/>.
    /// </summary>
    public class BasicExcelFunctionalityAsync {
        /// <summary>
        /// Creates a workbook, saves it and loads it asynchronously.
        /// </summary>
        /// <param name="folderPath">Path to the folder used for the workbook.</param>
        public static async Task Example_ExcelAsync(string folderPath) {
            Console.WriteLine("[*] Async example for ExcelDocument");
            string filePath = Path.Combine(folderPath, "AsyncExcel.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Sheet1");
                await document.SaveAsync();
            }

            using (var document = await ExcelDocument.LoadAsync(filePath)) {
                Console.WriteLine($"Sheet count: {document.Sheets.Count}");
            }

            File.Delete(filePath);
        }
    }
}
