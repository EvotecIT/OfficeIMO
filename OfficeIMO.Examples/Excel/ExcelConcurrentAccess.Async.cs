using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates concurrent asynchronous loading of Excel documents.
    /// </summary>
    public class ExcelConcurrentAccessAsync {
        /// <summary>
        /// Opens the same workbook concurrently in read/write mode.
        /// </summary>
        /// <param name="folderPath">Path to the folder used for the workbook.</param>
        public static async Task Example_ExcelAsyncConcurrent(string folderPath) {
            Console.WriteLine("[*] Async concurrent load example for ExcelDocument");
            string filePath = Path.Combine(folderPath, "AsyncExcelConcurrent.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            await using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Sheet1");
                await document.SaveAsync();
            }

            var loadTask1 = ExcelDocument.LoadAsync(filePath, false);
            var loadTask2 = ExcelDocument.LoadAsync(filePath, false);

            var documents = await Task.WhenAll(loadTask1, loadTask2);

            await using var document1 = documents[0];
            await using var document2 = documents[1];
            Console.WriteLine($"Document1 sheets: {document1.Sheets.Count}");
            Console.WriteLine($"Document2 sheets: {document2.Sheets.Count}");

            File.Delete(filePath);
        }
    }
}

