using System;
using System.IO;
using System.Threading;
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

        /// <summary>
        /// Demonstrates cancelling an asynchronous save operation.
        /// </summary>
        /// <param name="folderPath">Path to the folder used for the workbook.</param>
        public static async Task Example_ExcelAsync_Cancel(string folderPath) {
            Console.WriteLine("[*] Async cancel example for ExcelDocument");
            string sourcePath = Path.Combine(folderPath, "AsyncSource.xlsx");
            string targetPath = Path.Combine(folderPath, "AsyncCancelled.xlsx");
            if (File.Exists(sourcePath)) File.Delete(sourcePath);
            if (File.Exists(targetPath)) File.Delete(targetPath);

            using (var document = ExcelDocument.Create(sourcePath)) {
                document.AddWorkSheet("Sheet1");
                using var cts = new CancellationTokenSource();
                var saveTask = document.SaveAsync(targetPath, false, cts.Token);
                cts.Cancel();
                try {
                    await saveTask;
                } catch (OperationCanceledException) {
                    Console.WriteLine("Save operation was cancelled.");
                }
            }

            Console.WriteLine($"File exists after cancel: {File.Exists(targetPath)}");
            if (File.Exists(sourcePath)) File.Delete(sourcePath);
            if (File.Exists(targetPath)) File.Delete(targetPath);
        }
    }
}
