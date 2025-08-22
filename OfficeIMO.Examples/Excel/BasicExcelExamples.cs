using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates basic synchronous and asynchronous Excel workflows.
    /// </summary>
    public static class BasicExcelExamples {
        /// <summary>
        /// Creates a workbook, saves it and loads it again.
        /// </summary>
        /// <param name="folderPath">Target folder for the workbook.
        /// <param name="openExcel">When set to <c>true</c>, opens the workbook after saving.
        public static void BasicExample(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Basic example");
            string filePath = Path.Combine(folderPath, "BasicExcelExample.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Sheet1");
                document.Save(openExcel);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                Console.WriteLine($"Sheets count: {document.Sheets.Count}");
            }
        }

        /// <summary>
        /// Creates, saves and loads a workbook asynchronously.
        /// </summary>
        /// <param name="folderPath">Target folder for the workbook.
        public static async Task BasicExampleAsync(string folderPath) {
            Console.WriteLine("[*] Excel - Basic async example");
            string filePath = Path.Combine(folderPath, "AsyncBasicExcelExample.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            await using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Sheet1");
                await document.SaveAsync();
            }

            await using (ExcelDocument document = await ExcelDocument.LoadAsync(filePath)) {
                Console.WriteLine($"Sheets count: {document.Sheets.Count}");
            }
        }
    }
}