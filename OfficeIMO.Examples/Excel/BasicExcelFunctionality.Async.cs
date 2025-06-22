using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    public class BasicExcelFunctionalityAsync {
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
