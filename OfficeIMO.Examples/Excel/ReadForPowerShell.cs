using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel
{
    /// <summary>
    /// Demonstrates PowerShell-friendly reading: emit JSON rows so PS can pipe into ConvertFrom-Json.
    /// In PowerShell, you can call the compiled examples and pipe the output:
    ///   dotnet OfficeIMO.Examples.dll | ConvertFrom-Json
    /// </summary>
    internal static class ReadForPowerShell
    {
        public static void Example(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Read for PowerShell");
            string filePath = Path.Combine(folderPath, "Read-ForPowerShell.xlsx");

            // Create a tiny workbook
            using (var doc = ExcelDocument.Create(filePath, "Data"))
            {
                var s = doc.Sheets[0];
                s.CellValue(1, 1, "Name");
                s.CellValue(1, 2, "Value");
                s.CellValue(2, 1, "Alpha");
                s.CellValue(2, 2, 10);
                s.CellValue(3, 1, "Beta");
                s.CellValue(3, 2, 20);
                doc.Save(openExcel);
            }

            // Read dictionaries with Simple converters (optional)
            var rows = ExcelRead.ReadRangeObjects(filePath, "Data", "A1:B3", ExcelReadPresets.Simple());

            // Emit one JSON object per row (PowerShell-friendly)
            var jsonOptions = new JsonSerializerOptions { WriteIndented = false };
            foreach (var row in rows)
            {
                Console.WriteLine(JsonSerializer.Serialize(row, jsonOptions));
            }
        }
    }
}
