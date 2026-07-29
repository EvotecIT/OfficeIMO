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
                doc.Save();
                if (openExcel) doc.OpenInApplication();
            }

            // The OfficeIMO.Excel reader works for XLSX, XLSM, and XLSB without an A1 range.
            var jsonOptions = new JsonSerializerOptions { WriteIndented = false };
            using var reader = ExcelDocument.OpenDataReader(filePath);
            while (reader.Read())
            {
                var row = new Dictionary<string, object?>(reader.FieldCount, StringComparer.OrdinalIgnoreCase);
                for (int column = 0; column < reader.FieldCount; column++)
                    row[reader.GetName(column)] = reader.IsDBNull(column) ? null : reader.GetValue(column);
                Console.WriteLine(JsonSerializer.Serialize(row, jsonOptions));
            }
        }
    }
}
