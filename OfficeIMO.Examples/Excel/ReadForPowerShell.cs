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
            using (var doc = ExcelDocument.Create(filePath, new ExcelCreateOptions { FirstSheetName = "Data" }))
            {
                var s = doc.Sheets[0];
                s.Cell(1, 1).SetValue("Name");
                s.Cell(1, 2).SetValue("Value");
                s.Cell(2, 1).SetValue("Alpha");
                s.Cell(2, 2).SetValue(10);
                s.Cell(3, 1).SetValue("Beta");
                s.Cell(3, 2).SetValue(20);
                doc.Save();
                if (openExcel) doc.OpenInApplication();
            }

            // The OfficeIMO.Excel reader works for XLSX, XLSM, and XLSB without an A1 range.
            var jsonOptions = new JsonSerializerOptions { WriteIndented = false };
            using var reader = ExcelDocument.OpenReader(filePath);
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
