using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel
{
    /// <summary>
    /// PowerShell-style workflow: write → read → modify → write → emit JSON.
    /// In PowerShell you can pipe JSON to objects with ConvertFrom-Json.
    /// </summary>
    internal static class PowerShellRoundTrip
    {
        public static void Example(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - PowerShell round-trip (write/read/write/json)");
            string filePath = Path.Combine(folderPath, "PS-RoundTrip.xlsx");

            // 1) Write: create workbook with a simple sheet
            using (var doc = ExcelDocument.Create(filePath, new ExcelCreateOptions { FirstSheetName = "Data" }))
            {
                var s = doc.Sheets[0];
                s.Cell(1, 1).SetValue("Name");
                s.Cell(1, 2).SetValue("Value");
                s.Cell(1, 3).SetValue("Status");

                s.Cell(2, 1).SetValue("Alpha");
                s.Cell(2, 2).SetValue(10);
                s.Cell(2, 3).SetValue("New");

                s.Cell(3, 1).SetValue("Beta");
                s.Cell(3, 2).SetValue(20);
                s.Cell(3, 3).SetValue("New");

                doc.Save();
                if (openExcel) doc.OpenInApplication();
            }

            // 2) Modify: read through the package-owned Excel API, then update and save
            var updates = new List<(int RowNumber, int? Value, string Status)>();
            using (var reader = ExcelDocument.OpenReader(
                filePath,
                new ExcelReaderOptions { SheetName = "Data" }))
            {
                int rowNumber = 2;
                while (reader.Read())
                {
                    string? name = reader.IsDBNull(0) ? null : Convert.ToString(reader.GetValue(0));
                    int value = reader.IsDBNull(1) ? 0 : Convert.ToInt32(reader.GetValue(1));

                    if (string.Equals(name, "Alpha", StringComparison.OrdinalIgnoreCase) && value == 10)
                    {
                        updates.Add((rowNumber, 15, "Processed"));
                    }
                    else if (string.Equals(name, "Beta", StringComparison.OrdinalIgnoreCase))
                    {
                        updates.Add((rowNumber, null, "Hold"));
                    }

                    rowNumber++;
                }
            }

            using (var doc = ExcelDocument.Load(filePath))
            {
                var s = doc["Data"];
                foreach (var update in updates)
                {
                    if (update.Value.HasValue)
                    {
                        s.Cell(update.RowNumber, 2).SetValue(update.Value.Value);
                    }
                    s.Cell(update.RowNumber, 3).SetValue(update.Status);
                }
                doc.Save();
                if (openExcel) doc.OpenInApplication();
            }

            // 3) Read again and emit JSON lines for PowerShell consumption
            using var finalReader = ExcelDocument.OpenReader(
                filePath,
                new ExcelReaderOptions { SheetName = "Data" });
            var jsonOptions = new JsonSerializerOptions { WriteIndented = false };
            while (finalReader.Read())
            {
                var row = new Dictionary<string, object?>(finalReader.FieldCount, StringComparer.OrdinalIgnoreCase);
                for (int ordinal = 0; ordinal < finalReader.FieldCount; ordinal++)
                {
                    row[finalReader.GetName(ordinal)] = finalReader.IsDBNull(ordinal)
                        ? null
                        : finalReader.GetValue(ordinal);
                }
                Console.WriteLine(JsonSerializer.Serialize(row, jsonOptions));
            }
        }
    }
}
