using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using OfficeIMO.Excel;
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
            using (var doc = ExcelDocument.Create(filePath, "Data"))
            {
                var s = doc.Sheets[0];
                s.CellValue(1, 1, "Name");
                s.CellValue(1, 2, "Value");
                s.CellValue(1, 3, "Status");

                s.CellValue(2, 1, "Alpha");
                s.CellValue(2, 2, 10);
                s.CellValue(2, 3, "New");

                s.CellValue(3, 1, "Beta");
                s.CellValue(3, 2, 20);
                s.CellValue(3, 3, "New");

                doc.Save(openExcel);
            }

            // 2) Modify: read used range via sheet.Rows(), update cells and save again
            using (var doc = ExcelDocument.Load(filePath))
            {
                var s = doc["Data"];
                var rows = s.Rows().ToList();
                foreach (var row in rows)
                {
                    var name = Convert.ToString(row["Name"]);
                    var valueObj = row.ContainsKey("Value") ? row["Value"] : null;
                    int value = 0;
                    if (valueObj != null)
                    {
                        try { value = Convert.ToInt32(valueObj); } catch { /* ignore */ }
                    }

                    // Example rule: if Alpha has value 10 → set Value to 15 and Status to Processed
                    if (string.Equals(name, "Alpha", StringComparison.OrdinalIgnoreCase) && value == 10)
                    {
                        s.CellValue(2, 2, 15); // row 2, col 2
                        s.CellValue(2, 3, "Processed");
                    }
                    // If Beta → mark Status to Hold
                    if (string.Equals(name, "Beta", StringComparison.OrdinalIgnoreCase))
                    {
                        s.CellValue(3, 3, "Hold");
                    }
                }
                doc.Save(openExcel: openExcel);
            }

            // 3) Read again and emit JSON lines for PowerShell consumption
            using var doc2 = ExcelDocument.Load(filePath);
            var finalRows = doc2["Data"].Rows("A1:C3");
            var jsonOptions = new JsonSerializerOptions { WriteIndented = false };
            foreach (var r in finalRows)
            {
                Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(r, jsonOptions));
            }
        }
    }
}
