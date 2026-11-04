using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public class ExcelPowerShellRoundTrip
    {
        [Fact]
        public void RoundTrip_ReadModifyWrite_NoLockConflicts_ValuesUpdated()
        {
            string folder = Path.Combine(AppContext.BaseDirectory, "Documents");
            Directory.CreateDirectory(folder);
            string path = Path.Combine(folder, "PS-RoundTrip-Test.xlsx");

            if (File.Exists(path)) File.Delete(path);

            // Write initial file
            using (var doc = ExcelDocument.Create(path, new ExcelCreateOptions { FirstSheetName = "Data" }))
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
            }

            // Read whole sheet as dictionaries
            var rows = ReadRows(path);
            Assert.Equal(2, rows.Count); // two data rows

            // Modify and write again
            using (var doc = ExcelDocument.Load(path))
            {
                var s = doc.Sheets[0];
                    foreach (var row in rows)
                    {
                        string name = Convert.ToString(row["Name"]) ?? string.Empty;
                        int value = 0;
                        if (row.TryGetValue("Value", out var val) && val != null)
                        {
                            try { value = Convert.ToInt32(val); } catch { }
                        }

                    if (string.Equals(name, "Alpha", StringComparison.OrdinalIgnoreCase) && value == 10)
                    {
                        s.Cell(2, 2).SetValue(15);
                        s.Cell(2, 3).SetValue("Processed");
                    }
                    if (string.Equals(name, "Beta", StringComparison.OrdinalIgnoreCase))
                    {
                        s.Cell(3, 3).SetValue("Hold");
                    }
                }
                doc.Save();
            }

            // Read back and assert changes
            var finalRows = ReadRows(path);
            Assert.Equal(2, finalRows.Count);

            var alpha = finalRows.Find(d => string.Equals(Convert.ToString(d["Name"]), "Alpha", StringComparison.OrdinalIgnoreCase));
            var beta = finalRows.Find(d => string.Equals(Convert.ToString(d["Name"]), "Beta", StringComparison.OrdinalIgnoreCase));

            Assert.NotNull(alpha);
            Assert.NotNull(beta);

            Assert.Equal(15, Convert.ToInt32(alpha["Value"]));
            Assert.Equal("Processed", Convert.ToString(alpha["Status"]));
            Assert.Equal("Hold", Convert.ToString(beta["Status"]));
        }

        private static List<Dictionary<string, object?>> ReadRows(string path)
        {
            using DbDataReader reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions
            {
                SheetName = "Data"
            });
            var rows = new List<Dictionary<string, object?>>();
            while (reader.Read())
            {
                var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (int ordinal = 0; ordinal < reader.FieldCount; ordinal++)
                {
                    row[reader.GetName(ordinal)] = reader.IsDBNull(ordinal)
                        ? null
                        : reader.GetValue(ordinal);
                }
                rows.Add(row);
            }
            return rows;
        }
    }
}
