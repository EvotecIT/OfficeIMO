using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Read;
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
            using (var doc = ExcelDocument.Create(path, "Data"))
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

                doc.Save(openExcel: false);
            }

            // Read whole sheet as dictionaries
            var rows = ExcelRead.ReadUsedRangeObjects(path, "Data", ExcelReadPresets.Simple());
            Assert.Equal(2, rows.Count); // two data rows

            // Modify and write again
            using (var doc = ExcelDocument.Load(path))
            {
                var s = doc.Sheets[0];
                foreach (var row in rows)
                {
                    string name = Convert.ToString(row["Name"]);
                    int value = 0;
                    if (row.TryGetValue("Value", out var val) && val != null)
                    {
                        try { value = Convert.ToInt32(val); } catch { }
                    }

                    if (string.Equals(name, "Alpha", StringComparison.OrdinalIgnoreCase) && value == 10)
                    {
                        s.CellValue(2, 2, 15);
                        s.CellValue(2, 3, "Processed");
                    }
                    if (string.Equals(name, "Beta", StringComparison.OrdinalIgnoreCase))
                    {
                        s.CellValue(3, 3, "Hold");
                    }
                }
                doc.Save(openExcel: false);
            }

            // Read back and assert changes
            var finalRows = ExcelRead.ReadRangeObjects(path, "Data", "A1:C3", ExcelReadPresets.Simple());
            Assert.Equal(2, finalRows.Count);

            var alpha = finalRows.Find(d => string.Equals(Convert.ToString(d["Name"]), "Alpha", StringComparison.OrdinalIgnoreCase));
            var beta = finalRows.Find(d => string.Equals(Convert.ToString(d["Name"]), "Beta", StringComparison.OrdinalIgnoreCase));

            Assert.NotNull(alpha);
            Assert.NotNull(beta);

            Assert.Equal(15, Convert.ToInt32(alpha["Value"]));
            Assert.Equal("Processed", Convert.ToString(alpha["Status"]));
            Assert.Equal("Hold", Convert.ToString(beta["Status"]));
        }
    }
}
