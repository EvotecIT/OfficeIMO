using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelSheetComposerReportTests {
        private static string GetCellText(SpreadsheetDocument doc, WorksheetPart ws, string a1)
        {
            var cell = ws.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference != null && c.CellReference.Value == a1);
            if (cell == null) return string.Empty;
            var value = cell.CellValue?.Text ?? string.Empty;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var sst = doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                if (sst != null && int.TryParse(value, out int idx) && idx >= 0 && idx < sst.Count())
                    return sst.ChildElements[idx].InnerText;
            }
            return value;
        }

        [Fact]
        public void Composer_Callout_WritesTitleAndBody() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(filePath))
            {
                doc.Compose("Summary", c =>
                {
                    c.Title("Report");
                    c.Callout("warning", "Heads up", "This is a caution.");
                    c.Paragraph("After callout");
                    c.Finish();
                });
                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false))
            {
                var ws = ss.WorkbookPart!.WorksheetParts.First();
                // Title at A1; callout title at A3; callout body at A4; paragraph at A6
                Assert.Equal("Report", GetCellText(ss, ws, "A1"));
                Assert.Equal("Heads up", GetCellText(ss, ws, "A3"));
                Assert.Equal("This is a caution.", GetCellText(ss, ws, "A4"));
                Assert.Equal("After callout", GetCellText(ss, ws, "A6"));
            }
            File.Delete(filePath);
        }

        [Fact]
        public void Composer_PropertiesGrid_WritesKeyValuePairs() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(filePath))
            {
                doc.Compose("Details", c =>
                {
                    c.Section("Meta");
                    c.PropertiesGrid(new (string, object?)[]
                    {
                        ("Name", "Alice"),
                        ("Score", 95),
                        ("Status", "OK")
                    }, columns: 2);
                    c.Finish();
                });
                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false))
            {
                var ws = ss.WorkbookPart!.WorksheetParts.First();
                // Section header at A1; then first row of grid at A2/B2 and C2/D2
                Assert.Equal("Meta", GetCellText(ss, ws, "A1"));
                Assert.Equal("Name", GetCellText(ss, ws, "A2"));
                Assert.Equal("Alice", GetCellText(ss, ws, "B2"));
                Assert.Equal("Score", GetCellText(ss, ws, "C2"));
                Assert.Equal("95", GetCellText(ss, ws, "D2"));
                // Next row contains the remaining key/value
                Assert.Equal("Status", GetCellText(ss, ws, "A3"));
                Assert.Equal("OK", GetCellText(ss, ws, "B3"));
            }
            File.Delete(filePath);
        }
    }
}

