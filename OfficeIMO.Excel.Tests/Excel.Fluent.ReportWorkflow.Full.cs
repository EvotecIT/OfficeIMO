using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelFluentReportWorkflowTests {
        [Fact]
        public void Fluent_Report_Workflow_EndToEnd() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                // Optional: monitor AddTable scanning timing for perf visibility (no assert, just exercise)
                doc.Execution.OnTiming = (op, elapsed) => { /* no-op in test */ };

                doc.Compose("Report", c =>
                {
                    c.Title("Demo Report", "Subtitle");
                    c.Callout("info", "Heads up", "Generated via fluent API");
                    c.Section("Summary");
                    c.PropertiesGrid(new (string, object?)[] {
                        ("Author", "Tester"),
                        ("Date", DateTime.Today.ToString("yyyy-MM-dd"))
                    });

                    var items = new[] {
                        new { Name = "Alice", Score = 90, Status = "OK" },
                        new { Name = "Bob", Score = 80, Status = "Warning" }
                    };
                    c.TableFrom(items, title: "Scores", visuals: v => {
                        v.NumericColumnDecimals["Score"] = 0;
                        v.TextBackgrounds["Status"] = new System.Collections.Generic.Dictionary<string, string> { { "Warning", "#FFF3CD" } };
                    });

                    c.References(new[] { "https://example.com" });
                    c.HeaderFooter(h => h.Center("Demo Report").FooterRight("Page &P of &N"));
                    c.Finish(autoFitColumns: true);
                });
                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(path, false)) {
                var ws = ss.WorkbookPart!.WorksheetParts.First();
                // Title exists
                Assert.Equal("Demo Report", GetText(ss, ws, "A1"));
                // A table exists
                Assert.True(ws.TableDefinitionParts.Any());
                // Named range created at top
                var dn = ss.WorkbookPart!.Workbook.DefinedNames;
                Assert.NotNull(dn);
                Assert.Contains(dn!.Elements<DefinedName>(), d => (d.Text ?? string.Empty).Contains("$A$1"));
                // Header/footer present
                var hf = ws.Worksheet.GetFirstChild<HeaderFooter>();
                Assert.NotNull(hf);
            }

            File.Delete(path);
        }

        private static string GetText(SpreadsheetDocument doc, WorksheetPart ws, string a1)
        {
            var cell = ws.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference != null && c.CellReference.Value == a1);
            if (cell == null) return string.Empty;
            var value = cell.CellValue?.Text ?? string.Empty;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                var sst = doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                if (sst != null && int.TryParse(value, out int id) && id >= 0 && id < sst.Count())
                    return sst.ChildElements[id].InnerText;
            }
            return value;
        }
    }
}

