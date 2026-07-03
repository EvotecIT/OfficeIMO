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

        [Fact]
        public void Composer_LayoutHelpers_ReadOnlyListsDoNotSnapshotEnumerate() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var properties = new ThrowOnEnumerateReadOnlyList<(string Key, object? Value)>(
                ("Name", "Alice"),
                ("Score", 95),
                ("Status", "OK"));
            var kpis = new ThrowOnEnumerateReadOnlyList<(string Label, object? Value)>(
                ("Total", 2),
                ("Errors", 0));
            var urls = new ThrowOnEnumerateReadOnlyList<string>(
                "https://example.com",
                "https://evotec.xyz");

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.Compose("Details", c => {
                    c.PropertiesGrid(properties, columns: 2);
                    c.KpiRow(kpis, perRow: 2);
                    c.References(urls);
                    c.Finish(autoFitColumns: false);
                });

                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false)) {
                var ws = ss.WorkbookPart!.WorksheetParts.First();
                Assert.Equal("Name", GetCellText(ss, ws, "A1"));
                Assert.Equal("Alice", GetCellText(ss, ws, "B1"));
                Assert.Equal("Total", GetCellText(ss, ws, "A4"));
                Assert.Equal("2", GetCellText(ss, ws, "A5"));
                Assert.Equal("References", GetCellText(ss, ws, "A7"));
                Assert.Equal("example.com", GetCellText(ss, ws, "A8"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Composer_TableFrom_ReadOnlyListDoesNotSnapshotEnumerate() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var rows = new ThrowOnEnumerateReadOnlyList<ComposerTableRow>(
                new ComposerTableRow("Alpha", 10),
                new ComposerTableRow("Beta", 20));

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.Compose("Report", c => {
                    c.TableFrom(rows, title: "Scores");
                    c.Finish(autoFitColumns: false);
                });

                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false)) {
                var ws = ss.WorkbookPart!.WorksheetParts.First();
                Assert.True(ws.TableDefinitionParts.Any());
                Assert.Equal("Alpha", GetCellText(ss, ws, "A3"));
                Assert.Equal("10", GetCellText(ss, ws, "B3"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Composer_ColumnTableFrom_ReadOnlyListDoesNotSnapshotEnumerate() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var rows = new ThrowOnEnumerateReadOnlyList<ComposerTableRow>(
                new ComposerTableRow("Alpha", 10),
                new ComposerTableRow("Beta", 20));

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.Compose("Report", c => {
                    c.Columns(2, columns => {
                        columns[0].TableFrom(rows, title: "Scores");
                    });
                    c.Finish(autoFitColumns: false);
                });

                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false)) {
                var ws = ss.WorkbookPart!.WorksheetParts.First();
                Assert.True(ws.TableDefinitionParts.Any());
                Assert.Equal("Alpha", GetCellText(ss, ws, "A3"));
                Assert.Equal("10", GetCellText(ss, ws, "B3"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Composer_TableFrom_AppliesTableVisualStyleFlags() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var rows = new[] {
                new ComposerTableRow("Alpha", 10),
                new ComposerTableRow("Beta", 20)
            };

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.Compose("Report", c => {
                    c.TableFrom(
                        rows,
                        title: "Scores",
                        visuals: options => {
                            options.ShowFirstColumn = true;
                            options.ShowLastColumn = true;
                            options.ShowRowStripes = false;
                            options.ShowColumnStripes = true;
                        });
                    c.Finish(autoFitColumns: false);
                });

                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false)) {
                var tablePart = Assert.Single(ss.WorkbookPart!.WorksheetParts.First().TableDefinitionParts);
                Assert.NotNull(tablePart.Table);
                var styleInfo = Assert.IsType<TableStyleInfo>(tablePart.Table!.TableStyleInfo);
                Assert.True(styleInfo.ShowFirstColumn?.Value);
                Assert.True(styleInfo.ShowLastColumn?.Value);
                Assert.False(styleInfo.ShowRowStripes?.Value);
                Assert.True(styleInfo.ShowColumnStripes?.Value);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Composer_ColumnTableFrom_SummarizeOverflowPreservesMoreColumn() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var rows = new[] {
                new WideComposerTableRow("Alpha", 10, 5, 7),
                new WideComposerTableRow("Beta", 20, 6, 8)
            };

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.Compose("Report", c => {
                    c.Columns(2, columns => {
                        columns[0].TableFrom(rows, title: "Wide");
                    }, columnWidth: 2, overflow: OverflowMode.Summarize);
                    c.Finish(autoFitColumns: false);
                });

                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false)) {
                var ws = ss.WorkbookPart!.WorksheetParts.First();
                Assert.True(ws.TableDefinitionParts.Any());
                Assert.Equal("Metric A", GetCellText(ss, ws, "A2"));
                Assert.Equal("More", GetCellText(ss, ws, "B2"));
                Assert.Equal("5", GetCellText(ss, ws, "A3"));
                Assert.Contains("Name=Alpha", GetCellText(ss, ws, "B3"), StringComparison.Ordinal);
                Assert.Contains("Score=10", GetCellText(ss, ws, "B3"), StringComparison.Ordinal);
            }

            File.Delete(filePath);
        }

        private sealed class ComposerTableRow {
            public ComposerTableRow(string name, int score) {
                Name = name;
                Score = score;
            }

            public string Name { get; }

            public int Score { get; }
        }

        private sealed class WideComposerTableRow {
            public WideComposerTableRow(string name, int score, int metricA, int metricB) {
                Name = name;
                Score = score;
                MetricA = metricA;
                MetricB = metricB;
            }

            public string Name { get; }

            public int Score { get; }

            public int MetricA { get; }

            public int MetricB { get; }
        }

        private sealed class ThrowOnEnumerateReadOnlyList<T> : System.Collections.Generic.IReadOnlyList<T> {
            private readonly T[] _items;

            internal ThrowOnEnumerateReadOnlyList(params T[] items) {
                _items = items;
            }

            public int Count => _items.Length;

            public T this[int index] => _items[index];

            public System.Collections.Generic.IEnumerator<T> GetEnumerator() => throw new InvalidOperationException("Composer should use IReadOnlyList<T> indexing without snapshot enumeration.");

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }
}
