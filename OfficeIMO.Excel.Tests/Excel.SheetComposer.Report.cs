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
        public void ComposerColumnSizingRejectsAQualifierThatWouldRetargetAnotherSheet() {
            using ExcelDocument document = ExcelDocument.Create();
            var composer = new SheetComposer(document, "Data");

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                composer.ApplyColumnSizing("'Other'!A1:B2", _ => { }));

            Assert.Contains("must not qualify", exception.Message, StringComparison.Ordinal);
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
        public void Composer_TableFrom_LeavesHeaderAppearanceToTheTableStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var rows = new[] { new ComposerTableRow("Alpha", 10) };

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.Compose("Report", c => {
                    c.TableFrom(rows, style: ExcelTableStyle.TableStyleMedium9);
                    c.Finish(autoFitColumns: false);
                });
                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false)) {
                var worksheet = ss.WorkbookPart!.WorksheetParts.First().Worksheet;
                var headerCells = worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value == "A1" || cell.CellReference?.Value == "B1")
                    .ToList();

                Assert.Equal(2, headerCells.Count);
                Assert.All(headerCells, cell => Assert.Null(cell.StyleIndex));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Composer_TableFrom_PreservesAdAsAHeaderAcronym() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var rows = new[] { new AdStateComposerTableRow("Enabled") };

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.Compose("Report", c => {
                    c.TableFrom(rows);
                    c.Finish(autoFitColumns: false);
                });
                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false)) {
                var worksheet = ss.WorkbookPart!.WorksheetParts.First();
                Assert.Equal("AD State", GetCellText(ss, worksheet, "A1"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Composer_ColumnTableFrom_LeavesHeaderAppearanceToTheTableStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var rows = new[] { new ComposerTableRow("Alpha", 10) };

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.Compose("Report", c => {
                    c.Columns(2, columns => {
                        columns[0].TableFrom(rows, style: ExcelTableStyle.TableStyleMedium4);
                    });
                    c.Finish(autoFitColumns: false);
                });
                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false)) {
                var worksheet = ss.WorkbookPart!.WorksheetParts.First().Worksheet;
                var headerCells = worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value == "A1" || cell.CellReference?.Value == "B1")
                    .ToList();

                Assert.Equal(2, headerCells.Count);
                Assert.All(headerCells, cell => Assert.Null(cell.StyleIndex));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void TableOfContents_LeavesHeaderAppearanceToTheTableStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.AddWorksheet("Data").Cell(1, 1, "Value");
                doc.AddTableOfContents(sheetName: "Index", styled: true);
                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false)) {
                var worksheet = ss.WorkbookPart!.WorksheetParts.First().Worksheet;
                var headerCells = worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value == "A3" || cell.CellReference?.Value == "B3")
                    .ToList();

                Assert.Equal(2, headerCells.Count);
                Assert.All(headerCells, cell => Assert.Null(cell.StyleIndex));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void EmptyTableOfContents_KeepsReadableStyledHeaders() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.AddTableOfContents(sheetName: "Index", styled: true);
                doc.Save();
            }

            using (var ss = SpreadsheetDocument.Open(filePath, false)) {
                var worksheet = ss.WorkbookPart!.WorksheetParts.First().Worksheet;
                var headerCells = worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value == "A3" || cell.CellReference?.Value == "B3")
                    .ToList();

                Assert.Equal(2, headerCells.Count);
                Assert.All(headerCells, cell => Assert.NotNull(cell.StyleIndex));
            }

            using (ExcelDocument doc = ExcelDocument.Load(filePath)) {
                ExcelCellStyleSnapshot firstHeader = doc.Sheets[0].GetCellStyle(3, 1);
                ExcelCellStyleSnapshot secondHeader = doc.Sheets[0].GetCellStyle(3, 2);
                Assert.True(firstHeader.Bold);
                Assert.True(secondHeader.Bold);
                Assert.EndsWith("F2F2F2", firstHeader.FillColorArgb, StringComparison.OrdinalIgnoreCase);
                Assert.EndsWith("F2F2F2", secondHeader.FillColorArgb, StringComparison.OrdinalIgnoreCase);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Composer_TableFrom_ManagedSvgResolvesBuiltInHeaderStyle() {
            using var doc = ExcelDocument.Create();
            ExcelSheet? report = null;
            doc.Compose("Report", c => {
                report = c.Sheet;
                c.TableFrom(new[] { new ComposerTableRow("Alpha", 10) }, style: ExcelTableStyle.TableStyleMedium9);
                c.Finish(autoFitColumns: false);
            });

            string svg = report!.Range("A1:B2").ToSvg(new ExcelImageExportOptions { ShowGridlines = false });
            string accentArgb = Assert.IsType<string>(doc.ResolveThemeColorArgb(4U));
            string expectedAccent = accentArgb.Substring(accentArgb.Length - 6);

            int headerTextIndex = svg.IndexOf(">Name</text>", StringComparison.Ordinal);
            Assert.True(headerTextIndex > 0);
            int headerElementStart = svg.LastIndexOf("<text ", headerTextIndex, StringComparison.Ordinal);
            Assert.True(headerElementStart >= 0);
            string headerTextElement = svg.Substring(headerElementStart, headerTextIndex - headerElementStart);
            Assert.Contains("font-weight=\"700\"", headerTextElement, StringComparison.Ordinal);
            Assert.Contains("fill=\"#FFFFFF\"", headerTextElement, StringComparison.OrdinalIgnoreCase);
            Assert.Contains($"fill=\"#{expectedAccent}\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Composer_TableFrom_ManagedSvgPreservesDirectHeaderStyle() {
            using var doc = ExcelDocument.Create();
            ExcelSheet? report = null;
            doc.Compose("Report", c => {
                report = c.Sheet;
                c.TableFrom(new[] { new ComposerTableRow("Alpha", 10) }, style: ExcelTableStyle.TableStyleMedium9);
                c.Finish(autoFitColumns: false);
            });
            report!.Range("A1").SetFillColor("C00000").SetFontColor("FFFF00");

            string svg = report.Range("A1:B2").ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            int headerTextIndex = svg.IndexOf(">Name</text>", StringComparison.Ordinal);
            Assert.True(headerTextIndex > 0);
            int headerElementStart = svg.LastIndexOf("<text ", headerTextIndex, StringComparison.Ordinal);
            Assert.True(headerElementStart >= 0);
            string headerTextElement = svg.Substring(headerElementStart, headerTextIndex - headerElementStart);
            Assert.Contains("fill=\"#FFFF00\"", headerTextElement, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("font-weight=\"700\"", headerTextElement, StringComparison.Ordinal);
            Assert.Contains("fill=\"#C00000\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Composer_TableFrom_ManagedSvgKeepsTableBoldWhenOnlyDirectFillIsSet() {
            using var doc = ExcelDocument.Create();
            ExcelSheet? report = null;
            doc.Compose("Report", c => {
                report = c.Sheet;
                c.TableFrom(new[] { new ComposerTableRow("Alpha", 10) }, style: ExcelTableStyle.TableStyleMedium9);
                c.Finish(autoFitColumns: false);
            });
            report!.Range("A1").SetFillColor("C00000");

            string svg = report.Range("A1:B2").ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            int headerTextIndex = svg.IndexOf(">Name</text>", StringComparison.Ordinal);
            Assert.True(headerTextIndex > 0);
            int headerElementStart = svg.LastIndexOf("<text ", headerTextIndex, StringComparison.Ordinal);
            Assert.True(headerElementStart >= 0);
            string headerTextElement = svg.Substring(headerElementStart, headerTextIndex - headerElementStart);
            Assert.Contains("font-weight=\"700\"", headerTextElement, StringComparison.Ordinal);
            Assert.Contains("fill=\"#FFFFFF\"", headerTextElement, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("fill=\"#C00000\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Composer_TableFrom_ManagedSvgPreservesDirectHeaderGradientFill() {
            using var doc = ExcelDocument.Create();
            ExcelSheet? report = null;
            doc.Compose("Report", c => {
                report = c.Sheet;
                c.TableFrom(new[] { new ComposerTableRow("Alpha", 10) }, style: ExcelTableStyle.TableStyleMedium9);
                c.Finish(autoFitColumns: false);
            });
            report!.CellAt(1, 1).SetGradientFill("C00000", "00A000", 45D);

            ExcelRangeVisualSnapshot snapshot = report.Range("A1:B2").CreateVisualSnapshot();
            string svg = report.Range("A1:B2").ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            ExcelVisualCell header = Assert.Single(snapshot.Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.Equal("FFC00000", header.Style.FillGradientStartColorArgb);
            Assert.Equal("FF00A000", header.Style.FillGradientEndColorArgb);
            Assert.Equal(45D, header.Style.FillGradientDegree);
            Assert.Contains("xl-gradient-1-1", svg, StringComparison.Ordinal);
            Assert.Contains("stop-color=\"#C00000\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stop-color=\"#00A000\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Composer_TableFrom_ManagedSvgProjectsLightTableHeaderBorder() {
            using var doc = ExcelDocument.Create();
            ExcelSheet? report = null;
            doc.Compose("Report", c => {
                report = c.Sheet;
                c.TableFrom(new[] { new ComposerTableRow("Alpha", 10) }, style: ExcelTableStyle.TableStyleLight1);
                c.Finish(autoFitColumns: false);
            });

            ExcelRangeVisualSnapshot snapshot = report!.Range("A1:B2").CreateVisualSnapshot();
            string svg = report.Range("A1:B2").ToSvg(new ExcelImageExportOptions { ShowGridlines = false });
            string borderArgb = Assert.IsType<string>(doc.ResolveThemeColorArgb(0U, -0.35D));
            string expectedBorder = borderArgb.Substring(borderArgb.Length - 6);

            ExcelVisualCell header = Assert.Single(snapshot.Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.Equal(expectedBorder, header.Style.Border?.Bottom?.ColorArgb, ignoreCase: true);
            Assert.Contains($"stroke=\"#{expectedBorder}\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Composer_TableFrom_ManagedSvgPreservesDirectHeaderBorder() {
            using var doc = ExcelDocument.Create();
            ExcelSheet? report = null;
            doc.Compose("Report", c => {
                report = c.Sheet;
                c.TableFrom(new[] { new ComposerTableRow("Alpha", 10) }, style: ExcelTableStyle.TableStyleLight1);
                c.Finish(autoFitColumns: false);
            });
            report!.CellAt(1, 1).SetBorder(ExcelBorderStyle.Double, "C00000");

            ExcelRangeVisualSnapshot snapshot = report.Range("A1:B2").CreateVisualSnapshot();
            string svg = report.Range("A1:B2").ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            ExcelVisualCell header = Assert.Single(snapshot.Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.Equal("double", header.Style.Border?.Bottom?.Style);
            Assert.Equal("FFC00000", header.Style.Border?.Bottom?.ColorArgb);
            Assert.Contains("stroke=\"#C00000\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void TableHeaderVisualStyleResolver_PreservesDirectPatternFill() {
            var directStyle = new ExcelCellStyleSnapshot {
                StyleIndex = 1U,
                FillPatternType = "darkGrid",
                FillPatternForegroundColorArgb = "FFC00000",
                FillPatternBackgroundColorArgb = "FFFFE5E5"
            };
            var tableStyle = new ExcelTableHeaderVisualStyle("FF4472C4", "FFFFFFFF", bold: true, borderColorArgb: "FF4472C4");

            ExcelCellStyleSnapshot resolved = ExcelTableHeaderVisualStyleResolver.Apply(directStyle, tableStyle);

            Assert.Equal("darkGrid", resolved.FillPatternType);
            Assert.Equal("FFC00000", resolved.FillPatternForegroundColorArgb);
            Assert.Equal("FFFFE5E5", resolved.FillPatternBackgroundColorArgb);
        }

        [Fact]
        public void Composer_TableFrom_ManagedSvgPreservesNonBoldDarkHeaderContract() {
            using var doc = ExcelDocument.Create();
            ExcelSheet? report = null;
            doc.Compose("Report", c => {
                report = c.Sheet;
                c.TableFrom(new[] { new ComposerTableRow("Alpha", 10) }, style: ExcelTableStyle.TableStyleDark8);
                c.Finish(autoFitColumns: false);
            });

            string svg = report!.Range("A1:B2").ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            int headerTextIndex = svg.IndexOf(">Name</text>", StringComparison.Ordinal);
            Assert.True(headerTextIndex > 0);
            int headerElementStart = svg.LastIndexOf("<text ", headerTextIndex, StringComparison.Ordinal);
            Assert.True(headerElementStart >= 0);
            string headerTextElement = svg.Substring(headerElementStart, headerTextIndex - headerElementStart);
            Assert.DoesNotContain("font-weight=\"700\"", headerTextElement, StringComparison.Ordinal);
            Assert.Contains("fill=\"#FFFFFF\"", headerTextElement, StringComparison.OrdinalIgnoreCase);
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

        private sealed class AdStateComposerTableRow {
            public AdStateComposerTableRow(string adState) {
                ADState = adState;
            }

            public string ADState { get; }
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
