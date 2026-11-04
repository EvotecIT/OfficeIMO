using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Utilities;
using Xunit;

namespace OfficeIMO.Tests
{
    public partial class Excel
    {
        [Fact]
        public void ComplexWorkbookMaintainsCellOrderAndPackageIntegrity() {
            string filePath = Path.Combine(_directoryWithFiles, "RepairRegression.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheetNames = new[] {
                    "Overview", "evotec.pl", "evotec.xyz", "Summary", "Matrix",
                    "SPF Providers", "Recommendations", "References", "Index"
                };

                int tableIndex = 1;
                foreach (var name in sheetNames) {
                    var sheet = document.AddSheet(name);
                    SeedWideContent(sheet, name, ref tableIndex);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                foreach (var worksheetPart in spreadsheet.WorkbookPart!.WorksheetParts) {
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    if (sheetData == null) continue;
                    foreach (var row in sheetData.Elements<Row>()) {
                        int previousColumn = 0;
                        foreach (var cell in row.Elements<Cell>()) {
                            var reference = cell.CellReference?.Value;
                            if (string.IsNullOrEmpty(reference)) continue;
                            var referenceNonNull = reference!;
                            int current = ColumnIndex(referenceNonNull);
                            Assert.True(current >= previousColumn,
                                $"Row {row.RowIndex} contains out-of-order cell '{reference}'.");
                            previousColumn = current;
                        }
                    }
                }
            }

            var summary = ExcelPackageUtilities.GetContentTypesSummary(filePath);
            Assert.True(summary.HasXmlDefault, "Missing XML default content-type override.");
            Assert.Equal("application/xml", summary.XmlDefaultContentType, ignoreCase: true);
            Assert.Equal(1, summary.XmlDefaultCount);
            Assert.True(summary.HasWorkbookOverride, "Workbook override entry is missing.");
            Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                summary.WorkbookContentType, ignoreCase: true);

            Assert.False(ExcelPackageUtilities.NormalizeContentTypes(filePath),
                "Content types should already be normalized after save.");

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void MixedFeatureWorkbookValidatesWithoutRepairSignals()
        {
            string filePath = Path.Combine(_directoryWithFiles, "RepairRegression.MixedFeatures.xlsx");
            string logoPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            using (var document = ExcelDocument.Create(filePath))
            {
                var summary = document.AddSheet("Summary");
                summary.Cell(1, 1).SetValue("Label");
                summary.Cell(1, 2).SetValue("Value");
                summary.Cell(2, 1).SetValue("Status");
                summary.Cell(2, 2).SetValue("Open");
                summary.SetComment(2, 2, "Line 1\nLine 2", author: "Tester", initials: "TT");
                summary.SetHyperlink(3, 1, "https://example.org", display: "Example");
                summary.SetInternalLink(4, 1, "'Summary'!A1", display: "Back to top");
                summary.PageSetup.Header.Center.Text = "Repair Regression";
                summary.PageSetup.Header.Right.Text = "Page &P of &N";
                if (File.Exists(logoPath))
                {
                    summary.PageSetup.Header.Center.SetImage(File.ReadAllBytes(logoPath), "image/png", width: OfficeLength.Points(96), height: OfficeLength.Points(32));
                }

                var data = document.AddSheet("Data");
                data.Cell(1, 1).SetValue("Category");
                data.Cell(1, 2).SetValue("Amount");
                data.Cell(1, 3).SetValue("Trend1");
                data.Cell(1, 4).SetValue("Trend2");
                data.Cell(2, 1).SetValue("A");
                data.Cell(2, 2).SetValue(10);
                data.Cell(2, 3).SetValue(3d);
                data.Cell(2, 4).SetValue(5d);
                data.Cell(3, 1).SetValue("B");
                data.Cell(3, 2).SetValue(20);
                data.Cell(3, 3).SetValue(4d);
                data.Cell(3, 4).SetValue(6d);
                data.Cell(4, 1).SetValue("A");
                data.Cell(4, 2).SetValue(15);
                data.Cell(4, 3).SetValue(5d);
                data.Cell(4, 4).SetValue(7d);
                data.Range("A1:D4").AutoFilter();
                data.AddTable("A1:D4", hasHeader: true, name: "DataTable", OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium9, includeAutoFilter: true);
                data.AddSparklines("C2:D4", "E2:E4", displayMarkers: true, seriesColor: OfficeIMO.OfficeColor.ParseHex("#FF0000"));
                data.AddChart(OfficeIMO.OfficeChartType.ColumnClustered, "A1:B4", new OfficeIMO.Excel.ExcelChartOptions { Row = 7, Column = 1, Width = OfficeLength.Pixels(360), Height = OfficeLength.Pixels(220) });
                data.AddPivotTable("A1:B4", "G2", options: new ExcelPivotTableOptions { Name = "DataPivot" });
                document.Save(filePath, new ExcelSaveOptions { SafePreflight = true, SafeRepairDefinedNames = true, ValidateOpenXml = true });
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void SafeRepairDefinedNames_RemovesMalformedWorkbookNames() {
            string filePath = Path.Combine(_directoryWithFiles, "RepairRegression.DefinedNames.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddSheet("Data");
                var workbook = document._spreadSheetDocument.WorkbookPart!.Workbook;
                workbook.DefinedNames = new DefinedNames(
                    new DefinedName { Name = "DupName", Text = "'Data'!$A$1" },
                    new DefinedName { Name = "DupName", Text = "'Data'!$A$2" },
                    new DefinedName { Name = "BrokenRef", Text = "#REF!" },
                    new DefinedName { Name = "_xlnm.Print_Area", LocalSheetId = 0U, Text = "'Missing'!$A$1:$A$2" }
                );
                workbook.Save();

                document.Save(filePath, new ExcelSaveOptions {
                    SafePreflight = true,
                    SafeRepairDefinedNames = true,
                    ValidateOpenXml = true
                });
            }

            using (var package = SpreadsheetDocument.Open(filePath, false)) {
                var names = package.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().ToList();
                Assert.Single(names);
                Assert.Equal("DupName", names[0].Name?.Value);
                Assert.Equal("'Data'!$A$1", names[0].Text);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void WrapCellsWithWidthPinsColumnDuringAutoFit() {
            string filePath = Path.Combine(_directoryWithFiles, "WrapWidth.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("ID");
                sheet.Cell(1, 2).SetValue("Description");
                for (int r = 2; r <= 6; r++) {
                    sheet.Cell(r, 1).SetValue($"Row {r}");
                    sheet.Cell(r, 2).SetValue("Wrapped column should stay narrow even when auto-fit runs on neighbours.");
                }

                sheet.Range(2, 2, 6, 2).SetWrapText(); sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(24));
                sheet.AutoFitColumnsFor(new[] { 1 });
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                double? wrappedWidth = GetColumnWidth(wsPart, 2);
                Assert.True(wrappedWidth.HasValue);
                Assert.InRange(wrappedWidth!.Value, 23.5, 24.5);

                double? autoWidth = GetColumnWidth(wsPart, 1);
                Assert.True(autoWidth.HasValue);
                Assert.NotEqual(wrappedWidth.Value, autoWidth.Value);
            }
        }

        [Fact]
        public void ColumnSizingPinsWrappedAndAutoFitsOthers()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ColumnSizing.WrapAutoFit.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var composer = document.AddSheet("Report").Report();
                var sheet = composer.Sheet;
                sheet.Cell(1, 1).SetValue("Wrap");
                sheet.Cell(1, 2).SetValue("Auto");
                for (int r = 2; r <= 10; r++)
                {
                    sheet.Cell(r, 1).SetValue("Long wrapped content that should stay constrained.");
                    sheet.Cell(r, 2).SetValue("Extremely long value that benefits from auto-fitting to show full content without manual width tweaks.");
                }

                composer.ApplyColumnSizing("A1:B10", opts =>
                {
                    opts.WrapHeaders.Add("Wrap");
                    opts.WrapWidth = ExcelColumnWidth.Characters(22);
                    opts.AutoFitHeaders.Add("Auto");
                });
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                double? wrapWidth = GetColumnWidth(wsPart, 1);
                double? autoWidth = GetColumnWidth(wsPart, 2);
                Assert.True(wrapWidth.HasValue && autoWidth.HasValue);
                Assert.InRange(wrapWidth!.Value, 21.5, 22.5);
                Assert.True(autoWidth!.Value > wrapWidth.Value + 2, "Auto-fit column should be wider than the pinned wrap column.");
            }
        }

        private static double? GetColumnWidth(WorksheetPart part, int columnIndex)
        {
            var columns = part.Worksheet.GetFirstChild<Columns>();
            if (columns == null) return null;
            foreach (var column in columns.Elements<Column>())
            {
                uint min = column.Min?.Value ?? 0;
                uint max = column.Max?.Value ?? 0;
                if (columnIndex >= min && columnIndex <= max)
                {
                    return column.Width?.Value;
                }
            }
            return null;
        }

        private static void SeedWideContent(ExcelSheet sheet, string label, ref int tableIndex) {
            // Intentionally write cells out of order across the alphabet boundary to exercise ordering logic.
            sheet.Cell(1, 1).SetValue($"{label} Report");
            sheet.Cell(1, 27).SetValue("AA marker");
            sheet.Cell(1, 28).SetValue("AB marker");
            sheet.Cell(1, 53).SetValue("BA marker");
            sheet.Cell(1, 19).SetValue("S marker");
            sheet.Cell(9, 15).SetValue("Detail O9");
            sheet.Cell(9, 29).SetValue("Detail AC9");

            var safeLabel = label.Replace(' ', '_');
            AddTable(sheet, 11, 1, 4, 11, $"{safeLabel}_Primary_{tableIndex++}");
            AddTable(sheet, 20, 1, 5, 4, $"{safeLabel}_Summary_{tableIndex++}");
        }

        private static void AddTable(ExcelSheet sheet, int startRow, int startColumn, int rows, int columns, string tableName) {
            for (int c = 0; c < columns; c++) {
                sheet.Cell(startRow, startColumn + c).SetValue($"Header {c + 1}");
            }

            for (int r = 1; r < rows; r++) {
                for (int c = 0; c < columns; c++) {
                    sheet.Cell(startRow + r, startColumn + c).SetValue($"R{r}C{c}");
                }
            }

            string startColumnName = ColumnName(startColumn);
            string endColumnName = ColumnName(startColumn + columns - 1);
            string address = $"{startColumnName}{startRow}:{endColumnName}{startRow + rows - 1}";
            sheet.AddTable(address, hasHeader: true, tableName, OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium9);
        }

        private static int ColumnIndex(string reference) {
            int result = 0;
            foreach (char ch in reference) {
                if (char.IsLetter(ch)) {
                    result = result * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
                } else {
                    break;
                }
            }
            return result;
        }

        private static string ColumnName(int columnIndex) {
            if (columnIndex <= 0) throw new ArgumentOutOfRangeException(nameof(columnIndex));
            int dividend = columnIndex;
            var name = new System.Text.StringBuilder();
            while (dividend > 0) {
                int modulo = (dividend - 1) % 26;
                name.Insert(0, (char)('A' + modulo));
                dividend = (dividend - modulo) / 26;
            }
            return name.ToString();
        }
    }
}