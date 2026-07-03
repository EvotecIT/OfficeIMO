using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Utilities;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
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
                    var sheet = document.AddWorkSheet(name);
                    SeedWideContent(sheet, name, ref tableIndex);
                }

                document.Save(false);
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

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void MixedFeatureWorkbookValidatesWithoutRepairSignals() {
            string filePath = Path.Combine(_directoryWithFiles, "RepairRegression.MixedFeatures.xlsx");
            string logoPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using (var document = ExcelDocument.Create(filePath)) {
                var summary = document.AddWorkSheet("Summary");
                summary.CellValue(1, 1, "Label");
                summary.CellValue(1, 2, "Value");
                summary.CellValue(2, 1, "Status");
                summary.CellValue(2, 2, "Open");
                summary.SetComment(2, 2, "Line 1\nLine 2", author: "Tester", initials: "TT");
                summary.SetHyperlink(3, 1, "https://example.org", display: "Example");
                summary.SetInternalLink(4, 1, "'Summary'!A1", display: "Back to top");
                summary.SetHeaderFooter(headerCenter: "Repair Regression", headerRight: "Page &P of &N");
                if (File.Exists(logoPath)) {
                    summary.SetHeaderImage(HeaderFooterPosition.Center, File.ReadAllBytes(logoPath), "image/png", widthPoints: 96, heightPoints: 32);
                }

                var data = document.AddWorkSheet("Data");
                data.CellValue(1, 1, "Category");
                data.CellValue(1, 2, "Amount");
                data.CellValue(1, 3, "Trend1");
                data.CellValue(1, 4, "Trend2");

                data.CellValue(2, 1, "A");
                data.CellValue(2, 2, 10);
                data.CellValue(2, 3, 3d);
                data.CellValue(2, 4, 5d);

                data.CellValue(3, 1, "B");
                data.CellValue(3, 2, 20);
                data.CellValue(3, 3, 4d);
                data.CellValue(3, 4, 6d);

                data.CellValue(4, 1, "A");
                data.CellValue(4, 2, 15);
                data.CellValue(4, 3, 5d);
                data.CellValue(4, 4, 7d);

                data.AddAutoFilter("A1:D4");
                data.AddTable("A1:D4", hasHeader: true, name: "DataTable", OfficeIMO.Excel.TableStyle.TableStyleMedium9, includeAutoFilter: true);
                data.AddSparklines("C2:D4", "E2:E4", displayMarkers: true, seriesColor: "#FF0000");
                data.AddChartFromRange("A1:B4", row: 7, column: 1, widthPixels: 360, heightPixels: 220);
                data.AddPivotTable("A1:B4", "G2", name: "DataPivot");

                document.Save(filePath, false, new ExcelSaveOptions {
                    SafePreflight = true,
                    SafeRepairDefinedNames = true,
                    ValidateOpenXml = true
                });
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void SafeRepairDefinedNames_RemovesMalformedWorkbookNames() {
            string filePath = Path.Combine(_directoryWithFiles, "RepairRegression.DefinedNames.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data");
                var workbook = document._spreadSheetDocument.WorkbookPart!.Workbook;
                workbook.DefinedNames = new DefinedNames(
                    new DefinedName { Name = "DupName", Text = "'Data'!$A$1" },
                    new DefinedName { Name = "DupName", Text = "'Data'!$A$2" },
                    new DefinedName { Name = "BrokenRef", Text = "#REF!" },
                    new DefinedName { Name = "_xlnm.Print_Area", LocalSheetId = 0U, Text = "'Missing'!$A$1:$A$2" }
                );
                workbook.Save();

                document.Save(filePath, false, new ExcelSaveOptions {
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

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void WrapCellsWithWidthPinsColumnDuringAutoFit() {
            string filePath = Path.Combine(_directoryWithFiles, "WrapWidth.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "ID");
                sheet.CellValue(1, 2, "Description");
                for (int r = 2; r <= 6; r++) {
                    sheet.CellValue(r, 1, $"Row {r}");
                    sheet.CellValue(r, 2, "Wrapped column should stay narrow even when auto-fit runs on neighbours.");
                }

                sheet.WrapCells(2, 6, 2, 24);
                sheet.AutoFitColumnsFor(new[] { 1 });
                document.Save(false);
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
        public void ColumnSizingPinsWrappedAndAutoFitsOthers() {
            string filePath = Path.Combine(_directoryWithFiles, "ColumnSizing.WrapAutoFit.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var composer = new OfficeIMO.Excel.Fluent.SheetComposer(document, "Report");
                var sheet = composer.Sheet;

                sheet.CellValue(1, 1, "Wrap");
                sheet.CellValue(1, 2, "Auto");
                for (int r = 2; r <= 10; r++) {
                    sheet.CellValue(r, 1, "Long wrapped content that should stay constrained.");
                    sheet.CellValue(r, 2, "Extremely long value that benefits from auto-fitting to show full content without manual width tweaks.");
                }

                composer.ApplyColumnSizing("A1:B10", opts => {
                    opts.WrapHeaders.Add("Wrap");
                    opts.WrapWidth = 22;
                    opts.AutoFitHeaders.Add("Auto");
                });
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
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
            sheet.CellValue(1, 1, $"{label} Report");
            sheet.CellValue(1, 27, "AA marker");
            sheet.CellValue(1, 28, "AB marker");
            sheet.CellValue(1, 53, "BA marker");
            sheet.CellValue(1, 19, "S marker");
            sheet.CellValue(9, 15, "Detail O9");
            sheet.CellValue(9, 29, "Detail AC9");

            var safeLabel = label.Replace(' ', '_');
            AddTable(sheet, 11, 1, 4, 11, $"{safeLabel}_Primary_{tableIndex++}");
            AddTable(sheet, 20, 1, 5, 4, $"{safeLabel}_Summary_{tableIndex++}");
        }

        private static void AddTable(ExcelSheet sheet, int startRow, int startColumn, int rows, int columns, string tableName) {
            for (int c = 0; c < columns; c++) {
                sheet.CellValue(startRow, startColumn + c, $"Header {c + 1}");
            }

            for (int r = 1; r < rows; r++) {
                for (int c = 0; c < columns; c++) {
                    sheet.CellValue(startRow + r, startColumn + c, $"R{r}C{c}");
                }
            }

            string startColumnName = ColumnName(startColumn);
            string endColumnName = ColumnName(startColumn + columns - 1);
            string address = $"{startColumnName}{startRow}:{endColumnName}{startRow + rows - 1}";
            sheet.AddTable(address, hasHeader: true, tableName, OfficeIMO.Excel.TableStyle.TableStyleMedium9);
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
