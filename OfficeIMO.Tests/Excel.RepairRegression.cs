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
                            int current = ColumnIndex(reference);
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
