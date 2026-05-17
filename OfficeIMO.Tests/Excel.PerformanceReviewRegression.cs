using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void PerformanceReview_LoadedWorkbookDirectAppend_PersistsAfterSave() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.DirectAppendDirty.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets[0];
                var cells = Enumerable.Range(2, 20)
                    .Select(row => (row, 1, (object)("Row " + row.ToString(System.Globalization.CultureInfo.InvariantCulture))))
                    .ToList();
                sheet.CellValues(cells, ExecutionMode.Sequential);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.True(document.Sheets[0].TryGetCellText(21, 1, out string? text));
                Assert.Equal("Row 21", text);
            }
        }

        [Fact]
        public void PerformanceReview_LoadedWorkbookProtection_PersistsAfterSave() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.WorkbookProtectionDirty.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                    ProtectStructure = true,
                    LegacyPasswordHash = "CAFE"
                });
                document.Save();
            }

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var protection = spreadsheet.WorkbookPart!.Workbook.GetFirstChild<WorkbookProtection>();
            Assert.NotNull(protection);
            Assert.True(protection!.LockStructure!.Value);
            Assert.Equal("CAFE", protection.WorkbookPassword!.Value);
        }

        [Fact]
        public void PerformanceReview_StreamSaveWithPackageProperties_PreservesProperties() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Value");
                document.BuiltinDocumentProperties.Title = "Performance Review";
                document.ApplicationProperties.Company = "Evotec";
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.Equal("Performance Review", loaded.BuiltinDocumentProperties.Title);
            Assert.Equal("Evotec", loaded.ApplicationProperties.Company);
        }

        [Fact]
        public void PerformanceReview_StreamFastPackageFallback_PreservesRowMetadata() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Hidden");
                var row = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First();
                row.Hidden = true;
                row.Height = 24D;
                row.CustomHeight = true;
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var savedRow = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First();
            Assert.True(savedRow.Hidden!.Value);
            Assert.Equal(24D, savedRow.Height!.Value);
            Assert.True(savedRow.CustomHeight!.Value);
        }

        [Fact]
        public void PerformanceReview_ReadRangeFormulaWithoutCachedValue_ReturnsFormulaText() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.FormulaWithoutCache.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Formula");
                sheet.CellFormula(1, 1, "SUM(1,2)");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference?.Value == "A1");
                cell.CellValue = null;
                worksheetPart.Worksheet.Save();
            }

            using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions { UseCachedFormulaResult = true });
            object?[,] values = reader.GetSheet("Formula").ReadRange("A1:A1", ExecutionMode.Sequential);

            Assert.Equal("SUM(1,2)", values[0, 0]);
        }

        [Fact]
        public void PerformanceReview_HeaderlessTableObjectReaders_ThrowHelpfulError() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.HeaderlessTableReaders.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Alpha");
                sheet.CellValue(1, 2, 10d);
                sheet.CellValue(2, 1, "Beta");
                sheet.CellValue(2, 2, 20d);
                sheet.AddTable("A1:B2", false, "HeaderlessData", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                document.Save();
            }

            using var reader = ExcelDocumentReader.Open(filePath);
            var exception = Assert.Throws<InvalidOperationException>(() => reader.ReadTableObjects("HeaderlessData").ToList());
            Assert.Contains("requires table 'HeaderlessData' to have a header row", exception.Message);

            DataTable table = reader.ReadTableAsDataTable("HeaderlessData", headersInFirstRow: false);
            Assert.Equal(2, table.Rows.Count);
            Assert.Equal("Alpha", table.Rows[0][0]);
            Assert.Equal(20d, table.Rows[1][1]);
        }
    }
}
