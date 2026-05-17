using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
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

        [Fact]
        public void PerformanceReview_ReadObjectsSequential_FallsBackWhenRowsHaveImplicitIndex() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.ImplicitRowIndexReadObjects.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Score");
                sheet.CellValue(2, 1, "Alpha");
                sheet.CellValue(2, 2, 10d);
                document.Save();
            }

            RemoveFirstRowIndex(filePath);

            using var reader = ExcelDocumentReader.Open(filePath);
            var rows = reader.GetSheet("Data").ReadObjects("A1:B2", ExecutionMode.Sequential);

            var row = Assert.Single(rows);
            Assert.Equal("Alpha", row["Name"]);
            Assert.Equal(10d, row["Score"]);
        }

        [Fact]
        public void PerformanceReview_CellValuesAppend_FallsBackWhenExistingRowsHaveImplicitIndex() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.ImplicitRowIndexCellValues.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Existing");
                document.Save();
            }

            RemoveFirstRowIndex(filePath);

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets[0];
                var cells = Enumerable.Range(2, 20)
                    .Select(row => (row, 1, (object)("Row " + row.ToString(System.Globalization.CultureInfo.InvariantCulture))))
                    .ToList();
                sheet.CellValues(cells, ExecutionMode.Sequential);
                document.Save();
            }

            AssertWorksheetHasUniqueCellReferences(filePath);
            AssertWorksheetContainsCellReferences(filePath, "A1", "A21");
        }

        [Fact]
        public void PerformanceReview_DataTableAppend_FallsBackWhenExistingRowsHaveImplicitIndex() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.ImplicitRowIndexDataTable.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Existing");
                document.Save();
            }

            RemoveFirstRowIndex(filePath);

            var table = new DataTable();
            table.Columns.Add("Name", typeof(string));
            for (int i = 0; i < 20; i++) {
                table.Rows.Add("Row " + i.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }

            using (var document = ExcelDocument.Load(filePath)) {
                document.Sheets[0].InsertDataTable(table, startRow: 2, startColumn: 1);
                document.Save();
            }

            AssertWorksheetHasUniqueCellReferences(filePath);
            AssertWorksheetContainsCellReferences(filePath, "A1", "A22");
        }

        [Fact]
        public void PerformanceReview_ReadRangeSequential_HonorsTypedDatesAndDecimalOption() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.TypedDateAndDecimal.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "placeholder");
                sheet.CellValue(1, 2, 1d);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
                cells["A1"].DataType = CellValues.Date;
                cells["A1"].CellValue = new CellValue("2024-01-02T03:04:05");
                cells["B1"].DataType = CellValues.Number;
                cells["B1"].CellValue = new CellValue("123.45");
                spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();
            }

            using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions { NumericAsDecimal = true });
            object?[,] values = reader.GetSheet("Data").ReadRange("A1:B1", ExecutionMode.Sequential);

            var date = Assert.IsType<DateTime>(values[0, 0]);
            Assert.Equal(new DateTime(2024, 1, 2, 3, 4, 5), date);
            var number = Assert.IsType<decimal>(values[0, 1]);
            Assert.Equal(123.45m, number);
        }

        [Fact]
        public void PerformanceReview_StreamFastPackageFallback_PreservesHiddenSheetState() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                document.AddWorkSheet("Visible").CellValue(1, 1, "Visible");
                var hidden = document.AddWorkSheet("Hidden");
                hidden.CellValue(1, 1, "Hidden");
                hidden.SetHidden(true);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var hiddenSheet = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name == "Hidden");
            Assert.Equal(SheetStateValues.Hidden, hiddenSheet.State!.Value);
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_RejectsEmptyDataSet() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet();

            var exception = Assert.Throws<ArgumentException>(() => ExcelDocument.WriteDataSet(memory, dataSet));
            Assert.Contains("at least one DataTable", exception.Message);
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_SerializesNonNumericFormattableValuesAsText() {
            using var memory = new MemoryStream();
            var id = Guid.Parse("89f22c99-1d51-4de5-b3b4-b20c4a60164f");
            var dataSet = new DataSet();
            var table = new DataTable("Items");
            table.Columns.Add("Id", typeof(Guid));
            table.Rows.Add(id);
            dataSet.Tables.Add(table);

            ExcelDocument.WriteDataSet(memory, dataSet);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cell = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().Single(c => c.CellReference?.Value == "A2");
            Assert.Equal(CellValues.String, cell.DataType!.Value);
            Assert.Equal(id.ToString(), cell.CellValue!.Text);
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_RejectsOversizedTextValues() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet();
            var table = new DataTable("Items");
            table.Columns.Add("Notes", typeof(string));
            table.Rows.Add(new string('A', 32768));
            dataSet.Tables.Add(table);

            var exception = Assert.Throws<ArgumentException>(() => ExcelDocument.WriteDataSet(memory, dataSet));
            Assert.Contains("32,767", exception.Message);
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_FallsBackForOutOfRangeDateTimeOffset() {
            using var memory = new MemoryStream();
            var value = DateTimeOffset.MinValue;
            var dataSet = new DataSet();
            var table = new DataTable("Items");
            table.Columns.Add("When", typeof(DateTimeOffset));
            table.Rows.Add(value);
            dataSet.Tables.Add(table);

            ExcelDocument.WriteDataSet(memory, dataSet);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cell = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().Single(c => c.CellReference?.Value == "A2");
            Assert.Equal(CellValues.String, cell.DataType!.Value);
            Assert.Equal(value.ToString("o", CultureInfo.InvariantCulture), cell.CellValue!.Text);
        }

#if NET6_0_OR_GREATER
        [Fact]
        public void PerformanceReview_WriteDataSet_AppliesDateOnlyAndTimeOnlyStyles() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet();
            var table = new DataTable("Items");
            table.Columns.Add("Date", typeof(DateOnly));
            table.Columns.Add("Time", typeof(TimeOnly));
            table.Rows.Add(new DateOnly(2026, 5, 17), new TimeOnly(14, 30, 0));
            dataSet.Tables.Add(table);

            ExcelDocument.WriteDataSet(memory, dataSet);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
            Assert.Equal(1U, cells["A2"].StyleIndex!.Value);
            Assert.Equal(2U, cells["B2"].StyleIndex!.Value);
        }
#endif

        [Fact]
        public void PerformanceReview_CellValuesAppend_InsertsMissingDimensionAfterSheetProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.DimensionOrder.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Existing");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                worksheet.GetFirstChild<SheetDimension>()?.Remove();
                if (worksheet.GetFirstChild<SheetProperties>() == null) {
                    worksheet.PrependChild(new SheetProperties());
                }

                worksheet.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var cells = Enumerable.Range(2, 20)
                    .Select(row => (row, 1, (object)("Row " + row.ToString(CultureInfo.InvariantCulture))))
                    .ToList();
                document.Sheets[0].CellValues(cells, ExecutionMode.Sequential);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                var children = worksheet.ChildElements.ToList();
                int propertiesIndex = children.FindIndex(element => element is SheetProperties);
                int dimensionIndex = children.FindIndex(element => element is SheetDimension);
                Assert.True(propertiesIndex >= 0);
                Assert.True(dimensionIndex > propertiesIndex);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
            }
        }

        [Fact]
        public void PerformanceReview_StreamCreateClose_PersistsWorkbook() {
            using var memory = new MemoryStream();

            var document = ExcelDocument.Create(memory);
            document.AddWorkSheet("Data").CellValue(1, 1, "Closed");
            document.Close();

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? text));
            Assert.Equal("Closed", text);
        }

        private static void RemoveFirstRowIndex(string filePath) {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var row = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First();
            row.RowIndex = null;
            spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();
        }

        private static void AssertWorksheetHasUniqueCellReferences(string filePath) {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var references = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .Select(cell => cell.CellReference?.Value)
                .Where(reference => !string.IsNullOrWhiteSpace(reference))
                .ToList();

            Assert.Equal(references.Count, references.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        }

        private static void AssertWorksheetContainsCellReferences(string filePath, params string[] expectedReferences) {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var references = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .Select(cell => cell.CellReference?.Value)
                .Where(reference => !string.IsNullOrWhiteSpace(reference))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (string expectedReference in expectedReferences) {
                Assert.Contains(expectedReference, references);
            }
        }
    }
}
