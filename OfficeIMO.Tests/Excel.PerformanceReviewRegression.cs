using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
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
        public void PerformanceReview_StreamFastPackage_PreservesRowMetadata() {
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
        public void PerformanceReview_StreamFastPackage_PreservesHiddenSheetState() {
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

        [Fact]
        public void PerformanceReview_InsertDataTable_DirectPackagePreservesDateTimeOffsetFallbackThreshold() {
            using var memory = new MemoryStream();
            var value = new DateTimeOffset(1899, 12, 31, 23, 59, 0, TimeSpan.Zero);
            var table = new DataTable("Items");
            table.Columns.Add("When", typeof(DateTimeOffset));
            table.Rows.Add(value);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTable(table);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cell = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().Single(c => c.CellReference?.Value == "A2");
            Assert.Equal(CellValues.String, cell.DataType!.Value);
            Assert.Equal(value.ToString("o", CultureInfo.InvariantCulture), cell.CellValue!.Text);
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_StreamsLargeWorksheetPackageReadable() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Rows");
            table.Columns.Add("Index", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("When", typeof(DateTime));
            table.Columns.Add("Duration", typeof(TimeSpan));
            for (int i = 0; i < 5000; i++) {
                table.Rows.Add(
                    i,
                    "Row " + i.ToString(CultureInfo.InvariantCulture),
                    new DateTime(2026, 5, 18).AddMinutes(i),
                    TimeSpan.FromSeconds(i));
            }

            dataSet.Tables.Add(table);

            ExcelDocument.WriteDataSet(memory, dataSet);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("4999", cells["A5001"].CellValue!.Text);
            Assert.Equal(CellValues.String, cells["B5001"].DataType!.Value);
            Assert.Equal("Row 4999", cells["B5001"].CellValue!.Text);
            Assert.Equal(1U, cells["C5001"].StyleIndex!.Value);
            Assert.Equal(2U, cells["D5001"].StyleIndex!.Value);
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_ExplicitSaveUsesDirectDataSetPackageWhenUnchanged() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetDirectSave.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);
            table.Rows.Add("Beta", 20);
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                var results = document.InsertDataSet(dataSet);

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
                var result = Assert.Single(results);
                Assert.Equal("Items", result.SheetName);
                Assert.Equal("A1:B3", result.Range);
            }

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            Assert.NotNull(worksheet.Descendants<Cell>().FirstOrDefault(cell => cell.CellReference?.Value == "B3"));
            Assert.NotNull(spreadsheet.WorkbookPart.WorksheetParts.First().TableDefinitionParts.SingleOrDefault());
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_DirectPackageIncludesStandardDocumentProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetDirectSaveDocumentProperties.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Alpha");
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet);

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using var file = File.OpenRead(filePath);
            using var archive = new ZipArchive(file, ZipArchiveMode.Read);
            Assert.NotNull(archive.GetEntry("docProps/core.xml"));
            Assert.NotNull(archive.GetEntry("docProps/app.xml"));
            string contentTypes = ReadZipEntry(archive, "[Content_Types].xml");
            string packageRelationships = ReadZipEntry(archive, "_rels/.rels");
            Assert.Contains("/docProps/core.xml", contentTypes);
            Assert.Contains("/docProps/app.xml", contentTypes);
            Assert.Contains("metadata/core-properties", packageRelationships);
            Assert.Contains("extended-properties", packageRelationships);
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_NamedRangeAfterDeferredImportSkipsDirectPackageAndPersists() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetNamedRangeFallback.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Alpha");
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet);
                document.SetNamedRange("ExportNames", "'Items'!A1:A2", save: false);

                document.Save();

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var definedName = spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "ExportNames");
            Assert.Equal("'Items'!$A$1:$A$2", definedName.Text);
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_CalculationPolicySkipsDirectPackage() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetCalculationFallback.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Value", typeof(int));
            table.Rows.Add(1);
            dataSet.Tables.Add(table);

            using var document = ExcelDocument.Create(filePath);
            document.InsertDataSet(dataSet);
            document.Calculation.ForceFullCalculationOnOpen = true;

            document.Save();

            Assert.Equal(ExcelSavePackageWriter.StandardPackage, document.LastSaveDiagnostics.Writer);
            Assert.False(document.LastSaveDiagnostics.UsedFastPackageWriter);
            Assert.Contains("Calculation", document.LastSaveDiagnostics.FastPackageSkipReason, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task PerformanceReview_InsertDataSet_AsyncDirectFileSaveHonorsCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetCancelledSave.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Value", typeof(int));
            table.Rows.Add(1);
            dataSet.Tables.Add(table);

            using var document = ExcelDocument.Create(filePath);
            document.InsertDataSet(dataSet);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                document.SaveAsync(filePath, openExcel: false, options: null, cancellationToken: cancellation.Token));
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_WorkbookProtectionBeforeDeferredImportSkipsDirectPackageAndPersists() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetProtectedWorkbookFallback.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Value", typeof(int));
            table.Rows.Add(1);
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.ProtectWorkbook();
                document.InsertDataSet(dataSet);

                document.Save();

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var protection = spreadsheet.WorkbookPart!.Workbook.GetFirstChild<WorkbookProtection>();
            Assert.NotNull(protection);
            Assert.True(protection!.LockStructure?.Value ?? false);
        }

        [Fact]
        public async Task PerformanceReview_SimplePackageAsyncFileSaveHonorsCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.SimplePackageCancelledFileSave.xlsx");

            using var document = ExcelDocument.Create(filePath);
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValue(1, 1, "Value");
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                document.SaveAsync(filePath, openExcel: false, options: null, cancellationToken: cancellation.Token));
        }

        [Fact]
        public async Task PerformanceReview_SimplePackageAsyncStreamSaveHonorsCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.SimplePackageCancelledStreamSave.xlsx");

            using var document = ExcelDocument.Create(filePath);
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValue(1, 1, "Value");
            using var destination = new MemoryStream();
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                document.SaveAsync(destination, options: null, cancellationToken: cancellation.Token));
        }

        private static string ReadZipEntry(ZipArchive archive, string entryName) {
            var entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing ZIP entry '" + entryName + "'.");
            using var stream = entry.Open();
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_AutoFitUsesDirectDataSetPackageAndWritesColumnWidths() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetDirectSaveAutoFit.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Description", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("A very long value that should expand the exported column", 10);
            table.Rows.Add("Short", 20);
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet, autoFit: true);

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var columns = worksheet.GetFirstChild<Columns>();
            Assert.NotNull(columns);
            var column1 = columns!.Elements<Column>().Single(column => column.Min?.Value == 1U && column.Max?.Value == 1U);
            var column2 = columns.Elements<Column>().Single(column => column.Min?.Value == 2U && column.Max?.Value == 2U);
            Assert.True(column1.Width!.Value > column2.Width!.Value);
            Assert.True(column1.BestFit?.Value ?? false);
            Assert.True(column1.CustomWidth?.Value ?? false);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_PlainRangeAutoFitUsesDirectDataSetPackage() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetPlainRangeDirectSaveAutoFit.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("OfficeIMO automatic fast path");
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet, createTables: false, autoFit: true);

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            Assert.Empty(worksheetPart.TableDefinitionParts);
            Assert.NotNull(worksheetPart.Worksheet.GetFirstChild<Columns>());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_HeaderlessTableUsesDirectDataSetPackage() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetDirectSaveHeaderlessTable.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);
            table.Rows.Add("Beta", 20);
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                var results = document.InsertDataSet(dataSet, includeHeaders: false);

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                var result = Assert.Single(results);
                Assert.Equal("Items", result.TableName);
                Assert.Equal("A1:B2", result.Range);
            }

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", cells["A1"].CellValue!.Text);
            Assert.Equal("10", cells["B1"].CellValue!.Text);
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal(0U, tableDefinition.HeaderRowCount!.Value);
            Assert.Null(tableDefinition.AutoFilter);
            Assert.Equal("A1:B2", tableDefinition.Reference!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_SparseNullValuesUseDirectDataSetPackageWithExplicitEmptyCells() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetDirectSaveSparseNulls.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("OptionalScore", typeof(int));
            table.Columns.Add("Amount", typeof(double));
            table.Rows.Add("Alpha", DBNull.Value, 12.5d);
            table.Rows.Add("Beta", 20, DBNull.Value);
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet);

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.True(cells.ContainsKey("A2"));
            Assert.True(cells.ContainsKey("B2"));
            Assert.True(cells.ContainsKey("C2"));
            Assert.True(cells.ContainsKey("B3"));
            Assert.True(cells.ContainsKey("C3"));
            Assert.Equal(CellValues.String, cells["B2"].DataType!.Value);
            Assert.Equal(string.Empty, cells["B2"].CellValue!.Text);
            Assert.Equal(CellValues.String, cells["C3"].DataType!.Value);
            Assert.Equal(string.Empty, cells["C3"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_AllNullRowsUseDirectDataSetPackageWithExplicitEmptyRows() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetDirectSaveAllNullRows.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("OptionalScore", typeof(int));
            table.Rows.Add("Alpha", 10);
            table.Rows.Add(DBNull.Value, DBNull.Value);
            table.Rows.Add("Gamma", 30);
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet);

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var sheetData = worksheet.GetFirstChild<SheetData>()!;
            Assert.Contains(sheetData.Elements<Row>(), row => row.RowIndex?.Value == 3U);
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(CellValues.String, cells["A3"].DataType!.Value);
            Assert.Equal(string.Empty, cells["A3"].CellValue!.Text);
            Assert.Equal(CellValues.String, cells["B3"].DataType!.Value);
            Assert.Equal(string.Empty, cells["B3"].CellValue!.Text);
            var tableDefinition = spreadsheet.WorkbookPart.WorksheetParts.First().TableDefinitionParts.Single().Table!;
            Assert.Equal("A1:B4", tableDefinition.Reference!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTable_DirectPackagePreservesLargeIntegerNumberCells() {
            using var memory = new MemoryStream();
            var table = new DataTable("Numbers");
            table.Columns.Add("Value", typeof(long));
            table.Rows.Add(long.MaxValue);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTable(table);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Cell valueCell = cells["A2"];
            Assert.Equal(CellValues.Number, valueCell.DataType!.Value);
            Assert.Equal(long.MaxValue.ToString(CultureInfo.InvariantCulture), valueCell.CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValues_DirectPackagePreservesExplicitNullCells() {
            using var memory = new MemoryStream();
            var cells = new (int Row, int Column, object Value)[] {
                (1, 1, "Id"),
                (2, 1, 1),
                (3, 1, null!)
            };

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValues(cells);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var writtenCells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(CellValues.String, writtenCells["A3"].DataType!.Value);
            Assert.Equal(string.Empty, writtenCells["A3"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_DirectSaveEscapesAndSanitizesStrings() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetDirectSaveEscapedStrings.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("A&B <tag> \"quote\" 'single'");
            table.Rows.Add("Bad" + '\u0001' + "Value");
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet);

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using var loaded = ExcelDocument.Load(filePath, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? escaped));
            Assert.Equal("A&B <tag> \"quote\" 'single'", escaped);
            Assert.True(loaded.Sheets[0].TryGetCellText(3, 1, out string? sanitized));
            Assert.Equal("BadValue", sanitized);
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_DirectFileSaveKeepsDocumentEditable() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetDirectSaveEditable.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Alpha");
            table.Rows.Add("Beta");
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet);

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.Sheets[0].TryGetCellText(2, 1, out string? savedValue));
                Assert.Equal("Alpha", savedValue);

                document.Sheets[0].CellValue(4, 1, "Gamma");
                document.Save();
            }

            using var loaded = ExcelDocument.Load(filePath, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(4, 1, out string? value));
            Assert.Equal("Gamma", value);
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_SourceMutationAfterDeferredImportDoesNotChangeSavedData() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetDirectSaveSourceMutation.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Original");
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet);
                table.Rows.Add("Late");

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using var loaded = ExcelDocument.Load(filePath, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? original));
            Assert.Equal("Original", original);
            Assert.False(loaded.Sheets[0].TryGetCellText(3, 1, out _));
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_SourceSchemaMutationAfterDeferredImportDoesNotChangeSavedData() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetDirectSaveSourceSchemaMutation.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Original");
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet);
                table.Columns.Add("LateColumn", typeof(string));
                table.Rows[0]["LateColumn"] = "Late";
                dataSet.Tables.Add(new DataTable("LateTable"));

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using var loaded = ExcelDocument.Load(filePath, readOnly: true);
            Assert.Single(loaded.Sheets);
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.Equal("Name", header);
            Assert.False(loaded.Sheets[0].TryGetCellText(1, 2, out _));
            Assert.False(loaded.Sheets[0].TryGetCellText(2, 2, out _));
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_WorkbookMutationInvalidatesDirectDataSetPackageCandidate() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.InsertDataSetDirectSaveWorkbookMutation.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Original");
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet);
                document.Sheets[0].CellValue(3, 1, "Workbook");

                document.Save();

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using var loaded = ExcelDocument.Load(filePath, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(3, 1, out string? value));
            Assert.Equal("Workbook", value);
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

        [Fact]
        public void PerformanceReview_StreamLoadCopyWorksheet_PersistsInsteadOfWritingUnchangedPackage() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                document.AddWorkSheet("Source").CellValue(1, 1, "Copied");
            }

            memory.Position = 0;
            using (var document = ExcelDocument.Load(memory, autoSave: true)) {
                document.CopyWorkSheet("Source", "Copy");
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.Contains(loaded.Sheets, sheet => sheet.Name == "Copy");
        }

        [Fact]
        public void PerformanceReview_StreamFastPackage_PreservesColumnPhoneticAttribute() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Value");
                var worksheet = sheet.WorksheetPart.Worksheet;
                var sheetData = worksheet.GetFirstChild<SheetData>()!;
                var columns = new Columns(new Column {
                    Min = 1U,
                    Max = 1U,
                    Width = 12D,
                    CustomWidth = true,
                    Phonetic = true
                });
                worksheet.InsertBefore(columns, sheetData);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var column = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single();
            Assert.True(column.Phonetic!.Value);
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterWhenEligible() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(2, 1, "OfficeIMO");

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");
            Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? text));
            Assert.Equal("OfficeIMO", text);
        }

        [Fact]
        public void PerformanceReview_ExplicitFileSave_UsesSimplePackageWriterWhenEligible() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.SimplePackageExplicitSave.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "OfficeIMO");

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.SimplePackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            using (var loaded = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? text));
                Assert.Equal("OfficeIMO", text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_ReportsSimplePackageFallbackReason() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Comments");
            sheet.CellValue(1, 1, "Project");
            sheet.SetComment(1, 1, "Fallback coverage", "OfficeIMO");

            document.Save(memory);

            Assert.Equal(ExcelSavePackageWriter.StandardPackage, document.LastSaveDiagnostics.Writer);
            Assert.False(document.LastSaveDiagnostics.UsedFastPackageWriter);
            Assert.False(string.IsNullOrWhiteSpace(document.LastSaveDiagnostics.FastPackageSkipReason));
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForHyperlinks() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Links");
            sheet.SetHyperlink(1, 1, "https://github.com/EvotecIT/OfficeIMO", "OfficeIMO");
            sheet.SetInternalLink(2, 1, "'Links'!A1", "Back");

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var hyperlinks = worksheetPart.Worksheet.GetFirstChild<Hyperlinks>()!.Elements<Hyperlink>().ToList();
            Assert.Equal(2, hyperlinks.Count);
            Assert.Contains(hyperlinks, hyperlink => hyperlink.Id != null && hyperlink.Reference?.Value == "A1");
            Assert.Contains(hyperlinks, hyperlink => hyperlink.Location?.Value == "'Links'!A1" && hyperlink.Reference?.Value == "A2");
            Assert.Single(worksheetPart.HyperlinkRelationships);
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForDefinedNames() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(2, 1, "OfficeIMO");
            document.SetNamedRange("GlobalData", "'Data'!A1:A2", save: false);
            sheet.SetNamedRange("LocalData", "A2", save: false);

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var definedNames = spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().ToList();
            Assert.Contains(definedNames, name => name.Name?.Value == "GlobalData" && name.LocalSheetId == null);
            Assert.Contains(definedNames, name => name.Name?.Value == "LocalData" && name.LocalSheetId?.Value == 0U);
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForInlineStrings() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Inline");
            var sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var row = new Row { RowIndex = 1U };
            row.Append(
                new Cell {
                    CellReference = "A1",
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text("Plain inline"))
                },
                new Cell {
                    CellReference = "B1",
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(
                        new Run(new Text("Rich ")),
                        new Run(new Text("inline")))
                });
            sheetData.Append(row);

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(CellValues.InlineString, cells["A1"].DataType!.Value);
            Assert.Equal("Plain inline", cells["A1"].InlineString!.InnerText);
            Assert.Equal(CellValues.InlineString, cells["B1"].DataType!.Value);
            Assert.Equal("Rich inline", cells["B1"].InlineString!.InnerText);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForRichSharedStrings() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Shared");
            var sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            sheetData.Append(new Row(
                new Cell {
                    CellReference = "A1",
                    DataType = CellValues.SharedString,
                    CellValue = new CellValue("0")
                }) {
                    RowIndex = 1U
                });

            var workbookPart = sheet.WorksheetPart.GetParentParts().OfType<WorkbookPart>().Single();
            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringPart.SharedStringTable = new SharedStringTable(
                new SharedStringItem(
                    new Run(new Text("Rich ")),
                    new Run(new Text("shared")))) {
                    Count = 1U,
                    UniqueCount = 1U
                };

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var savedCell = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().First(cell => cell.CellReference?.Value == "A1");
            var sharedString = spreadsheet.WorkbookPart.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().Single();
            Assert.Equal(CellValues.SharedString, savedCell.DataType!.Value);
            Assert.Equal("0", savedCell.CellValue!.Text);
            Assert.Equal("Rich shared", sharedString.InnerText);
            Assert.Equal(2, sharedString.Elements<Run>().Count());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForRowMetadata() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Rows");
            sheet.CellValue(1, 1, "Hidden");
            var row = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First();
            row.Hidden = true;
            row.Height = 24D;
            row.CustomHeight = true;
            row.OutlineLevel = 1;
            row.Collapsed = true;

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var savedRow = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First();
            Assert.True(savedRow.Hidden!.Value);
            Assert.Equal(24D, savedRow.Height!.Value);
            Assert.True(savedRow.CustomHeight!.Value);
            Assert.Equal(1, savedRow.OutlineLevel!.Value);
            Assert.True(savedRow.Collapsed!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForPrintMetadata() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Report");
            sheet.CellValue(1, 1, "Report");
            sheet.SetMargins(0.25D, 0.25D, 0.5D, 0.5D, 0.3D, 0.3D);
            sheet.SetOrientation(ExcelPageOrientation.Landscape);
            sheet.SetPageSetup(fitToWidth: 1U, fitToHeight: 0U);
            sheet.SetHeaderFooter(headerCenter: "OfficeIMO", footerRight: "Page &P of &N");

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
            var margins = worksheet.GetFirstChild<PageMargins>()!;
            var setup = worksheet.GetFirstChild<PageSetup>()!;
            var headerFooter = worksheet.GetFirstChild<HeaderFooter>()!;
            Assert.Equal(0.25D, margins.Left!.Value);
            Assert.Equal(0.5D, margins.Top!.Value);
            Assert.Equal(OrientationValues.Landscape, setup.Orientation!.Value);
            Assert.Equal(1U, setup.FitToWidth!.Value);
            Assert.Equal(0U, setup.FitToHeight!.Value);
            Assert.Contains("&COfficeIMO", headerFooter.OddHeader!.Text, StringComparison.Ordinal);
            Assert.Contains("&RPage &P of &N", headerFooter.OddFooter!.Text, StringComparison.Ordinal);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForWorksheetMetadata() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Metadata");
            sheet.CellValue(1, 1, "Watched");
            var worksheet = sheet.WorksheetPart.Worksheet;
            worksheet.Append(new SheetCalculationProperties { FullCalculationOnLoad = true });
            worksheet.Append(new PhoneticProperties { FontId = 0U });
            worksheet.Append(new CellWatches(new CellWatch { CellReference = "A1" }));

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var savedWorksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
            Assert.True(savedWorksheet.GetFirstChild<SheetCalculationProperties>()!.FullCalculationOnLoad!.Value);
            Assert.Equal(0U, savedWorksheet.GetFirstChild<PhoneticProperties>()!.FontId!.Value);
            Assert.Equal("A1", savedWorksheet.GetFirstChild<CellWatches>()!.Elements<CellWatch>().Single().CellReference!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForHiddenSheets() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var visible = document.AddWorkSheet("Visible");
            visible.CellValue(1, 1, "Shown");
            var hidden = document.AddWorkSheet("Hidden");
            hidden.CellValue(1, 1, "Hidden");
            hidden.SetHidden(true);

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var sheets = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().ToList();
            Assert.Equal(2, sheets.Count);
            Assert.Null(sheets[0].State);
            Assert.Equal(SheetStateValues.Hidden, sheets[1].State!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForWorkbookMetadata() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValue(1, 1, "Value");

            var workbookPart = sheet.WorksheetPart.GetParentParts().OfType<WorkbookPart>().Single();
            var workbook = workbookPart.Workbook;
            var sheets = workbook.GetFirstChild<Sheets>()!;
            workbook.InsertBefore(new WorkbookProperties { Date1904 = true }, sheets);
            workbook.InsertBefore(new BookViews(new WorkbookView { ActiveTab = 0U, FirstSheet = 0U }), sheets);
            document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                ProtectStructure = true,
                ProtectWindows = true,
                LegacyPasswordHash = "CAFE"
            });

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var savedWorkbook = spreadsheet.WorkbookPart!.Workbook;
            var properties = savedWorkbook.GetFirstChild<WorkbookProperties>()!;
            var protection = savedWorkbook.GetFirstChild<WorkbookProtection>()!;
            var bookViews = savedWorkbook.GetFirstChild<BookViews>()!;
            Assert.True(properties.Date1904!.Value);
            Assert.True(protection.LockStructure!.Value);
            Assert.True(protection.LockWindows!.Value);
            Assert.Equal("CAFE", protection.WorkbookPassword!.Value);
            Assert.Equal(0U, bookViews.Elements<WorkbookView>().Single().ActiveTab!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_ExplicitFileSave_CanUseSimplePackageWriterAfterPriorFastSave() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.SimplePackageRepeatedSave.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "OfficeIMO");

                document.Save();
                Assert.Equal(ExcelSavePackageWriter.SimplePackage, document.LastSaveDiagnostics.Writer);

                sheet = document.Sheets[0];
                sheet.CellValue(3, 1, "Again");
                document.Save();

                Assert.True(
                    document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                    document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");
            }

            using (var loaded = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.True(loaded.Sheets[0].TryGetCellText(3, 1, out string? text));
                Assert.Equal("Again", text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForSimpleFormulas() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Calc");
            sheet.CellValue(1, 1, 2d);
            sheet.CellValue(2, 1, 3d);
            sheet.CellFormula(3, 1, "SUM(A1:A2)");
            Assert.Equal(1, document.RecalculateSupportedFormulas());

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");
            Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            Cell formulaCell = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().First(cell => cell.CellReference?.Value == "A3");
            Assert.Equal("SUM(A1:A2)", formulaCell.CellFormula!.Text);
            Assert.Equal("5", formulaCell.CellValue!.Text);
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_FallsBackWhenCalculationPolicyIsPending() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            var sheet = document.AddWorkSheet("Calc");
            sheet.CellValue(1, 1, 2d);
            sheet.CellValue(2, 1, 3d);
            sheet.CellFormula(3, 1, "SUM(A1:A2)");
            document.Calculation.ForceFullCalculationOnOpen = true;

            document.Save(memory);

            Assert.Equal(ExcelSavePackageWriter.StandardPackage, document.LastSaveDiagnostics.Writer);
            Assert.False(document.LastSaveDiagnostics.UsedFastPackageWriter);
            Assert.Contains("Calculation", document.LastSaveDiagnostics.FastPackageSkipReason, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PerformanceReview_CellValuesRectangle_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var created = new DateTime(2026, 5, 19);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Name"),
                    (1, 2, (object)"Score"),
                    (1, 3, (object)"Created"),
                    (2, 1, (object)"Alpha"),
                    (2, 2, (object)10),
                    (2, 3, (object)created),
                    (3, 1, (object)"Beta"),
                    (3, 2, (object)20),
                    (3, 3, (object)created.AddDays(1))
                });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", cells["A1"].CellValue!.Text);
            Assert.Equal("Alpha", cells["A2"].CellValue!.Text);
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Empty(worksheetPart.TableDefinitionParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValuesSingleA1_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var created = new DateTime(2026, 5, 19);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)created)
                });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(1U, cells["A1"].StyleIndex!.Value);
            Assert.Empty(worksheetPart.TableDefinitionParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValuesSingleA1_WorkbookMutationInvalidatesDirectPackageCandidate() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Original")
                });
                sheet.CellValue(2, 1, "Manual edit");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? original));
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? text));
            Assert.Equal("Original", original);
            Assert.Equal("Manual edit", text);
        }

        [Fact]
        public void PerformanceReview_CellValuesRectangle_WorkbookMutationInvalidatesDirectPackageCandidate() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Name"),
                    (1, 2, (object)"Score"),
                    (2, 1, (object)"Alpha"),
                    (2, 2, (object)10)
                });
                sheet.CellValue(5, 1, "Manual edit");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(5, 1, out string? text));
            Assert.Equal("Manual edit", text);
        }

        [Fact]
        public void PerformanceReview_CellValuesSparseRange_DoesNotUseDirectPackageCandidate() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Left"),
                    (1, 3, (object)"Right")
                });

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? left));
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 3, out string? right));
            Assert.Equal("Left", left);
            Assert.Equal("Right", right);
        }

        [Fact]
        public void PerformanceReview_CellValues_NewlineStringSkipsDirectPackageAndPreservesWrapFormatting() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Line 1\nLine 2")
                });

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cell = worksheetPart.Worksheet.Descendants<Cell>().Single(c => c.CellReference!.Value == "A1");
            Assert.NotNull(cell.StyleIndex);

            var stylesheet = spreadsheet.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
            var format = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
            Assert.True(format.Alignment!.WrapText!.Value);
            Assert.True(format.ApplyAlignment!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ExplicitSelectorsUseDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", cells["A2"].CellValue!.Text);
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ExplicitSelectorsPreserveBlankHeaders() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19))
            };

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertObjects(rows, ("", row => row.Name));

                Assert.True(sheet.TryGetCellText(1, 1, out string? header));
                Assert.Equal(string.Empty, header);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? loadedHeader));
            Assert.Equal(string.Empty, loadedHeader);
        }

        [Fact]
        public void PerformanceReview_InsertObjects_DictionaryWhitespaceHeaderPreservesHeader() {
            using var memory = new MemoryStream();
            var rows = new List<Dictionary<string, object?>> {
                new Dictionary<string, object?> {
                    [" "] = "Alpha"
                }
            };

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertObjects(rows);

                Assert.True(sheet.TryGetCellText(1, 1, out string? header));
                Assert.Equal(" ", header);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(" ", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ReflectionOverloadUsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_WorkbookMutationInvalidatesDirectPackageCandidate() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19))
            };

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.CellValue(4, 1, "Manual edit");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(4, 1, out string? text));
            Assert.Equal("Manual edit", text);
        }

        [Fact]
        public void PerformanceReview_InsertObjectsThenAddTable_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C3", hasHeader: true, name: "Object Sales", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", cells["A2"].CellValue!.Text);
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("A1:C3", tableDefinition.Reference!.Value);
            Assert.Equal("Object_Sales", tableDefinition.Name!.Value);
            Assert.Equal("TableStyleMedium4", tableDefinition.TableStyleInfo!.Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjectsThenAddTable_LaterMutationInvalidatesDirectPackageCandidate() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19))
            };

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C2", hasHeader: true, name: "ObjectSales", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.CellValue(4, 1, "Manual edit");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Manual edit", GetSpreadsheetCellText(spreadsheet, cells["A4"]));
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("A1:C2", tableDefinition.Reference!.Value);
            Assert.Equal("ObjectSales", tableDefinition.Name!.Value);
        }

        [Fact]
        public void PerformanceReview_FluentRowsFrom_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                document.AsFluent()
                    .Sheet("Data", sheet => sheet.RowsFrom(rows))
                    .End()
                    .Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", cells["A1"].CellValue!.Text);
            Assert.Equal("Alpha", cells["A2"].CellValue!.Text);
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Empty(worksheetPart.TableDefinitionParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_FluentRowsFromThenTable_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                document.AsFluent()
                    .Sheet("Data", sheet => sheet
                        .RowsFrom(rows)
                        .Table("Fluent Rows", table => table.Style(OfficeIMO.Excel.TableStyle.TableStyleMedium5)))
                    .End()
                    .Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("A1:C3", tableDefinition.Reference!.Value);
            Assert.Equal("Fluent_Rows", tableDefinition.Name!.Value);
            Assert.Equal("TableStyleMedium5", tableDefinition.TableStyleInfo!.Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_FluentRowsFrom_WorkbookMutationInvalidatesDirectPackageCandidate() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19))
            };

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                document.AsFluent()
                    .Sheet("Data", sheet => sheet.RowsFrom(rows))
                    .End();
                document.Sheets[0].CellValue(5, 1, "Manual edit");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(5, 1, out string? text));
            Assert.Equal("Manual edit", text);
        }

        [Fact]
        public void PerformanceReview_InsertDataTable_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Columns.Add("Created", typeof(DateTime));
            table.Rows.Add("Alpha", 10, new DateTime(2026, 5, 19));
            table.Rows.Add("Beta", 20, new DateTime(2026, 5, 20));

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTable(table);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", cells["A1"].CellValue!.Text);
            Assert.Equal("Alpha", cells["A2"].CellValue!.Text);
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Empty(worksheetPart.TableDefinitionParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTable_ObjectColumnDateAndTimeValuesUseValueStyles() {
            using var memory = new MemoryStream();
            var table = new DataTable("Mixed");
            table.Columns.Add("Kind", typeof(string));
            table.Columns.Add("Value", typeof(object));
            table.Rows.Add("When", new DateTime(2026, 5, 19, 8, 30, 0));
            table.Rows.Add("Duration", TimeSpan.FromMinutes(95));

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTable(table);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(1U, cells["B2"].StyleIndex!.Value);
            Assert.Equal(2U, cells["B3"].StyleIndex!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTable_ObjectColumnDateAndTimeValuesKeepStylesAfterFallback() {
            using var memory = new MemoryStream();
            var table = new DataTable("Mixed");
            table.Columns.Add("Kind", typeof(string));
            table.Columns.Add("Value", typeof(object));
            table.Rows.Add("When", new DateTime(2026, 5, 19, 8, 30, 0));
            table.Rows.Add("Duration", TimeSpan.FromMinutes(95));

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTable(table);
                sheet.CellValue(5, 1, "Manual edit");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("yyyy-mm-dd hh:mm", GetCellNumberFormatCode(spreadsheet, cells["B2"]));
            Assert.Equal("[h]:mm:ss", GetCellNumberFormatCode(spreadsheet, cells["B3"]));
            Assert.Equal("Manual edit", GetSpreadsheetCellText(spreadsheet, cells["A5"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTable_HeaderlessUsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTable(table, includeHeaders: false);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("10", cells["B1"].CellValue!.Text);
            Assert.Empty(worksheetPart.TableDefinitionParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTableHeaderlessThenAddTable_UsesGeneratedColumnNames() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTable(table, includeHeaders: false);
                sheet.AddTable("A1:B1", hasHeader: false, name: "HeaderlessSales", style: OfficeIMO.Excel.TableStyle.TableStyleMedium9, includeAutoFilter: true);

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("10", cells["B1"].CellValue!.Text);
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("0", tableDefinition.HeaderRowCount!.Value.ToString(CultureInfo.InvariantCulture));
            var columns = tableDefinition.TableColumns!.Elements<TableColumn>().ToList();
            Assert.Equal("Column1", columns[0].Name!.Value);
            Assert.Equal("Column2", columns[1].Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTable_SourceMutationAfterInsertDoesNotChangeDirectCandidate() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTable(table);
                table.Rows[0]["Name"] = "Changed";
                table.Rows.Add("Late", 20);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", cells["A2"].CellValue!.Text);
            Assert.False(cells.ContainsKey("A3"));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTable_WorkbookMutationInvalidatesDirectPackageCandidate() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTable(table);
                sheet.CellValue(5, 1, "Manual edit");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(5, 1, out string? text));
            Assert.Equal("Manual edit", text);
        }

        [Fact]
        public void PerformanceReview_InsertDataTable_HiddenSheetSkipsDirectPackageAndPreservesState() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Alpha");

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetHidden(true);
                sheet.InsertDataTable(table);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var workbookSheet = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single();
            Assert.Equal(SheetStateValues.Hidden, workbookSheet.State!.Value);
            var cells = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTableAsTable_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Columns.Add("Created", typeof(DateTime));
            table.Rows.Add("Alpha", 10, new DateTime(2026, 5, 19));
            table.Rows.Add("Beta", 20, new DateTime(2026, 5, 20));

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                Assert.Equal("A1:C3", sheet.InsertDataTableAsTable(table, tableName: "Sales Table", style: OfficeIMO.Excel.TableStyle.TableStyleMedium9));

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", cells["A2"].CellValue!.Text);
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("A1:C3", tableDefinition.Reference!.Value);
            Assert.Equal("Sales_Table", tableDefinition.Name!.Value);
            Assert.Equal("TableStyleMedium9", tableDefinition.TableStyleInfo!.Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTableAsTable_HeaderlessTableUsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                Assert.Equal("A1:B1", sheet.InsertDataTableAsTable(table, includeHeaders: false, tableName: "HeaderlessSales"));

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", cells["A1"].CellValue!.Text);
            Assert.Equal("10", cells["B1"].CellValue!.Text);
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("0", tableDefinition.HeaderRowCount!.Value.ToString(CultureInfo.InvariantCulture));
            var columns = tableDefinition.TableColumns!.Elements<TableColumn>().ToList();
            Assert.Equal("Column1", columns[0].Name!.Value);
            Assert.Equal("Column2", columns[1].Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTableAsTable_WorkbookMutationInvalidatesDirectPackageCandidate() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");
                sheet.CellValue(4, 1, "Manual edit");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Manual edit", GetSpreadsheetCellText(spreadsheet, cells["A4"]));
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("A1:B2", tableDefinition.Reference!.Value);
            Assert.Equal("SalesTable", tableDefinition.Name!.Value);
        }

        [Fact]
        public void PerformanceReview_InsertDataTableAsTable_PackagePropertiesSkipDirectPackageAndPersist() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Alpha");

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                document.BuiltinDocumentProperties.Title = "Sales Export";
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.Equal("Sales Export", loaded.BuiltinDocumentProperties.Title);
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? text));
            Assert.Equal("Alpha", text);
        }

        [Fact]
        public void PerformanceReview_InsertDataReader_KeepsStreamingWriterWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Columns.Add("Created", typeof(DateTime));
            table.Rows.Add("Alpha", 10, new DateTime(2026, 5, 19));
            table.Rows.Add("Beta", 20, new DateTime(2026, 5, 20));

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                using IDataReader reader = table.CreateDataReader();
                Assert.Equal("A1:C3", sheet.InsertDataReader(reader, tableName: "Reader Table", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4));

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.True(cells["C2"].StyleIndex?.Value > 0U);
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("A1:C3", tableDefinition.Reference!.Value);
            Assert.Equal("Reader_Table", tableDefinition.Name!.Value);
            Assert.Equal("TableStyleMedium4", tableDefinition.TableStyleInfo!.Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataReader_AutoFitKeepsStreamingWriterWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Description", typeof(string));
            table.Rows.Add("Alpha", "A longer value for sizing");

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                using IDataReader reader = table.CreateDataReader();
                sheet.InsertDataReader(reader, tableName: "ReaderAutoFit", autoFit: true);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var columns = worksheet.GetFirstChild<Columns>();
            Assert.NotNull(columns);
            Assert.Contains(columns!.Elements<Column>(), column => column.Width?.Value > 0D);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataReader_WorkbookMutationInvalidatesDirectPackageCandidate() {
            using var memory = new MemoryStream();
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                using IDataReader reader = table.CreateDataReader();
                sheet.InsertDataReader(reader, tableName: "ReaderTable");
                sheet.CellValue(5, 1, "Manual edit");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, readOnly: true);
            Assert.True(loaded.Sheets[0].TryGetCellText(5, 1, out string? text));
            Assert.Equal("Manual edit", text);
        }

        [Fact]
        public void PerformanceReview_InsertDataReader_HeaderlessTableKeepsStreamingWriterWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                using IDataReader reader = table.CreateDataReader();
                Assert.Equal("A1:B1", sheet.InsertDataReader(reader, includeHeaders: false, tableName: "HeaderlessReader"));

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("10", cells["B1"].CellValue!.Text);
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("0", tableDefinition.HeaderRowCount!.Value.ToString(CultureInfo.InvariantCulture));
            var columns = tableDefinition.TableColumns!.Elements<TableColumn>().ToList();
            Assert.Equal("Column1", columns[0].Name!.Value);
            Assert.Equal("Column2", columns[1].Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_ObjectColumnDateAndTimeValuesUseValueStyles() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet();
            var table = new DataTable("Mixed");
            table.Columns.Add("Kind", typeof(string));
            table.Columns.Add("Value", typeof(object));
            table.Rows.Add("When", new DateTime(2026, 5, 19, 8, 30, 0));
            table.Rows.Add("Duration", TimeSpan.FromMinutes(95));
            dataSet.Tables.Add(table);

            ExcelDocument.WriteDataSet(memory, dataSet);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(1U, cells["B2"].StyleIndex!.Value);
            Assert.Equal(2U, cells["B3"].StyleIndex!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_RejectsTooManyColumns() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet();
            var table = new DataTable("TooWide");
            for (int i = 0; i <= A1.MaxColumns; i++) {
                table.Columns.Add("Column" + i.ToString(CultureInfo.InvariantCulture));
            }

            dataSet.Tables.Add(table);

            var exception = Assert.Throws<ArgumentException>(() => ExcelDocument.WriteDataSet(memory, dataSet));
            Assert.Contains(A1.MaxColumns.ToString(CultureInfo.InvariantCulture), exception.Message);
        }

        private sealed class PerformanceObjectExportRow {
            public PerformanceObjectExportRow(string name, int score, DateTime created) {
                Name = name;
                Score = score;
                Created = created;
            }

            public string Name { get; }

            public int Score { get; }

            public DateTime Created { get; }
        }

        private static void RemoveFirstRowIndex(string filePath) {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var row = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First();
            row.RowIndex = null;
            spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet!.Save();
        }

        private static void AssertWorksheetHasUniqueCellReferences(string filePath) {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var references = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>()
                .Select(cell => cell.CellReference?.Value)
                .Where(reference => !string.IsNullOrWhiteSpace(reference))
                .ToList();

            Assert.Equal(references.Count, references.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        }

        private static void AssertWorksheetContainsCellReferences(string filePath, params string[] expectedReferences) {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var references = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>()
                .Select(cell => cell.CellReference?.Value)
                .Where(reference => !string.IsNullOrWhiteSpace(reference))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (string expectedReference in expectedReferences) {
                Assert.Contains(expectedReference, references);
            }
        }

        private static string? GetCellNumberFormatCode(SpreadsheetDocument spreadsheet, Cell cell) {
            if (cell.StyleIndex == null) {
                return null;
            }

            var stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            var cellFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex.Value);
            if (cellFormat.NumberFormatId == null) {
                return null;
            }

            uint numberFormatId = cellFormat.NumberFormatId.Value;
            return stylesheet.NumberingFormats?.Elements<NumberingFormat>()
                .FirstOrDefault(format => format.NumberFormatId?.Value == numberFormatId)
                ?.FormatCode
                ?.Value;
        }

        private static string? GetSpreadsheetCellText(SpreadsheetDocument spreadsheet, Cell cell) {
            string? value = cell.CellValue?.Text;
            if (cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sharedStringIndex)) {
                return spreadsheet.WorkbookPart?.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sharedStringIndex)?.InnerText;
            }

            return value;
        }
    }
}
