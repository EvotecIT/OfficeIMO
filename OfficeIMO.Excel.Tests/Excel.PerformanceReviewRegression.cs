using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void PerformanceReview_InvariantNumberTextExpandsForMediumIndexes() {
            string text = InvariantNumberText.Get(25000);
            double doubleValue = 1d / 7d;
            float floatValue = 1f / 7f;

            Assert.Equal("25000", text);
            Assert.Same(text, InvariantNumberText.Get(25000));
            Assert.True(InvariantNumberText.TryGet(25000, out string cached));
            Assert.Same(text, cached);
            Assert.Equal("-1", InvariantNumberText.Get(-1));
            Assert.Equal(doubleValue, double.Parse(InvariantNumberText.Get(doubleValue), CultureInfo.InvariantCulture));
            Assert.Equal(floatValue, float.Parse(InvariantNumberText.Get(floatValue), CultureInfo.InvariantCulture));
            Assert.Equal("2.35", InvariantNumberText.Get(2.35d));
        }

        [Fact]
        public void PerformanceReview_SheetBatchWritesCellsWithSinglePublicCall() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Batch");

                sheet.Batch(s => {
                    for (int row = 1; row <= 100; row++) {
                        s.CellValue(row, 1, "Item " + (row % 10).ToString(CultureInfo.InvariantCulture));
                        s.CellValue(row, 2, row);
                        s.CellValue(row, 3, row % 2 == 0);
                    }
                });

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("Item 0", GetSpreadsheetCellText(spreadsheet, cells["A100"]));
            Assert.Equal("100", cells["B100"].CellValue!.Text);
            Assert.Equal(CellValues.Boolean, cells["C100"].DataType!.Value);
            Assert.Equal("1", cells["C100"].CellValue!.Text);
        }

        [Fact]
        public void PerformanceReview_SheetBatchRejectsNullAction() {
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Batch");

            Assert.Throws<ArgumentNullException>(() => sheet.Batch(null!));
        }

        [Fact]
        public void PerformanceReview_SheetBatchReadOnlyActionDoesNotDirtyLoadedWorkbook() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Batch");
                sheet.CellValue(1, 1, "Status");
                document.Save(memory);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory);
            var loadedSheet = loaded.Sheets[0];

            Assert.False(loaded.IsPackageDirty);
            Assert.False(loadedSheet.RequiresSavePreparation);

            loadedSheet.Batch(s => {
                Assert.True(s.TryGetCellText(1, 1, out string? text));
                Assert.Equal("Status", text);
            });

            Assert.False(loaded.IsPackageDirty);
            Assert.False(loadedSheet.RequiresSavePreparation);
        }

        [Fact]
        public void PerformanceReview_SheetBatchAllowsNestedWriteLockOperations() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Batch");

                sheet.Batch(s => {
                    s.CellValue(1, 1, "Name");
                    s.CellValue(1, 2, "Score");
                    s.CellValue(2, 1, "Alpha");
                    s.CellValue(2, 2, 42);
                    s.AddTable("A1:B2", hasHeader: true, name: "BatchTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                });

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);

            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            var table = spreadsheet.WorkbookPart!.WorksheetParts.First().TableDefinitionParts.Single().Table!;
            Assert.Equal("BatchTable", table.Name!.Value);
            Assert.Equal("A1:B2", table.Reference!.Value);
        }

        [Fact]
        public void PerformanceReview_SheetBatchHeaderMutationRefreshesCachedHeadersInsideBatch() {
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Batch");
            sheet.CellValue(1, 1, "Status");
            Assert.True(sheet.TryGetColumnIndexByHeader("Status", out int statusColumn));
            Assert.Equal(1, statusColumn);

            sheet.Batch(s => {
                s.CellValue(1, 1, "State");

                Assert.True(s.TryGetColumnIndexByHeader("State", out int stateColumn));
                Assert.Equal(1, stateColumn);
                Assert.False(s.TryGetColumnIndexByHeader("Status", out _));
            });
        }

        [Fact]
        public void PerformanceReview_CellValuesEmptyStringsUseDirectStringCells() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Empty");
                sheet.CellValues(new[] {
                    (1, 1, (object)string.Empty),
                    (1, 2, (object)"Header"),
                    (2, 1, (object)string.Empty)
                });

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal(CellValues.String, cells["A1"].DataType!.Value);
            Assert.Equal(string.Empty, cells["A1"].CellValue!.Text);
            Assert.Equal(CellValues.String, cells["A2"].DataType!.Value);
            Assert.Equal(string.Empty, cells["A2"].CellValue!.Text);
            Assert.Equal(CellValues.SharedString, cells["B1"].DataType!.Value);
            Assert.Single(spreadsheet.WorkbookPart.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>());
        }

        [Fact]
        public void PerformanceReview_CellValuesReadOnlyListUsesDirectPackageWithoutSnapshotEnumeration() {
            var values = new ThrowOnEnumerateReadOnlyList<(int Row, int Column, object Value)>(
                (1, 1, (object)"Id"),
                (1, 2, (object)"Name"),
                (2, 1, (object)1),
                (2, 2, (object)"Alpha"));

            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(values);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("1", cells["A2"].CellValue!.Text);
        }

        [Fact]
        public void PerformanceReview_LoadedWorkbookDirectAppend_PersistsAfterSave() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.DirectAppendDirty.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Data");
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.True(document.Sheets[0].TryGetCellText(21, 1, out string? text));
                Assert.Equal("Row 21", text);
            }
        }

        [Fact]
        public void PerformanceReview_LoadedWorkbookFastCellValueOverloadsPersistAfterSave() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.FastCellValueOverloadsDirty.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Seed");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets[0];
                sheet.CellValue(2, 1, "Text");
                sheet.CellValue(3, 1, 12.5d);
                sheet.CellValue(4, 1, 12.5m);
                sheet.CellValue(5, 1, new DateTime(2026, 5, 20));
                sheet.CellValue(6, 1, TimeSpan.FromHours(2));
                sheet.CellValue(7, 1, true);
                sheet.CellFormula(8, 1, "SUM(A3:A4)");
                sheet.CellValue(9, 1, (object)"Object");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

                Assert.True(cells.ContainsKey("A2"));
                Assert.True(cells.ContainsKey("A3"));
                Assert.True(cells.ContainsKey("A4"));
                Assert.True(cells.ContainsKey("A5"));
                Assert.True(cells.ContainsKey("A6"));
                Assert.True(cells.ContainsKey("A7"));
                Assert.Equal("SUM(A3:A4)", cells["A8"].CellFormula!.Text);
                Assert.True(cells.ContainsKey("A9"));
            }
        }

        [Fact]
        public void PerformanceReview_LoadedWorkbookProtection_PersistsAfterSave() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.WorkbookProtectionDirty.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
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

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                document.BuiltinDocumentProperties.Title = "Performance Review";
                document.ApplicationProperties.Company = "Evotec";
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.Equal("Performance Review", loaded.BuiltinDocumentProperties.Title);
            Assert.Equal("Evotec", loaded.ApplicationProperties.Company);
        }

        [Fact]
        public void PerformanceReview_StreamFastPackage_PreservesRowMetadata() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
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
        public void PerformanceReview_StreamFastPackage_PreservesPlainFormulas() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Formulas");
                for (int row = 1; row <= 5; row++) {
                    sheet.CellValue(row, 1, (double)row);
                    sheet.CellValue(row, 2, (double)(row + 1));
                    sheet.CellValue(row, 3, (double)(row + 2));
                    sheet.CellFormula(row, 4, $"SUM(A{row}:C{row})");
                }

                document.Save(memory);

                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("SUM(A5:C5)", cells["D5"].CellFormula!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_ReadRangeFormulaWithoutCachedValue_ReturnsFormulaText() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.FormulaWithoutCache.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Formula");
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
                var sheet = document.AddWorksheet("Data");
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
                var sheet = document.AddWorksheet("Data");
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
                document.AddWorksheet("Data").CellValue(1, 1, "Existing");
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
                document.AddWorksheet("Data").CellValue(1, 1, "Existing");
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
                var sheet = document.AddWorksheet("Data");
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
                var row = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First();
                cells["C1"] = new Cell {
                    CellReference = "C1",
                    DataType = CellValues.Number,
                    CellValue = new CellValue("1.2345E+2")
                };
                cells["D1"] = new Cell {
                    CellReference = "D1",
                    DataType = CellValues.Number,
                    CellValue = new CellValue("0.000001")
                };
                cells["E1"] = new Cell {
                    CellReference = "E1",
                    DataType = CellValues.Number,
                    CellValue = new CellValue("-9876543210.1234")
                };
                row.Append(cells["C1"]);
                row.Append(cells["D1"]);
                row.Append(cells["E1"]);
                spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();
            }

            using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions { NumericAsDecimal = true });
            object?[,] values = reader.GetSheet("Data").ReadRange("A1:E1", ExecutionMode.Sequential);

            var date = Assert.IsType<DateTime>(values[0, 0]);
            Assert.Equal(new DateTime(2024, 1, 2, 3, 4, 5), date);
            var number = Assert.IsType<decimal>(values[0, 1]);
            Assert.Equal(123.45m, number);
            var exponentNumber = Assert.IsType<decimal>(values[0, 2]);
            Assert.Equal(123.45m, exponentNumber);
            var smallFraction = Assert.IsType<decimal>(values[0, 3]);
            Assert.Equal(0.000001m, smallFraction);
            var negativeNumber = Assert.IsType<decimal>(values[0, 4]);
            Assert.Equal(-9876543210.1234m, negativeNumber);
        }

        [Fact]
        public void PerformanceReview_StreamFastPackage_PreservesHiddenSheetState() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                document.AddWorksheet("Visible").CellValue(1, 1, "Visible");
                var hidden = document.AddWorksheet("Hidden");
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
            Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
            Assert.Equal(id.ToString(), GetSpreadsheetCellText(spreadsheet, cell));
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
        public void PerformanceReview_WriteDataSet_RejectsOversizedObjectStringValues() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet();
            var table = new DataTable("Items");
            table.Columns.Add("Notes", typeof(object));
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
            Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
            Assert.Equal(value.ToString("o", CultureInfo.InvariantCulture), GetSpreadsheetCellText(spreadsheet, cell));
        }

        [Fact]
        public void PerformanceReview_InsertDataTable_DirectPackagePreservesDateTimeOffsetFallbackThreshold() {
            using var memory = new MemoryStream();
            var value = new DateTimeOffset(1899, 12, 31, 23, 59, 0, TimeSpan.Zero);
            var table = new DataTable("Items");
            table.Columns.Add("When", typeof(DateTimeOffset));
            table.Rows.Add(value);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataTable(table);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cell = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().Single(c => c.CellReference?.Value == "A2");
            Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
            Assert.Equal(value.ToString("o", CultureInfo.InvariantCulture), GetSpreadsheetCellText(spreadsheet, cell));
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
            Assert.Equal(CellValues.InlineString, cells["B5001"].DataType!.Value);
            Assert.Equal("Row 4999", GetSpreadsheetCellText(spreadsheet, cells["B5001"]));
            Assert.Equal(1U, cells["C5001"].StyleIndex!.Value);
            Assert.Equal(2U, cells["D5001"].StyleIndex!.Value);
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_SmallStringExportSkipsSharedStringTable() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));

            for (int i = 0; i < 20; i++) {
                table.Rows.Add("Row " + i.ToString(CultureInfo.InvariantCulture));
            }

            dataSet.Tables.Add(table);

            ExcelDocument.WriteDataSet(memory, dataSet);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            Assert.Null(spreadsheet.WorkbookPart!.SharedStringTablePart);

            var worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(CellValues.InlineString, cells["A2"].DataType!.Value);
            Assert.Equal("Row 0", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_DuplicateStringsBuildSharedStringIndexes() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));

            for (int i = 0; i < 600; i++) {
                table.Rows.Add("Repeated");
            }

            dataSet.Tables.Add(table);

            ExcelDocument.WriteDataSet(memory, dataSet);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var sharedStrings = spreadsheet.WorkbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ToList();
            var worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal(new[] { "Name", "Repeated" }, sharedStrings.Select(item => item.InnerText).ToArray());
            Assert.Equal(CellValues.SharedString, cells["A2"].DataType!.Value);
            Assert.Equal("1", cells["A601"].CellValue!.Text);
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_SharedStringPlannerRejectsOverLimitSharedValue() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            string tooLong = new('A', 32_768);

            for (int i = 0; i < 511; i++) {
                table.Rows.Add(tooLong);
            }

            dataSet.Tables.Add(table);

            var exception = Assert.Throws<ArgumentException>(() => ExcelDocument.WriteDataSet(memory, dataSet));
            Assert.Contains("32,767", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_ForcedSharedHeaderCountsPriorDataOccurrence() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");

            var first = new DataTable("First");
            first.Columns.Add("Value", typeof(string));
            first.Rows.Add("CrossSheetHeader");
            for (int i = 0; i < 600; i++) {
                first.Rows.Add("Repeated");
            }

            var second = new DataTable("Second");
            second.Columns.Add("CrossSheetHeader", typeof(string));
            second.Rows.Add("Other");

            dataSet.Tables.Add(first);
            dataSet.Tables.Add(second);

            ExcelDocument.WriteDataSet(memory, dataSet);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var sharedStringTable = spreadsheet.WorkbookPart!.SharedStringTablePart!.SharedStringTable;
            int actualSharedCellReferences = spreadsheet.WorkbookPart.WorksheetParts
                .SelectMany(part => part.Worksheet.Descendants<Cell>())
                .Count(cell => cell.DataType?.Value == CellValues.SharedString);

            Assert.Contains(sharedStringTable.Elements<SharedStringItem>(), item => item.InnerText == "CrossSheetHeader");
            Assert.Equal((uint)actualSharedCellReferences, sharedStringTable.Count!.Value);
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_SharedStringsSkipUniqueDataValues() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Unique", typeof(string));

            for (int i = 0; i < 600; i++) {
                table.Rows.Add("Repeated", "Unique " + i.ToString(CultureInfo.InvariantCulture));
            }

            dataSet.Tables.Add(table);

            ExcelDocument.WriteDataSet(memory, dataSet);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var sharedStrings = spreadsheet.WorkbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().Select(item => item.InnerText).ToList();
            var worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal(new[] { "Name", "Unique", "Repeated" }, sharedStrings);
            Assert.Equal(CellValues.SharedString, cells["A2"].DataType!.Value);
            Assert.Equal("2", cells["A2"].CellValue!.Text);
            Assert.Equal(CellValues.InlineString, cells["B2"].DataType!.Value);
            Assert.Equal("Unique 0", GetSpreadsheetCellText(spreadsheet, cells["B2"]));
        }

        [Fact]
        public void PerformanceReview_CellValueDistinctStringsPromotesNewValuesToPlainStringsAfterSharedStringThreshold() {
            using var memory = new MemoryStream();
            int rowCount = ExcelSheet.CellValuePlainStringPromotionSharedStringCount + 3;

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Distinct");
                for (int row = 1; row <= rowCount; row++) {
                    sheet.CellValue(row, 1, "Distinct " + row.ToString(CultureInfo.InvariantCulture));
                    if (row == 1) {
                        sheet.MaterializePendingDirectCellValues();
                    }
                }

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.SimplePackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
            var sharedStrings = spreadsheet.WorkbookPart.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>().ToList();

            Assert.Equal(ExcelSheet.CellValuePlainStringPromotionSharedStringCount, sharedStrings.Count);
            Assert.Equal(CellValues.SharedString, cells["A1"].DataType!.Value);
            Assert.Equal("Distinct 1", sharedStrings[int.Parse(cells["A1"].CellValue!.Text, CultureInfo.InvariantCulture)].InnerText);

            string lastReference = "A" + rowCount.ToString(CultureInfo.InvariantCulture);
            Assert.Equal(CellValues.String, cells[lastReference].DataType!.Value);
            Assert.Equal("Distinct " + rowCount.ToString(CultureInfo.InvariantCulture), cells[lastReference].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_DeferredObjectExportMaterializesExistingDirectCellValuesBeforeFallback() {
            using var memory = new MemoryStream();
            var seedCells = Enumerable.Range(1, 160)
                .Select(row => (Row: row, Column: 1, Value: (object)("seed-" + row.ToString(CultureInfo.InvariantCulture))))
                .ToList();
            var rows = new[] {
                new { Name = "Alpha" }
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(seedCells);
                sheet.InsertObjects(rows);

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("seed-160", GetSpreadsheetCellText(spreadsheet, cells["A160"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_WriteDataSet_OmitsSparseBlankCellsButPreservesEmptyStrings() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Sparse");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("OptionalCode", typeof(string));
            table.Columns.Add("ExplicitEmpty", typeof(string));
            table.Columns.Add("ReviewDate", typeof(DateTime));
            table.Rows.Add(1, DBNull.Value, string.Empty, new DateTime(2026, 5, 20));
            table.Rows.Add(2, "C2", DBNull.Value, DBNull.Value);
            dataSet.Tables.Add(table);

            ExcelDocument.WriteDataSet(memory, dataSet);

            memory.Position = 0;
            using (var spreadsheet = SpreadsheetDocument.Open(memory, false)) {
                var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
                var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                Assert.False(cells.ContainsKey("B2"));
                Assert.True(cells.ContainsKey("C2"));
                Assert.Equal(CellValues.InlineString, cells["C2"].DataType!.Value);
                Assert.Equal(string.Empty, GetSpreadsheetCellText(spreadsheet, cells["C2"]));
                Assert.True(cells.ContainsKey("D2"));
                Assert.True(cells.ContainsKey("A3"));
                Assert.True(cells.ContainsKey("B3"));
                Assert.False(cells.ContainsKey("C3"));
                Assert.False(cells.ContainsKey("D3"));
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            DataTable imported = reader.ReadTableAsDataTable("Sparse");
            Assert.Equal(2, imported.Rows.Count);
            Assert.Equal(DBNull.Value, imported.Rows[0]["OptionalCode"]);
            Assert.Equal(string.Empty, imported.Rows[0]["ExplicitEmpty"]);
            Assert.Equal(DBNull.Value, imported.Rows[1]["ExplicitEmpty"]);
            Assert.Equal(DBNull.Value, imported.Rows[1]["ReviewDate"]);
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
                document.SaveAsync(filePath, null, cancellation.Token));
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
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Value");
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                document.SaveAsync(filePath, null, cancellation.Token));
        }

        [Fact]
        public async Task PerformanceReview_SimplePackageAsyncStreamSaveHonorsCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.SimplePackageCancelledStreamSave.xlsx");

            using var document = ExcelDocument.Create(filePath);
            var sheet = document.AddWorksheet("Data");
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
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataTable(table);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Cell valueCell = cells["A2"];
            Assert.True(valueCell.DataType == null || valueCell.DataType.Value == CellValues.Number);
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(cells);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using (var archive = new ZipArchive(memory, ZipArchiveMode.Read, leaveOpen: true)) {
                var worksheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
                using var reader = new StreamReader(worksheetEntry.Open());
                string worksheetXml = reader.ReadToEnd();
                Assert.Contains("<c r=\"A3\" t=\"str\"><v/></c>", worksheetXml, StringComparison.Ordinal);
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

            using var loaded = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
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

            using var loaded = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
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

            using var loaded = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
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

            using var loaded = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.Single(loaded.Sheets);
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.Equal("Name", header);
            Assert.False(loaded.Sheets[0].TryGetCellText(1, 2, out _));
            Assert.False(loaded.Sheets[0].TryGetCellText(2, 2, out _));
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_DisposedCandidateUnsubscribesRemovedSourceTable() {
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Original");
            dataSet.Tables.Add(table);

            using var document = ExcelDocument.Create(new MemoryStream());
            typeof(ExcelDocument).GetMethod("RegisterDirectDataSetSaveCandidate", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(
                document,
                new object[] {
                    dataSet,
                    true,
                    OfficeIMO.Excel.TableStyle.TableStyleMedium2,
                    true,
                    true,
                    false,
                    Array.Empty<ExcelDataSetImportResult>()
                });

            var candidateField = typeof(ExcelDocument).GetField("_directDataSetSaveCandidate", BindingFlags.Instance | BindingFlags.NonPublic)!;
            object? candidate = candidateField.GetValue(document);
            Assert.NotNull(candidate);
            var subscribeTableMethod = candidate.GetType()
                .GetMethods(BindingFlags.Instance | BindingFlags.NonPublic)
                .Single(method => method.Name == "Subscribe"
                    && method.GetParameters() is { Length: 1 } parameters
                    && parameters[0].ParameterType == typeof(DataTable));
            subscribeTableMethod.Invoke(candidate, new object[] { table });
            var subscribedTablesField = candidate.GetType().GetField("_subscribedTables", BindingFlags.Instance | BindingFlags.NonPublic)!;
            var subscribedTables = Assert.IsAssignableFrom<ICollection<DataTable>>(subscribedTablesField.GetValue(candidate));
            Assert.Contains(table, subscribedTables);

            dataSet.Tables.Remove(table);

            ((IDisposable)candidate).Dispose();

            Assert.Empty(subscribedTables);
            table.Rows.Add("Late");
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

            using var loaded = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
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
                document.AddWorksheet("Data").CellValue(1, 1, "Existing");
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
        public void PerformanceReview_StreamCreateDispose_PersistsWorkbook() {
            using var memory = new MemoryStream();

            var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose });
            document.AddWorksheet("Data").CellValue(1, 1, "Closed");
            document.Dispose();

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? text));
            Assert.Equal("Closed", text);
        }

        [Fact]
        public void PerformanceReview_StreamLoadCopyWorksheet_PersistsInsteadOfWritingUnchangedPackage() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                document.AddWorksheet("Source").CellValue(1, 1, "Copied");
            }

            memory.Position = 0;
            using (var document = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                document.CopyWorksheet("Source", "Copy");
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.Contains(loaded.Sheets, sheet => sheet.Name == "Copy");
        }

        [Fact]
        public void PerformanceReview_StreamFastPackage_PreservesColumnPhoneticAttribute() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
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
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(2, 1, "OfficeIMO");

            document.Save(memory);

            Assert.True(
                document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.FastPackageSkipReason ?? "Simple package writer was not used.");
            Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? text));
            Assert.Equal("OfficeIMO", text);
        }

        [Fact]
        public void PerformanceReview_ExplicitFileSave_UsesSimplePackageWriterWhenEligible() {
            string filePath = Path.Combine(_directoryWithFiles, "PerformanceReview.SimplePackageExplicitSave.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "OfficeIMO");

                document.Save();

                Assert.Equal(ExcelSavePackageWriter.SimplePackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            using (var loaded = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? text));
                Assert.Equal("OfficeIMO", text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_ReportsSimplePackageFallbackReason() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Comments");
            sheet.CellValue(1, 1, "Project");
            sheet.SetComment(1, 1, "Fallback coverage", "OfficeIMO");

            document.Save(memory);

            Assert.Equal(ExcelSavePackageWriter.StandardPackage, document.LastSaveDiagnostics.Writer);
            Assert.False(document.LastSaveDiagnostics.UsedFastPackageWriter);
            Assert.False(string.IsNullOrWhiteSpace(document.LastSaveDiagnostics.FastPackageSkipReason));
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_FallsBackForUnknownSheetDataChild() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Unknown");
            sheet.CellValue(1, 1, "Project");

            var sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var row = sheetData.Elements<Row>().First();
            row.AppendChild(new OpenXmlUnknownElement("x", "unknown", "urn:officeimo:test"));

            document.Save(memory);

            Assert.Equal(ExcelSavePackageWriter.StandardPackage, document.LastSaveDiagnostics.Writer);
            Assert.False(document.LastSaveDiagnostics.UsedFastPackageWriter);
            Assert.Contains("unknown", document.LastSaveDiagnostics.FastPackageSkipReason, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_FallsBackForUnknownDirectSheetDataChild() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Unknown");
            sheet.CellValue(1, 1, "Project");

            var sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            sheetData.AppendChild(new OpenXmlUnknownElement("x", "unknown", "urn:officeimo:test"));

            document.Save(memory);

            Assert.Equal(ExcelSavePackageWriter.StandardPackage, document.LastSaveDiagnostics.Writer);
            Assert.False(document.LastSaveDiagnostics.UsedFastPackageWriter);
            Assert.Contains("unknown", document.LastSaveDiagnostics.FastPackageSkipReason, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForHyperlinks() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Links");
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
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Data");
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
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Inline");
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
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Shared");
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
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Rows");
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
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Report");
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
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Metadata");
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
            using var document = ExcelDocument.Create(new MemoryStream());
            var visible = document.AddWorksheet("Visible");
            visible.CellValue(1, 1, "Shown");
            var hidden = document.AddWorksheet("Hidden");
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
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Data");
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
                var sheet = document.AddWorksheet("Data");
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

            using (var loaded = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.True(loaded.Sheets[0].TryGetCellText(3, 1, out string? text));
                Assert.Equal("Again", text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_UsesSimplePackageWriterForSimpleFormulas() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Calc");
            sheet.CellValue(1, 1, 2d);
            sheet.CellValue(2, 1, 3d);
            sheet.CellFormula(3, 1, "SUM(A1:A2)");
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCachedFormulaValue(3, 1, out string? cachedValue));
            Assert.Equal("5", cachedValue);

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
        public void PerformanceReview_RecalculateMaterializesPendingDirectCellValuesBeforeFormulaScan() {
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Calc");

            for (int row = 1; row <= 130; row++) {
                sheet.CellValue(row, 1, row);
            }

            sheet.CellFormula(131, 1, "SUM(A1:A130)");

            Assert.Equal(1, document.RecalculateSupportedFormulas());
            Assert.True(sheet.TryGetCachedFormulaValue(131, 1, out string? value));
            Assert.Equal("8515", value);
        }

        [Fact]
        public void PerformanceReview_ExplicitStreamSave_FallsBackWhenCalculationPolicyIsPending() {
            using var memory = new MemoryStream();
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Calc");
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
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
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Empty(worksheetPart.TableDefinitionParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValuesRectangle_WithMultilineTextDoesNotUseDirectPackage() {
            AssertCellValuesMultilineFallback(new[] {
                (1, 1, (object)"Name\nWrapped"),
                (1, 2, (object)"Score"),
                (2, 1, (object)"Alpha"),
                (2, 2, (object)10)
            }, "A1", "Name\nWrapped");

            AssertCellValuesMultilineFallback(new[] {
                (1, 1, (object)"Name"),
                (1, 2, (object)"Notes"),
                (2, 1, (object)"Alpha"),
                (2, 2, (object)"Line one\nLine two")
            }, "B2", "Line one\nLine two");

            static void AssertCellValuesMultilineFallback((int Row, int Column, object Value)[] cells, string reference, string expectedText) {
                using var memory = new MemoryStream();
                using (var document = ExcelDocument.Create(new MemoryStream())) {
                    var sheet = document.AddWorksheet("Data");
                    sheet.CellValues(cells);
                    document.Save(memory);

                    Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                }

                memory.Position = 0;
                using var spreadsheet = SpreadsheetDocument.Open(memory, false);
                var savedCells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
                    .ToDictionary(cell => cell.CellReference!.Value!);
                Assert.Equal(expectedText, GetSpreadsheetCellText(spreadsheet, savedCells[reference]));
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
            }
        }

        [Fact]
        public void PerformanceReview_CellValuesRectangleParallel_UsesDirectPackageWithSharedStrings() {
            using var memory = new MemoryStream();
            var cells = new List<(int Row, int Column, object Value)> {
                (1, 1, (object)"Group"),
                (1, 2, (object)"Name"),
                (1, 3, (object)"Notes")
            };

            for (int row = 2; row <= 701; row++) {
                cells.Add((row, 1, (object)("Repeated value " + (row % 12).ToString(CultureInfo.InvariantCulture))));
                cells.Add((row, 2, (object)("Distinct value " + row.ToString(CultureInfo.InvariantCulture))));
                cells.Add((row, 3, (object)("Long segment " + new string((char)('A' + (row % 26)), 48))));
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.AddWorksheet("Data").CellValues(cells, ExecutionMode.Parallel);
                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using (var spreadsheet = SpreadsheetDocument.Open(memory, false)) {
                var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

                Assert.NotNull(spreadsheet.WorkbookPart.SharedStringTablePart);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString, savedCells["A2"].DataType!.Value);
                Assert.Equal("Repeated value 2", GetSpreadsheetCellText(spreadsheet, savedCells["A2"]));
                Assert.Equal("Distinct value 701", GetSpreadsheetCellText(spreadsheet, savedCells["B701"]));
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(701, 2, out string? value));
            Assert.Equal("Distinct value 701", value);
        }

        [Fact]
        public void PerformanceReview_CellValuesHeaderlessRectangleParallel_UsesDeferredDirectPackageWithSharedStrings() {
            using var memory = new MemoryStream();
            var cells = new List<(int Row, int Column, object Value)>();

            for (int row = 1; row <= 700; row++) {
                cells.Add((row, 1, (object)("Repeated value " + (row % 12).ToString(CultureInfo.InvariantCulture))));
                cells.Add((row, 2, (object)("Distinct value " + row.ToString(CultureInfo.InvariantCulture))));
                cells.Add((row, 3, (object)("Long segment " + new string((char)('A' + (row % 26)), 48))));
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.AddWorksheet("Strings").CellValues(cells, ExecutionMode.Parallel);
                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

            Assert.NotNull(spreadsheet.WorkbookPart.SharedStringTablePart);
            Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString, savedCells["A1"].DataType!.Value);
            Assert.Equal("Repeated value 1", GetSpreadsheetCellText(spreadsheet, savedCells["A1"]));
            Assert.Equal("Distinct value 700", GetSpreadsheetCellText(spreadsheet, savedCells["B700"]));
            Assert.Empty(worksheetPart.TableDefinitionParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValuesHeaderlessRectangleDefault_UsesDeferredDirectPackageAndMaterializesOnRead() {
            using var memory = new MemoryStream();
            var cells = new List<(int Row, int Column, object Value)>();

            for (int row = 1; row <= 100; row++) {
                cells.Add((row, 1, (object)(row * 1.25d)));
                cells.Add((row, 2, (object)(row % 2 == 0)));
                cells.Add((row, 3, (object)("Item " + row.ToString(CultureInfo.InvariantCulture))));
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(cells);

                Assert.True(sheet.TryGetCellText(100, 3, out string? immediateText));
                Assert.Equal("Item 100", immediateText);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? firstValue));
            Assert.True(loaded.Sheets[0].TryGetCellText(100, 3, out string? lastValue));
            Assert.Equal("1.25", firstValue);
            Assert.Equal("Item 100", lastValue);
        }

        [Fact]
        public void PerformanceReview_CellValuesHeaderlessRectangleDefault_UsesDirectPackageWhenWorkbookStaysClean() {
            using var memory = new MemoryStream();
            var cells = new List<(int Row, int Column, object Value)>();

            for (int row = 1; row <= 100; row++) {
                cells.Add((row, 1, (object)(row * 1.25d)));
                cells.Add((row, 2, (object)(row % 2 == 0)));
                cells.Add((row, 3, (object)("Item " + row.ToString(CultureInfo.InvariantCulture))));
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.AddWorksheet("Data").CellValues(cells);
                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("1.25", savedCells["A1"].CellValue!.Text);
            Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean, savedCells["B2"].DataType!.Value);
            Assert.Equal("Item 100", GetSpreadsheetCellText(spreadsheet, savedCells["C100"]));
            Assert.Empty(worksheetPart.TableDefinitionParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValueRowMajorLoop_UsesDirectPackageWhenWorkbookStaysClean() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                for (int row = 1; row <= 100; row++) {
                    sheet.CellValue(row, 1, (double)row * 1.25d);
                    sheet.CellValue(row, 2, row % 2 == 0);
                    sheet.CellValue(row, 3, row);
                }

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("1.25", savedCells["A1"].CellValue!.Text);
            Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean, savedCells["B2"].DataType!.Value);
            Assert.Equal("100", savedCells["C100"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValueTemporalRowMajorLoop_UsesDirectPackageWithCellValueFormats() {
            using var memory = new MemoryStream();
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                for (int row = 1; row <= 100; row++) {
                    sheet.CellValue(row, 1, start.AddDays(row));
                    sheet.CellValue(row, 2, TimeSpan.FromMinutes(row * 7));
                }

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            var formats = spreadsheet.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ToList();

            AssertRoundTripNumericText(start.AddDays(1).ToOADate(), savedCells["A1"].CellValue!.Text);
            AssertRoundTripNumericText(TimeSpan.FromMinutes(700).TotalDays, savedCells["B100"].CellValue!.Text);
            Assert.Equal(14U, formats[(int)savedCells["A1"].StyleIndex!.Value].NumberFormatId!.Value);
            Assert.Equal(46U, formats[(int)savedCells["B1"].StyleIndex!.Value].NumberFormatId!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValueObjectScalarRowMajorLoop_UsesDirectPackageWhenWorkbookStaysClean() {
            using var memory = new MemoryStream();
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                for (int row = 1; row <= 100; row++) {
                    sheet.CellValue(row, 1, (object)("Item " + row.ToString(CultureInfo.InvariantCulture)));
                    sheet.CellValue(row, 2, (object)(row * 1.5d));
                    sheet.CellValue(row, 3, (object)(row % 2 == 0));
                    sheet.CellValue(row, 4, (object)start.AddMinutes(row));
                }

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("Item 100", GetSpreadsheetCellText(spreadsheet, savedCells["A100"]));
            Assert.Equal("150", savedCells["B100"].CellValue!.Text);
            Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean, savedCells["C100"].DataType!.Value);
            AssertRoundTripNumericText(start.AddMinutes(100).ToOADate(), savedCells["D100"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValueObjectSparseNullRowMajorLoop_UsesDirectPackageWithTypedCells() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                WriteSparseCellValueObjectRows(sheet);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            AssertSparseCellValueObjectRows(memory);
        }

        [Fact]
        public void PerformanceReview_CellValueObjectSparseNullBatch_UsesDirectPackageWithTypedCells() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.Batch(WriteSparseCellValueObjectRows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            AssertSparseCellValueObjectRows(memory);
        }

        private static void WriteSparseCellValueObjectRows(ExcelSheet sheet) {
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
            for (int row = 1; row <= 100; row++) {
                object? name = row % 3 == 0 ? null : "Item " + (row % 12).ToString(CultureInfo.InvariantCulture);
                object? amount = row % 4 == 0 ? DBNull.Value : row * 1.25d;
                object? active = row % 5 == 0 ? null : row % 2 == 0;
                object? created = row % 7 == 0 ? null : start.AddDays(row);
                sheet.CellValue(row, 1, name);
                sheet.CellValue(row, 2, amount);
                sheet.CellValue(row, 3, active);
                sheet.CellValue(row, 4, created);
            }
        }

        private static void AssertSparseCellValueObjectRows(MemoryStream memory) {
            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            double expectedCreated = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified)
                .AddDays(100)
                .ToOADate();

            Assert.Equal(CellValues.String, savedCells["A3"].DataType!.Value);
            Assert.Equal(string.Empty, savedCells["A3"].CellValue!.Text);
            Assert.Equal(CellValues.String, savedCells["B4"].DataType!.Value);
            Assert.Equal(string.Empty, savedCells["B4"].CellValue!.Text);
            Assert.Equal(CellValues.String, savedCells["C5"].DataType!.Value);
            Assert.Equal(string.Empty, savedCells["C5"].CellValue!.Text);
            Assert.Equal(CellValues.String, savedCells["D7"].DataType!.Value);
            Assert.Equal(string.Empty, savedCells["D7"].CellValue!.Text);
            Assert.Equal("123.75", savedCells["B99"].CellValue!.Text);
            Assert.Equal(CellValues.Boolean, savedCells["C98"].DataType!.Value);
            Assert.Equal("1", savedCells["C98"].CellValue!.Text);
            AssertRoundTripNumericText(expectedCreated, savedCells["D100"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValueStringRowMajorLoop_UsesDirectPackageWhenWorkbookStaysClean() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                for (int row = 1; row <= 100; row++) {
                    sheet.CellValue(row, 1, "Region " + (row % 8).ToString(CultureInfo.InvariantCulture));
                    sheet.CellValue(row, 2, "Owner " + (row % 16).ToString(CultureInfo.InvariantCulture));
                    sheet.CellValue(row, 3, "Status " + (row % 4).ToString(CultureInfo.InvariantCulture));
                }

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("Region 1", GetSpreadsheetCellText(spreadsheet, savedCells["A1"]));
            Assert.Equal("Owner 4", GetSpreadsheetCellText(spreadsheet, savedCells["B100"]));
            Assert.Equal("Status 0", GetSpreadsheetCellText(spreadsheet, savedCells["C100"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValueStringRowMajorLoop_MaterializesOnReadBeforeSave() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Alpha");
                sheet.CellValue(1, 2, "Beta");
                sheet.CellValue(2, 1, "Gamma");
                sheet.CellValue(2, 2, "Delta");

                Assert.True(sheet.TryGetCellText(2, 1, out string? text));
                Assert.Equal("Gamma", text);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 2, out string? value));
            Assert.Equal("Delta", value);
        }

        [Fact]
        public void PerformanceReview_CellValueMultilineString_UsesDomPathToPreserveWrapStyle() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Top\nBottom");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cell = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().Single();
            Assert.NotNull(cell.StyleIndex);
            Assert.Equal("Top\nBottom", GetSpreadsheetCellText(spreadsheet, cell));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValueRowMajorLoop_MaterializesOnReadBeforeSave() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(1, 2, 2d);
                sheet.CellValue(2, 1, 3d);
                sheet.CellValue(2, 2, 10d);

                Assert.True(sheet.TryGetCellText(2, 1, out string? text));
                Assert.Equal("3", text);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 2, out string? score));
            Assert.Equal("10", score);
        }

        [Fact]
        public void PerformanceReview_CellValueOutOfOrderWriteFallsBackToDomPath() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(1, 2, 2d);
                sheet.CellValue(2, 1, 3d);
                sheet.CellValue(3, 1, 4d);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("1", savedCells["A1"].CellValue!.Text);
            Assert.Equal("3", savedCells["A2"].CellValue!.Text);
            Assert.Equal("4", savedCells["A3"].CellValue!.Text);
            Assert.False(savedCells.ContainsKey("B2"));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValuesDeferredDirectCandidate_MaterializesForNoLockCellMutation() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Name"),
                    (1, 2, (object)"Score"),
                    (2, 1, (object)"Alpha"),
                    (2, 2, (object)10),
                    (3, 1, (object)"Beta"),
                    (3, 2, (object)20)
                }, ExecutionMode.Parallel);

                using (sheet.BeginNoLock()) {
                    sheet.CellValue(2, 2, 999);
                }

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, savedCells["A2"]));
            Assert.Equal("999", GetSpreadsheetCellText(spreadsheet, savedCells["B2"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValuesHeaderlessRectangleThenHeaderedTable_MaterializesBeforeHeaderRepair() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Name"),
                    (1, 2, (object)"Name"),
                    (2, 1, (object)"Alpha"),
                    (2, 2, (object)"Beta")
                }, ExecutionMode.Parallel);
                sheet.AddTable("A1:B2", hasHeader: true, name: "RepairedHeaders", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, savedCells["A1"]));
            Assert.Equal("Name (2)", GetSpreadsheetCellText(spreadsheet, savedCells["B1"]));
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            var columns = tableDefinition.TableColumns!.Elements<TableColumn>().ToList();
            Assert.Equal("Name", columns[0].Name!.Value);
            Assert.Equal("Name (2)", columns[1].Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValuesHeaderThenAppend_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Name"),
                    (1, 2, (object)"Score")
                }, ExecutionMode.Sequential);
                sheet.CellValues(new[] {
                    (2, 1, (object)"Alpha"),
                    (2, 2, (object)10),
                    (3, 1, (object)"Beta"),
                    (3, 2, (object)20)
                }, ExecutionMode.Parallel);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("20", cells["B3"].CellValue!.Text);
            Assert.Empty(worksheetPart.TableDefinitionParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValuesHeaderThenAppend_ReadBeforeSavePreservesDirectPackage() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Name"),
                    (1, 2, (object)"Score")
                }, ExecutionMode.Sequential);
                sheet.CellValues(new[] {
                    (2, 1, (object)"Alpha"),
                    (2, 2, (object)10),
                    (3, 1, (object)"Beta"),
                    (3, 2, (object)20)
                }, ExecutionMode.Parallel);

                Assert.True(sheet.TryGetCellText(3, 2, out string? score));
                Assert.Equal("20", score);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? name));
            Assert.True(loaded.Sheets[0].TryGetCellText(3, 2, out string? scoreAfterSave));
            Assert.Equal("Alpha", name);
            Assert.Equal("20", scoreAfterSave);
        }

        [Fact]
        public void PerformanceReview_DataTableDeferredThenAppend_MaterializesBeforeFallbackSave() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataTable(CreateSingleColumnDataTable("First", "Alpha"));
                sheet.InsertDataTable(CreateSingleColumnDataTable("Second", "Beta"), startRow: 3, includeHeaders: false);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? first));
            Assert.True(loaded.Sheets[0].TryGetCellText(3, 1, out string? second));
            Assert.Equal("Alpha", first);
            Assert.Equal("Beta", second);
        }

        [Fact]
        public void PerformanceReview_DataTableDeferredHeaderMap_MaterializesBeforeDomHeaderFastPath() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataTable(CreateSingleColumnDataTable("Items", "Alpha"));

                Assert.True(sheet.TryGetColumnIndexByHeader("Name", out int columnIndex));
                Assert.Equal(1, columnIndex);

                document.Save(memory);
                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? value));
            Assert.Equal("Alpha", value);
        }

        [Fact]
        public void PerformanceReview_DataTableDeferredThenParallelWrite_MaterializesBeforeFallbackSave() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataTable(CreateSingleColumnDataTable("First", "Alpha"));
                sheet.InsertDataTable(CreateSingleColumnDataTable("Second", "Beta"), startRow: 3, includeHeaders: false, mode: ExecutionMode.Parallel);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? first));
            Assert.True(loaded.Sheets[0].TryGetCellText(3, 1, out string? second));
            Assert.Equal("Alpha", first);
            Assert.Equal("Beta", second);
        }

        [Fact]
        public void PerformanceReview_DataReaderDeferredThenAppend_MaterializesBeforeFallbackSave() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataTable(CreateSingleColumnDataTable("First", "Alpha"));
                using var reader = CreateSingleColumnDataTable("Second", "Beta").CreateDataReader();
                sheet.InsertDataReader(reader, startRow: 3, includeHeaders: false, createTable: false);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? first));
            Assert.True(loaded.Sheets[0].TryGetCellText(3, 1, out string? second));
            Assert.Equal("Alpha", first);
            Assert.Equal("Beta", second);
        }

        [Fact]
        public void PerformanceReview_DataReaderDeferredRegistrationFailureWritesBufferedRows() {
            using var memory = new MemoryStream();
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Alpha");

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                using var reader = table.CreateDataReader();

                MethodInfo? method = typeof(ExcelSheet).GetMethod(
                    "TryInsertDataReaderAsDeferredDirectSave",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                Assert.NotNull(method);

                object?[] args = [
                    reader,
                    new[] { "Name" },
                    new[] { typeof(string), typeof(string) },
                    1,
                    1,
                    true,
                    null,
                    OfficeIMO.Excel.TableStyle.TableStyleMedium2,
                    true,
                    false,
                    false,
                    CancellationToken.None,
                    string.Empty
                ];

                Assert.True((bool)method.Invoke(sheet, args)!);
                Assert.Equal("A1:A2", Assert.IsType<string>(args[12]));

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? text));
            Assert.Equal("Name", header);
            Assert.Equal("Alpha", text);
        }

        [Fact]
        public void PerformanceReview_CellValuesHeaderThenAppend_ExternalMutationPreservesDirectPackageCandidate() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Name"),
                    (1, 2, (object)"Score")
                }, ExecutionMode.Sequential);
                sheet.CellValues(new[] {
                    (2, 1, (object)"Alpha"),
                    (2, 2, (object)10)
                }, ExecutionMode.Parallel);
                sheet.CellValue(5, 1, "Manual edit");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? appended));
            Assert.True(loaded.Sheets[0].TryGetCellText(5, 1, out string? manual));
            Assert.Equal("Alpha", appended);
            Assert.Equal("Manual edit", manual);
        }

        [Fact]
        public void PerformanceReview_CellValuesSingleA1_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var created = new DateTime(2026, 5, 19);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Original")
                });
                sheet.CellValue(2, 1, "Manual edit");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? original));
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? text));
            Assert.Equal("Original", original);
            Assert.Equal("Manual edit", text);
        }

        [Fact]
        public void PerformanceReview_CellValuesRectangle_ExternalMutationPreservesDirectPackageCandidate() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Name"),
                    (1, 2, (object)"Score"),
                    (2, 1, (object)"Alpha"),
                    (2, 2, (object)10)
                });
                sheet.CellValue(5, 1, "Manual edit");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(5, 1, out string? text));
            Assert.Equal("Manual edit", text);
        }

        [Fact]
        public void PerformanceReview_CellValuesSparseRange_DoesNotUseDirectPackageCandidate() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Left"),
                    (1, 3, (object)"Right")
                });

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? left));
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 3, out string? right));
            Assert.Equal("Left", left);
            Assert.Equal("Right", right);
        }

        [Fact]
        public void PerformanceReview_CellValues_NewlineStringSkipsDirectPackageAndPreservesWrapFormatting() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
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
        public void PerformanceReview_CellValueObjectMixed_ReusesDateAndDurationStyles() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
                for (int row = 1; row <= 100; row++) {
                    object? name = "Item " + row.ToString(CultureInfo.InvariantCulture);
                    object? created = start.AddDays(row);
                    object? duration = TimeSpan.FromMinutes(row * 7);
                    object? offset = new DateTimeOffset(start.AddHours(row), TimeSpan.Zero);
                    sheet.CellValue(row, 1, name);
                    sheet.CellValue(row, 2, created);
                    sheet.CellValue(row, 3, duration);
                    sheet.CellValue(row, 4, offset);
                }

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            uint dateStyle = cells["B1"].StyleIndex!.Value;
            uint durationStyle = cells["C1"].StyleIndex!.Value;
            Assert.Equal(dateStyle, cells["B100"].StyleIndex!.Value);
            Assert.Equal(dateStyle, cells["D100"].StyleIndex!.Value);
            Assert.Equal(durationStyle, cells["C100"].StyleIndex!.Value);
            Assert.NotEqual(dateStyle, durationStyle);

            var formats = spreadsheet.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ToList();
            Assert.Equal(14U, formats[(int)dateStyle].NumberFormatId!.Value);
            Assert.Equal(46U, formats[(int)durationStyle].NumberFormatId!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_CellValueDateTimeOffsetRowMajorLoop_UsesDirectPackageWithCustomStrategy() {
            using var memory = new MemoryStream();
            var start = new DateTimeOffset(2026, 1, 1, 8, 30, 0, TimeSpan.FromHours(2));

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
                var sheet = document.AddWorksheet("Data");
                for (int row = 1; row <= 150; row++) {
                    sheet.CellValue(row, 1, start.AddMinutes(row));
                }

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            double expected = start.AddMinutes(150).UtcDateTime.ToOADate();

            AssertRoundTripNumericText(expected, savedCells["A150"].CellValue!.Text);
            Assert.Equal(14U, spreadsheet.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ElementAt((int)savedCells["A150"].StyleIndex!.Value).NumberFormatId!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ExplicitSelectorsUseDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
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
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ExplicitSelectorsReadOnlyListUsesDirectPackageWithoutSnapshotEnumeration() {
            var rows = new ThrowOnEnumerateReadOnlyList<PerformanceObjectExportRow>(
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20)));

            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ExplicitSelectorsSnapshotValuesBeforeDeferredSave() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new MutablePerformanceObjectExportRow { Name = "Alpha", Score = 10 }
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score));

                rows[0].Name = "Changed";
                rows[0].Score = 99;
                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ExplicitSelectorsPreserveBlankHeaders() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows, ("", row => row.Name));

                Assert.True(sheet.TryGetCellText(1, 1, out string? header));
                Assert.Equal(string.Empty, header);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
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
        public void PerformanceReview_InsertObjects_FlatDictionaryRowsUseDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new List<Dictionary<string, object?>> {
                new Dictionary<string, object?> {
                    ["Name"] = "Alpha",
                    ["Score"] = 10,
                    ["Created"] = new DateTime(2026, 5, 19)
                },
                new Dictionary<string, object?> {
                    ["Name"] = "Beta",
                    ["Score"] = 20,
                    ["Created"] = new DateTime(2026, 5, 20),
                    ["Active"] = true
                }
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Score", GetSpreadsheetCellText(spreadsheet, cells["B1"]));
            Assert.Equal("Created", GetSpreadsheetCellText(spreadsheet, cells["C1"]));
            Assert.Equal("Active", GetSpreadsheetCellText(spreadsheet, cells["D1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Equal(string.Empty, GetSpreadsheetCellText(spreadsheet, cells["D2"]));
            Assert.Equal("Beta", GetSpreadsheetCellText(spreadsheet, cells["A3"]));
            Assert.Equal("20", cells["B3"].CellValue!.Text);
            Assert.Equal("1", cells["D3"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_FlatDictionaryRowsAutoFitColumnsForUsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new List<Dictionary<string, object?>> {
                new Dictionary<string, object?> {
                    ["Name"] = "Alpha Region",
                    ["Score"] = 10,
                    ["Created"] = new DateTime(2026, 5, 19)
                },
                new Dictionary<string, object?> {
                    ["Name"] = "Beta Region With Long Name",
                    ["Score"] = 20,
                    ["Created"] = new DateTime(2026, 5, 20)
                }
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.Execution.SaveWorksheetAfterAutoFit = false;
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);
                sheet.AutoFitColumnsFor(new[] { 1, 3 });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
            Assert.NotNull(columns);
            Assert.Equal(new uint[] { 1U, 3U }, columns!.Elements<Column>().Select(column => column.Min!.Value).ToArray());
            Assert.All(columns.Elements<Column>(), column => {
                Assert.True(column.Width?.Value > 0D);
                Assert.True(column.CustomWidth?.Value);
                Assert.True(column.BestFit?.Value);
            });

            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Beta Region With Long Name", GetSpreadsheetCellText(spreadsheet, cells["A3"]));
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_FlatDictionaryWideLateColumnsUseDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new List<Dictionary<string, object?>> {
                new Dictionary<string, object?> {
                    ["Name"] = "Alpha"
                },
                new Dictionary<string, object?> {
                    ["Name"] = "Beta"
                }
            };

            for (int i = 1; i <= 32; i++) {
                rows[1]["Metric" + i.ToString(CultureInfo.InvariantCulture)] = i;
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Metric32", GetSpreadsheetCellText(spreadsheet, cells["AG1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal(string.Empty, GetSpreadsheetCellText(spreadsheet, cells["AG2"]));
            Assert.Equal("Beta", GetSpreadsheetCellText(spreadsheet, cells["A3"]));
            Assert.Equal("32", cells["AG3"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_FlatHashtableRowsUseDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new List<Hashtable> {
                new Hashtable {
                    ["Name"] = "Alpha"
                },
                new Hashtable {
                    ["Name"] = "Beta"
                }
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("Beta", GetSpreadsheetCellText(spreadsheet, cells["A3"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_FlatHashtableRowsPreserveCaseDistinctKeysInDirectPackage() {
            using var memory = new MemoryStream();
            var rows = new List<System.Collections.Specialized.OrderedDictionary> {
                new System.Collections.Specialized.OrderedDictionary {
                    ["Key"] = "Upper",
                    ["key"] = "Lower"
                },
                new System.Collections.Specialized.OrderedDictionary {
                    ["Key"] = "ExactUpper",
                    ["key"] = "ExactLower"
                }
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Key", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("key", GetSpreadsheetCellText(spreadsheet, cells["B1"]));
            Assert.Equal("Upper", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("Lower", GetSpreadsheetCellText(spreadsheet, cells["B2"]));
            Assert.Equal("ExactUpper", GetSpreadsheetCellText(spreadsheet, cells["A3"]));
            Assert.Equal("ExactLower", GetSpreadsheetCellText(spreadsheet, cells["B3"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ReflectionOverloadUsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Score", GetSpreadsheetCellText(spreadsheet, cells["B1"]));
            Assert.Equal("Created", GetSpreadsheetCellText(spreadsheet, cells["C1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ReflectionOverloadSnapshotsValuesBeforeDeferredSave() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new MutablePerformanceObjectExportRow { Name = "Alpha", Score = 10 }
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                rows[0].Name = "Changed";
                rows[0].Score = 99;
                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ReflectionOverloadDirtyWorkbookPreservesSimpleRows() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(5, 5, "Manual edit");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Score", GetSpreadsheetCellText(spreadsheet, cells["B1"]));
            Assert.Equal("Created", GetSpreadsheetCellText(spreadsheet, cells["C1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Equal("Beta", GetSpreadsheetCellText(spreadsheet, cells["A3"]));
            Assert.Equal("Manual edit", GetSpreadsheetCellText(spreadsheet, cells["E5"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ReflectionReadOnlyListUsesDirectPackageWithoutSnapshotEnumeration() {
            var rows = new ThrowOnEnumerateReadOnlyList<PerformanceObjectExportRow>(
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20)));

            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ExternalMutationPreservesDirectPackageCandidate() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.CellValue(4, 1, "Manual edit");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
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
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("A1:C3", tableDefinition.Reference!.Value);
            Assert.Equal("Object_Sales", tableDefinition.Name!.Value);
            Assert.Equal("TableStyleMedium4", tableDefinition.TableStyleInfo!.Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjectsThenAddTableAndAutoFit_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha Region", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta Region With Long Name", 200, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.Execution.SaveWorksheetAfterAutoFit = false;
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C3", hasHeader: true, name: "Object Sales", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AutoFitColumns();

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
            Assert.NotNull(columns);
            var firstColumn = columns!.Elements<Column>().FirstOrDefault(column => column.Min?.Value == 1U && column.Max?.Value == 1U);
            Assert.NotNull(firstColumn);
            Assert.True(firstColumn!.Width?.Value > 10D);
            Assert.True(firstColumn.CustomWidth?.Value);
            Assert.True(firstColumn.BestFit?.Value);
            Assert.Equal("A1:C3", worksheetPart.TableDefinitionParts.Single().Table!.Reference!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjectsThenFullAutoFitColumnsFor_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha Region", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta Region With Long Name", 200, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.Execution.SaveWorksheetAfterAutoFit = false;
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AutoFitColumnsFor(new[] { 1, 2, 3 });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
            Assert.NotNull(columns);
            Assert.Equal(3, columns!.Elements<Column>().Count());
            Assert.All(columns.Elements<Column>(), column => {
                Assert.True(column.Width?.Value > 0D);
                Assert.True(column.CustomWidth?.Value);
                Assert.True(column.BestFit?.Value);
            });
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjectsThenPartialAutoFitColumnsFor_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha Region", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta Region With Long Name", 200, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.Execution.SaveWorksheetAfterAutoFit = false;
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AutoFitColumnsFor(new[] { 1, 3 });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
            Assert.NotNull(columns);
            var columnIndexes = columns!.Elements<Column>().Select(column => column.Min!.Value).ToArray();
            Assert.Equal(new uint[] { 1U, 3U }, columnIndexes);
            Assert.All(columns.Elements<Column>(), column => {
                Assert.True(column.Width?.Value > 0D);
                Assert.True(column.CustomWidth?.Value);
                Assert.True(column.BestFit?.Value);
            });
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjectsThenAutoFitColumn_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 200, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.Execution.SaveWorksheetAfterAutoFit = false;
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AutoFitColumn(2);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
            Assert.NotNull(columns);
            var column = Assert.Single(columns!.Elements<Column>());
            Assert.Equal(2U, column.Min!.Value);
            Assert.Equal(2U, column.Max!.Value);
            Assert.True(column.Width?.Value > 0D);
            Assert.True(column.CustomWidth?.Value);
            Assert.True(column.BestFit?.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_FlatDictionaryPowerShellMixedShapeUsesDirectPackage() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new Dictionary<string, object?> {
                    ["Id"] = 1,
                    ["Name"] = "Server-000001",
                    ["Department"] = "Department-1",
                    ["Region"] = "EU",
                    ["IsEnabled"] = true,
                    ["Created"] = new DateTime(2026, 5, 19, 8, 30, 0),
                    ["Score"] = 123.456D,
                    ["Owner"] = "owner@example.test",
                    ["TicketCount"] = 3,
                    ["Notes"] = "Benchmark row 1"
                },
                new Dictionary<string, object?> {
                    ["Id"] = 2,
                    ["Name"] = "Server-000002",
                    ["Department"] = "Department-2",
                    ["Region"] = "US",
                    ["IsEnabled"] = false,
                    ["Created"] = new DateTime(2026, 5, 20, 9, 45, 0),
                    ["Score"] = 456.789D,
                    ["Owner"] = "owner2@example.test",
                    ["TicketCount"] = 7,
                    ["Notes"] = "Benchmark row 2"
                }
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Id", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("1", cells["A2"].CellValue!.Text);
            Assert.Equal("Server-000001", GetSpreadsheetCellText(spreadsheet, cells["B2"]));
            Assert.Equal(CellValues.Boolean, cells["E2"].DataType!.Value);
            Assert.Equal("1", cells["E2"].CellValue!.Text);
            Assert.Equal(1U, cells["F2"].StyleIndex!.Value);
            Assert.Equal("123.456", cells["G2"].CellValue!.Text);
            Assert.Equal("3", cells["I2"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ObjectTypedFlatDictionaryRowsUseDirectPackage() {
            using var memory = new MemoryStream();
            IReadOnlyList<object?> rows = [
                new Dictionary<string, object?> {
                    ["Id"] = 1,
                    ["Name"] = "Server-000001",
                    ["IsEnabled"] = true,
                    ["Created"] = new DateTime(2026, 5, 19, 8, 30, 0),
                    ["Score"] = 123.456D
                },
                new Dictionary<string, object?> {
                    ["Id"] = 2,
                    ["Name"] = "Server-000002",
                    ["IsEnabled"] = false,
                    ["Created"] = new DateTime(2026, 5, 20, 9, 45, 0),
                    ["Score"] = 456.789D
                }
            ];

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("1", cells["A2"].CellValue!.Text);
            Assert.Equal("Server-000001", GetSpreadsheetCellText(spreadsheet, cells["B2"]));
            Assert.Equal(CellValues.Boolean, cells["C2"].DataType!.Value);
            Assert.Equal("1", cells["C2"].CellValue!.Text);
            Assert.Equal(1U, cells["D2"].StyleIndex!.Value);
            Assert.Equal("123.456", cells["E2"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PowerShellObjectBagShapeUsesDirectPackage() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new System.Management.Automation.PSObject(
                    new System.Management.Automation.PSPropertyInfo("Id", 1),
                    new System.Management.Automation.PSPropertyInfo("Name", "Server-000001"),
                    new System.Management.Automation.PSPropertyInfo("IsEnabled", true),
                    new System.Management.Automation.PSPropertyInfo("Created", new DateTime(2026, 5, 19, 8, 30, 0)),
                    new System.Management.Automation.PSPropertyInfo("Score", 123.456D)),
                new System.Management.Automation.PSObject(
                    new System.Management.Automation.PSPropertyInfo("Id", 2),
                    new System.Management.Automation.PSPropertyInfo("Name", "Server-000002"),
                    new System.Management.Automation.PSPropertyInfo("IsEnabled", false),
                    new System.Management.Automation.PSPropertyInfo("Created", new DateTime(2026, 5, 20, 9, 45, 0)),
                    new System.Management.Automation.PSPropertyInfo("Score", 456.789D))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Id", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Server-000001", GetSpreadsheetCellText(spreadsheet, cells["B2"]));
            Assert.Equal(CellValues.Boolean, cells["C2"].DataType!.Value);
            Assert.Equal(1U, cells["D2"].StyleIndex!.Value);
            Assert.Equal("123.456", cells["E2"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PowerShellObjectBagWideShapeUsesDirectPackage() {
            using var memory = new MemoryStream();
            var rows = new[] {
                CreateWidePowerShellObject(1),
                CreateWidePowerShellObject(2)
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Metric36", GetSpreadsheetCellText(spreadsheet, cells["AN1"]));
            Assert.Equal("36", cells["AN2"].CellValue!.Text);
            Assert.Equal("72", cells["AN3"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());

            static System.Management.Automation.PSObject CreateWidePowerShellObject(int id) {
                var properties = new System.Management.Automation.PSPropertyInfo[40];
                properties[0] = new System.Management.Automation.PSPropertyInfo("Id", id);
                properties[1] = new System.Management.Automation.PSPropertyInfo("Name", "Server-" + id.ToString("D6", CultureInfo.InvariantCulture));
                properties[2] = new System.Management.Automation.PSPropertyInfo("Created", new DateTime(2026, 5, 20).AddDays(id));
                properties[3] = new System.Management.Automation.PSPropertyInfo("Enabled", id % 2 == 1);
                for (int metric = 1; metric <= 36; metric++) {
                    properties[metric + 3] = new System.Management.Automation.PSPropertyInfo("Metric" + metric.ToString(CultureInfo.InvariantCulture), id * metric);
                }

                return new System.Management.Automation.PSObject(properties);
            }
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PowerShellObjectBagLateColumnsUseDirectPackage() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new System.Management.Automation.PSObject(
                    new System.Management.Automation.PSPropertyInfo("Id", 1),
                    new System.Management.Automation.PSPropertyInfo("Name", "Server-000001")),
                new System.Management.Automation.PSObject(
                    new System.Management.Automation.PSPropertyInfo("Id", 2),
                    new System.Management.Automation.PSPropertyInfo("Name", "Server-000002"),
                    new System.Management.Automation.PSPropertyInfo("TicketCount", 3))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("TicketCount", GetSpreadsheetCellText(spreadsheet, cells["C1"]));
            Assert.True(!cells.TryGetValue("C2", out var blankTicketCount)
                || string.IsNullOrEmpty(GetSpreadsheetCellText(spreadsheet, blankTicketCount)));
            Assert.Equal("3", cells["C3"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

#if NET6_0_OR_GREATER
        [Fact]
        public void PerformanceReview_InsertObjects_ExtendedSimpleScalarTypesUseDirectPackage() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new ExtendedPerformanceObjectExportRow("Alpha", PerformanceExportStatus.Ready, TimeSpan.FromMinutes(95), new DateOnly(2026, 5, 19), new TimeOnly(8, 30)),
                new ExtendedPerformanceObjectExportRow("Beta", PerformanceExportStatus.Waiting, TimeSpan.FromMinutes(125), new DateOnly(2026, 5, 20), new TimeOnly(9, 45))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("Ready", GetSpreadsheetCellText(spreadsheet, cells["B2"]));
            Assert.Equal(2U, cells["C2"].StyleIndex!.Value);
            Assert.Equal(1U, cells["D2"].StyleIndex!.Value);
            Assert.Equal(2U, cells["E2"].StyleIndex!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }
#endif

        [Fact]
        public void PerformanceReview_InsertObjectsThenSubsetAutoFit_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha Region With Longer Name", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.Execution.SaveWorksheetAfterAutoFit = false;
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AutoFitColumnsFor(new[] { 1 });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
            Assert.NotNull(columns);
            var firstColumn = columns!.Elements<Column>().FirstOrDefault(column => column.Min?.Value == 1U && column.Max?.Value == 1U);
            Assert.NotNull(firstColumn);
            Assert.True(firstColumn!.Width?.Value > 10D);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_MetadataOnlyRulesUseDirectPackage() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("Gamma", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddConditionalRule("B2:B4", ConditionalFormattingOperatorValues.GreaterThan, "15");
                sheet.AddConditionalColorScale("B2:B4", OfficeIMO.Drawing.OfficeColor.LightPink, OfficeIMO.Drawing.OfficeColor.LightGreen);
                sheet.AddConditionalDataBar("B2:B4", OfficeIMO.Drawing.OfficeColor.SteelBlue);
                sheet.ValidationWholeNumber("B2:B4", DataValidationOperatorValues.Between, 1, 100);
                sheet.Freeze(topRows: 1, leftCols: 1);
                sheet.AddAutoFilter("A1:C4");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("30", cells["B4"].CellValue!.Text);
            var ruleTypes = worksheetPart.Worksheet.Elements<ConditionalFormatting>()
                .SelectMany(formatting => formatting.Elements<ConditionalFormattingRule>())
                .Select(rule => rule.Type?.Value)
                .ToList();
            Assert.Equal(3, ruleTypes.Count);
            Assert.Contains(ConditionalFormatValues.CellIs, ruleTypes);
            Assert.Contains(ConditionalFormatValues.ColorScale, ruleTypes);
            Assert.Contains(ConditionalFormatValues.DataBar, ruleTypes);
            Assert.NotNull(worksheetPart.Worksheet.GetFirstChild<SheetViews>());
            Assert.Equal("A1:C4", worksheetPart.Worksheet.GetFirstChild<AutoFilter>()!.Reference!.Value);
            Assert.Equal("B2:B4", worksheetPart.Worksheet.GetFirstChild<DataValidations>()!.Elements<DataValidation>().Single().SequenceOfReferences!.InnerText);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_DeferredFreezePreservesExistingSheetViewAttributes() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.SetGridlinesVisible(false);
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.Freeze(topRows: 1, leftCols: 1);

                document.Save(memory);

                Assert.True(
                    document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.DirectDataSetPackage
                    || document.LastSaveDiagnostics.Writer == ExcelSavePackageWriter.SimplePackage,
                    document.LastSaveDiagnostics.FastPackageSkipReason ?? "Expected a fast package writer.");
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var sheetView = worksheetPart.Worksheet.GetFirstChild<SheetViews>()!.GetFirstChild<SheetView>()!;
            Assert.False(sheetView.ShowGridLines!.Value);
            Assert.Equal(PaneValues.BottomRight, sheetView.GetFirstChild<Pane>()!.ActivePane!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ChartWorkbookPreservesChartAndDeferredDataFastPath() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                var chartData = new ExcelChartData(
                    rows.Select(row => row.Name),
                    new[] { new ExcelChartSeries("Score", rows.Select(row => (double)row.Score)) });

                sheet.AddChart(chartData, row: 6, column: 5, widthPixels: 480, heightPixels: 280, type: ExcelChartType.ColumnClustered, title: "Scores");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var workbookPart = spreadsheet.WorkbookPart!;
            Assert.Single(workbookPart.GetPartsOfType<ThemePart>());
            var worksheetParts = workbookPart.WorksheetParts.ToList();
            Assert.Equal(2, worksheetParts.Count);
            var chartHostPart = Assert.Single(worksheetParts, part => part.DrawingsPart?.ChartParts.Count() == 1);
            Assert.Single(chartHostPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_RangeChartBeforePivotPreservesDeferredDataAndDrawing() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AddChartFromRange("A1:B4", row: 2, column: 6, widthPixels: 480, heightPixels: 280, type: ExcelChartType.ColumnClustered, title: "Scores", includeCachedData: false);
                sheet.AddPivotTable(
                    "A1:C4",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var workbookPart = spreadsheet.WorkbookPart!;
            var worksheetPart = workbookPart.WorksheetParts.First();
            Assert.Single(worksheetPart.PivotTableParts);
            Assert.Single(worksheetPart.DrawingsPart!.ChartParts);
            Assert.Single(worksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("North", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("30", GetSpreadsheetCellText(spreadsheet, cells["B4"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_IndependentChartBeforePivotResolvesDeferredHeaders() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                var chartData = new ExcelChartData(
                    rows.Select(row => row.Name),
                    new[] { new ExcelChartSeries("Score", rows.Select(row => (double)row.Score)) });
                sheet.AddChart(chartData, row: 2, column: 6, widthPixels: 480, heightPixels: 280, type: ExcelChartType.ColumnClustered, title: "Scores");
                sheet.AddPivotTable(
                    "A1:C4",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var workbookPart = spreadsheet.WorkbookPart!;
            var dataPart = workbookPart.WorksheetParts.Single(part => part.PivotTableParts.Any());
            Assert.Single(dataPart.PivotTableParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PostFeatureExternalCellMutationKeepsExtendedFastPath() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AddPivotTable(
                    "A1:C4",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });
                var chartData = new ExcelChartData(
                    rows.Select(row => row.Name),
                    new[] { new ExcelChartSeries("Score", rows.Select(row => (double)row.Score)) });
                sheet.AddChart(chartData, row: 2, column: 6, widthPixels: 480, heightPixels: 280, type: ExcelChartType.ColumnClustered, title: "Scores");
                sheet.CellValue(rows.Length + 4, 1, "Manual note after report features");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.PivotTableParts.Any());
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.True(cells.ContainsKey("A7"), "Saved cells: " + string.Join(", ", cells.Keys.OrderBy(static key => key, StringComparer.Ordinal)));
            Assert.Equal("Manual note after report features", GetSpreadsheetCellText(spreadsheet, cells["A7"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ChartKeepsDeferredRowsWhenFastSaveModelUnavailable() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);

                var relationship = sheet.WorksheetPart.AddHyperlinkRelationship(new Uri("https://example.org/deferred"), true);
                var hyperlinks = sheet.WorksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
                if (hyperlinks == null) {
                    hyperlinks = new Hyperlinks();
                    var tableParts = sheet.WorksheetPart.Worksheet.Elements<TableParts>().FirstOrDefault();
                    if (tableParts == null) {
                        sheet.WorksheetPart.Worksheet.Append(hyperlinks);
                    } else {
                        sheet.WorksheetPart.Worksheet.InsertBefore(hyperlinks, tableParts);
                    }
                }

                hyperlinks.Append(new Hyperlink { Reference = "D7", Id = relationship.Id });

                var chartData = new ExcelChartData(
                    rows.Select(row => row.Name),
                    new[] { new ExcelChartSeries("Score", rows.Select(row => (double)row.Score)) });
                sheet.AddChart(chartData, row: 2, column: 6, widthPixels: 480, heightPixels: 280, type: ExcelChartType.ColumnClustered, title: "Scores");

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart != null);
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("North", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("East", GetSpreadsheetCellText(spreadsheet, cells["A4"]));
            Assert.Single(worksheetPart.HyperlinkRelationships);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PostFeatureSideColumnMutationKeepsExtendedFastPath() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var links = document.AddWorksheet("Links");
                links.SetHyperlink(1, 1, "https://example.org/review", display: "Review link", style: false);

                var customPart = document._spreadSheetDocument.AddCustomFilePropertiesPart();
                customPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties(
                    new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty(
                        new Vt.VTLPWSTR("preserved")) {
                            FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                            PropertyId = 2,
                            Name = "OfficeIMOReview"
                        });

                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AddPivotTable(
                    "A1:C4",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });
                sheet.CellValue(2, 4, 123.45d);
                sheet.CellAt(2, 4).SetNumberFormat("0.00");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.PivotTableParts.Any());
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("North", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("123.45", GetSpreadsheetCellText(spreadsheet, cells["D2"]));
            Assert.Equal("0.00", GetCellNumberFormatCode(spreadsheet, cells["D2"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ExternalCellMutationDoesNotMaterializeDeferredRows() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));

                Assert.True(document.HasDeferredDirectDataSetImport);

                sheet.CellValue(2, 4, "Manual side note");

                Assert.True(document.HasDeferredDirectDataSetImport);
                var sourceCells = sheet.WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                Assert.Equal("Manual side note", GetSpreadsheetCellText(document._spreadSheetDocument, sourceCells["D2"]));

                document.Save(memory);

                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("North", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("Manual side note", GetSpreadsheetCellText(spreadsheet, cells["D2"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_DirectOverlayNumberFormatRemapsIntoDirectStyles() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.CellValue(2, 4, 123.45d);
                sheet.CellAt(2, 4).SetNumberFormat("0.00");
                sheet.CellValue(2, 5, 45000d);
                sheet.CellValue(2, 6, 30d);

                var stylesheet = document._spreadSheetDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                stylesheet.CellFormats ??= new CellFormats();
                uint builtInDateStyleIndex = (uint)stylesheet.CellFormats.Elements<CellFormat>().Count();
                stylesheet.CellFormats.Append(new CellFormat { NumberFormatId = 14U, ApplyNumberFormat = true });
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Elements<CellFormat>().Count();
                var sourceCells = sheet.WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                sourceCells["E2"].StyleIndex = builtInDateStyleIndex;
                sourceCells["F2"].CellFormula = new CellFormula("B2*3");
                sourceCells["F2"].CellValue = new CellValue("30");
                sourceCells["F2"].DataType = null;

                document.Save(memory);

                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("North", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("123.45", GetSpreadsheetCellText(spreadsheet, cells["D2"]));
            Assert.Equal("0.00", GetCellNumberFormatCode(spreadsheet, cells["D2"]));
            Assert.Equal("45000", GetSpreadsheetCellText(spreadsheet, cells["E2"]));
            Assert.Equal("mm-dd-yy", GetCellNumberFormatCode(spreadsheet, cells["E2"]));
            Assert.Equal("B2*3", cells["F2"].CellFormula!.Text);
            Assert.Equal("30", cells["F2"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_DirectOverlayRichStyleFallsBackAndPersists() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.CellValue(2, 4, "Styled side note");

                uint styleIndex = AddBoldCellStyle(document._spreadSheetDocument);
                var sourceCells = sheet.WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                sourceCells["D2"].StyleIndex = styleIndex;

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            AssertRichOverlayStylePersisted(memory);
        }

        [Fact]
        public void PerformanceReview_InsertObjects_MaterializationPreservesRichOverlayStyle() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.CellValue(2, 4, "Styled side note");

                uint styleIndex = AddBoldCellStyle(document._spreadSheetDocument);
                var sourceCells = sheet.WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                sourceCells["D2"].StyleIndex = styleIndex;

                document.MaterializeDeferredDataSetImport();
                document.Save(memory);
            }

            AssertRichOverlayStylePersisted(memory);
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ChartFallbackPreservesRichOverlayStyle() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.CellValue(2, 4, "Styled side note");

                uint styleIndex = AddBoldCellStyle(document._spreadSheetDocument);
                var sourceCells = sheet.WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                sourceCells["D2"].StyleIndex = styleIndex;

                var chartData = new ExcelChartData(
                    rows.Select(row => row.Name),
                    new[] { new ExcelChartSeries("Score", rows.Select(row => (double)row.Score)) });
                sheet.AddChart(chartData, row: 2, column: 6, widthPixels: 480, heightPixels: 280, type: ExcelChartType.ColumnClustered, title: "Scores");

                document.Save(memory);
            }

            AssertRichOverlayStylePersisted(memory);
        }

        [Fact]
        public void PerformanceReview_InsertObjects_DirectOverlayClearRemovesPreservedCell() {
            using var firstMemory = new MemoryStream();
            using var secondMemory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.CellValue(2, 4, "Manual");

                document.Save(firstMemory);

                var sourceCells = sheet.WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                sourceCells["D2"].CellFormula = null;
                sourceCells["D2"].CellValue = null;
                sourceCells["D2"].DataType = null;
                sourceCells["D2"].InlineString = null;

                document.Save(secondMemory);
            }

            secondMemory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(secondMemory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.True(!cells.ContainsKey("D2") || string.IsNullOrEmpty(GetSpreadsheetCellText(spreadsheet, cells["D2"])));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_DirectOverlayClearBeyondTableDoesNotPreserveStaleValue() {
            using var firstMemory = new MemoryStream();
            using var secondMemory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AddPivotTable(
                    "A1:C4",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });
                sheet.CellValue(10, 4, "Manual");

                document.Save(firstMemory);

                sheet.CellValue(10, 4, (object?)null);

                document.Save(secondMemory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            secondMemory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(secondMemory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.True(!cells.ContainsKey("D10") || string.IsNullOrEmpty(GetSpreadsheetCellText(spreadsheet, cells["D10"])));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PathSaveUsesExtendedFastSaveModel() {
            var path = Path.Combine(Path.GetTempPath(), "officeimo-path-extended-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try {
                var rows = new[] {
                    new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                    new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                    new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
                };

                using (var document = ExcelDocument.Create(new MemoryStream())) {
                    var sheet = document.AddWorksheet("Data");
                    sheet.InsertObjects(rows,
                        ("Name", row => row.Name),
                        ("Score", row => row.Score),
                        ("Created", row => row.Created));
                    sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                    sheet.AddPivotTable(
                        "A1:C4",
                        "F20",
                        rowFields: new[] { "Name" },
                        dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                    document.Save(path, null);

                    Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                    Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
                }

                using var spreadsheet = SpreadsheetDocument.Open(path, false);
                var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.PivotTableParts.Any());
                var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                Assert.Equal("North", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
                Assert.Equal("30", GetSpreadsheetCellText(spreadsheet, cells["B4"]));
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
            } finally {
                if (File.Exists(path)) {
                    File.Delete(path);
                }
            }
        }

        [Fact]
        public void PerformanceReview_InsertObjects_NonSeekableStreamSaveUsesExtendedFastSaveModel() {
            using var stream = new NonSeekableWriteStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AddPivotTable(
                    "A1:C4",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                document.Save(stream);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            using var memory = new MemoryStream(stream.ToArray());
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.PivotTableParts.Any());
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("North", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("East", GetSpreadsheetCellText(spreadsheet, cells["A4"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ExtendedSimpleDrawingSheetDeclaresRelationshipNamespace() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var chartSheet = document.AddWorksheet("ChartOnly");
                chartSheet.CellValue(1, 1, "Name");
                chartSheet.CellValue(1, 2, "Score");
                chartSheet.CellValue(2, 1, "North");
                chartSheet.CellValue(2, 2, 10);
                chartSheet.CellValue(3, 1, "South");
                chartSheet.CellValue(3, 2, 20);
                chartSheet.CellValue(4, 1, "East");
                chartSheet.CellValue(4, 2, 30);
                chartSheet.AddChartFromRange("A1:B4", row: 2, column: 5, widthPixels: 480, heightPixels: 280, type: ExcelChartType.ColumnClustered, title: "Scores", includeCachedData: false);

                var dataSheet = document.AddWorksheet("Data");
                dataSheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                dataSheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                dataSheet.AddPivotTable(
                    "A1:C4",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var chartHostPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart?.ChartParts.Any() == true);
            using (var reader = new StreamReader(chartHostPart.GetStream(FileMode.Open, FileAccess.Read))) {
                string xml = reader.ReadToEnd();
                Assert.Contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", xml);
                Assert.Contains("r:id=", xml);
            }

            Assert.Single(chartHostPart.DrawingsPart!.ChartParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_OverlaySharedFormulaMetadataIsPreserved() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AddPivotTable(
                    "A1:C4",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                var sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                var row2 = sheetData.Elements<Row>().FirstOrDefault(row => row.RowIndex?.Value == 2U)
                           ?? sheetData.AppendChild(new Row { RowIndex = 2U });
                row2.Append(new Cell {
                    CellReference = "D2",
                    CellFormula = new CellFormula("B2*2") {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 0U,
                        Reference = "D2:D3"
                    },
                    CellValue = new CellValue("20")
                });
                var row3 = sheetData.Elements<Row>().FirstOrDefault(row => row.RowIndex?.Value == 3U)
                           ?? sheetData.AppendChild(new Row { RowIndex = 3U });
                row3.Append(new Cell {
                    CellReference = "D3",
                    CellFormula = new CellFormula {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 0U
                    },
                    CellValue = new CellValue("40")
                });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.PivotTableParts.Any());
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(CellFormulaValues.Shared, cells["D2"].CellFormula!.FormulaType!.Value);
            Assert.Equal(0U, cells["D2"].CellFormula!.SharedIndex!.Value);
            Assert.Equal("D2:D3", cells["D2"].CellFormula!.Reference!.Value);
            Assert.Equal("B2*2", cells["D2"].CellFormula!.Text);
            Assert.Equal(CellFormulaValues.Shared, cells["D3"].CellFormula!.FormulaType!.Value);
            Assert.Equal(0U, cells["D3"].CellFormula!.SharedIndex!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_OverlayTypedCellsPreserveDataTypes() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AddPivotTable(
                    "A1:C4",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });
                sheet.CellValue(2, 4, string.Empty);
                sheet.CellValue(2, 5, string.Empty);

                var sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                var row2 = sheetData.Elements<Row>().First(row => row.RowIndex?.Value == 2U);
                var d2 = row2.Elements<Cell>().First(cell => cell.CellReference?.Value == "D2");
                d2.DataType = CellValues.Error;
                d2.CellValue = new CellValue("#DIV/0!");
                var e2 = row2.Elements<Cell>().First(cell => cell.CellReference?.Value == "E2");
                e2.DataType = CellValues.Date;
                e2.CellValue = new CellValue("2026-05-19");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.PivotTableParts.Any());
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(CellValues.Error, cells["D2"].DataType!.Value);
            Assert.Equal("#DIV/0!", cells["D2"].CellValue!.Text);
            Assert.Equal(CellValues.Date, cells["E2"].DataType!.Value);
            Assert.Equal("2026-05-19", cells["E2"].CellValue!.Text);
        }

        [Fact]
        public void PerformanceReview_InsertObjects_CompactOverlayCellReferencesKeepSideColumnMutation() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AddPivotTable(
                    "A1:C4",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });
                sheet.CellValue(2, 4, "Compact side note");

                var sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                sheetData.RemoveAllChildren<Row>();
                sheetData.Append(new Row());
                sheetData.Append(new Row(
                    new Cell(),
                    new Cell(),
                    new Cell(),
                    new Cell {
                        DataType = CellValues.String,
                        CellValue = new CellValue("Compact side note")
                    }));

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.PivotTableParts.Any());
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("North", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("Compact side note", GetSpreadsheetCellText(spreadsheet, cells["D2"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PostFeatureInRangeMutationMaterializesSourceRows() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AddPivotTable(
                    "A1:C4",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });
                sheet.CellValue(2, 2, 42);

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.PivotTableParts.Any());
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("North", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("42", cells["B2"].CellValue!.Text);
            Assert.Equal("East", GetSpreadsheetCellText(spreadsheet, cells["A4"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_ExtendedPackagePreservesSharedStringsStylesHyperlinksAndRootParts() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Columns.Add("Created", typeof(DateTime));
            for (int index = 0; index < 600; index++) {
                table.Rows.Add(
                    index % 2 == 0 ? "North Region Repeated" : "South Region Repeated",
                    index,
                    new DateTime(2026, 5, 19).AddDays(index % 7));
            }

            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document._spreadSheetDocument.PackageProperties.Title = "Review Metadata";
                document._spreadSheetDocument.PackageProperties.Creator = "OfficeIMO Tests";
                document._spreadSheetDocument.PackageProperties.LastModifiedBy = "Reviewer";
                document._spreadSheetDocument.PackageProperties.Created = new DateTime(2026, 5, 26, 12, 0, 0, DateTimeKind.Utc);
                document._spreadSheetDocument.PackageProperties.Modified = new DateTime(2026, 5, 26, 13, 0, 0, DateTimeKind.Utc);

                var appPart = document._spreadSheetDocument.ExtendedFilePropertiesPart
                              ?? document._spreadSheetDocument.AddExtendedFilePropertiesPart();
                appPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties {
                    Application = new DocumentFormat.OpenXml.ExtendedProperties.Application { Text = "OfficeIMO Review App" },
                    Company = new DocumentFormat.OpenXml.ExtendedProperties.Company { Text = "Evotec" },
                    Manager = new DocumentFormat.OpenXml.ExtendedProperties.Manager { Text = "OfficeIMO" }
                };

                var links = document.AddWorksheet("Links");
                links.SetHyperlink(1, 1, "https://example.org/review", display: "Review link", style: false);
                links.CellValue(2, 1, 123.45d);
                links.CellAt(2, 1).SetNumberFormat("0.00");
                links.CellValue(3, 1, "#DIV/0!");
                var linkSourceCells = links.WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                linkSourceCells["A3"].DataType = CellValues.Error;
                linkSourceCells["A3"].CellValue = new CellValue("#DIV/0!");

                var customPart = document._spreadSheetDocument.AddCustomFilePropertiesPart();
                customPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties(
                    new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty(
                        new Vt.VTLPWSTR("preserved")) {
                            FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                            PropertyId = 2,
                            Name = "OfficeIMOReview"
                        });
                document._spreadSheetDocument.AddExternalRelationship(
                    "https://schemas.example.org/officeimo/review",
                    new Uri("https://example.org/package-root"),
                    "rIdCore");

                document.InsertDataSet(dataSet);
                var sheet = document.Sheets.First(item => string.Equals(item.Name, "Items", StringComparison.Ordinal));
                sheet.AddPivotTable(
                    "A1:C601",
                    "F20",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            Assert.Equal("Review Metadata", spreadsheet.PackageProperties.Title);
            Assert.Equal("OfficeIMO Tests", spreadsheet.PackageProperties.Creator);
            Assert.Equal("Reviewer", spreadsheet.PackageProperties.LastModifiedBy);
            Assert.Equal("OfficeIMO Review App", spreadsheet.ExtendedFilePropertiesPart!.Properties!.Application!.Text);
            Assert.Equal("Evotec", spreadsheet.ExtendedFilePropertiesPart.Properties.Company!.Text);
            Assert.Equal("OfficeIMO", spreadsheet.ExtendedFilePropertiesPart.Properties.Manager!.Text);
            Assert.NotNull(spreadsheet.WorkbookPart!.SharedStringTablePart);
            Assert.NotNull(spreadsheet.CustomFilePropertiesPart);
            var packageRelationship = Assert.Single(spreadsheet.ExternalRelationships);
            Assert.Equal("rIdCore", packageRelationship.Id);
            Assert.Equal("https://schemas.example.org/officeimo/review", packageRelationship.RelationshipType);
            Assert.Equal("https://example.org/package-root", packageRelationship.Uri.ToString());

            var sheets = spreadsheet.WorkbookPart.Workbook.Sheets!.Elements<Sheet>().ToDictionary(sheet => sheet.Name!.Value!);
            var dataPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(sheets["Items"].Id!);
            var dataCells = dataPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("North Region Repeated", GetSpreadsheetCellText(spreadsheet, dataCells["A2"]));
            Assert.Equal("South Region Repeated", GetSpreadsheetCellText(spreadsheet, dataCells["A601"]));

            var linksPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(sheets["Links"].Id!);
            var linkCells = linksPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Review link", GetSpreadsheetCellText(spreadsheet, linkCells["A1"]));
            Assert.Equal("0.00", GetCellNumberFormatCode(spreadsheet, linkCells["A2"]));
            Assert.Equal(CellValues.Error, linkCells["A3"].DataType!.Value);
            Assert.Equal("#DIV/0!", linkCells["A3"].CellValue!.Text);
            Assert.Contains("xmlns:r=", linksPart.Worksheet.OuterXml, StringComparison.Ordinal);
            var relationship = Assert.Single(linksPart.HyperlinkRelationships);
            Assert.Equal("https://example.org/review", relationship.Uri.ToString());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_StyledChartPreservesDeferredRowsInExtendedPackage() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.DefaultChartStylePreset = ExcelChartStylePreset.Default;
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                var chartData = new ExcelChartData(
                    rows.Select(row => row.Name),
                    new[] { new ExcelChartSeries("Score", rows.Select(row => (double)row.Score)) });

                sheet.AddChart(chartData, row: 1, column: 6, widthPixels: 360, heightPixels: 240, type: ExcelChartType.ColumnClustered, title: "Styled");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var dataSheet = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>()
                .Single(sheet => string.Equals(sheet.Name?.Value, "Data", StringComparison.Ordinal));
            var dataPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(dataSheet.Id!);
            var cells = dataPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("20", cells["B3"].CellValue!.Text);

            var chartPart = dataPart.DrawingsPart!.ChartParts.Single();
            var stylePart = Assert.Single(chartPart.GetPartsOfType<ChartStylePart>());
            var colorStylePart = Assert.Single(chartPart.GetPartsOfType<ChartColorStylePart>());
            using (var styleStream = stylePart.GetStream(FileMode.Open, FileAccess.Read)) {
                Assert.True(styleStream.Length > 0);
            }
            using (var colorStyleStream = colorStylePart.GetStream(FileMode.Open, FileAccess.Read)) {
                Assert.True(colorStyleStream.Length > 0);
            }
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ReportMetadataPersistsFeaturesWithDirectPackage() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C3", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AutoFitColumns();
                sheet.Freeze(topRows: 1, leftCols: 1);
                sheet.AddAutoFilter("A1:C3");
                sheet.AddConditionalRule("B2:B3", ConditionalFormattingOperatorValues.GreaterThan, "15");
                sheet.ValidationWholeNumber("B2:B3", DataValidationOperatorValues.Between, 1, 100);
                sheet.CellValue(4, 1, "Manual note");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("20", cells["B3"].CellValue!.Text);
            Assert.Equal("Manual note", GetSpreadsheetCellText(spreadsheet, cells["A4"]));

            Assert.NotNull(worksheetPart.Worksheet.GetFirstChild<SheetViews>());
            Assert.NotNull(worksheetPart.Worksheet.GetFirstChild<Columns>());
            Assert.Single(worksheetPart.Worksheet.Elements<ConditionalFormatting>());
            Assert.Equal("B2:B3", worksheetPart.Worksheet.GetFirstChild<DataValidations>()!.Elements<DataValidation>().Single().SequenceOfReferences!.InnerText);

            var tablePart = Assert.Single(worksheetPart.TableDefinitionParts);
            Assert.Equal("A1:C3", tablePart.Table!.Reference!.Value);
            Assert.Equal("ReportData", tablePart.Table.Name!.Value);
            Assert.Equal("A1:C3", tablePart.Table.GetFirstChild<AutoFilter>()!.Reference!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_ReportMetadataUsesDirectPackageWhenColumnsAreDeferred() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C3", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AutoFitColumns();
                sheet.Freeze(topRows: 1, leftCols: 1);
                sheet.AddAutoFilter("A1:C3");
                sheet.AddConditionalRule("B2:B3", ConditionalFormattingOperatorValues.GreaterThan, "15");
                sheet.ValidationWholeNumber("B2:B3", DataValidationOperatorValues.Between, 1, 100);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("20", cells["B3"].CellValue!.Text);

            Assert.NotNull(worksheetPart.Worksheet.GetFirstChild<SheetViews>());
            var columns = Assert.Single(worksheetPart.Worksheet.Elements<Columns>());
            Assert.Equal(3, columns.Elements<Column>().Count());
            Assert.Single(worksheetPart.Worksheet.Elements<ConditionalFormatting>());
            Assert.Equal("B2:B3", worksheetPart.Worksheet.GetFirstChild<DataValidations>()!.Elements<DataValidation>().Single().SequenceOfReferences!.InnerText);

            var tablePart = Assert.Single(worksheetPart.TableDefinitionParts);
            Assert.Equal("A1:C3", tablePart.Table!.Reference!.Value);
            Assert.Equal("ReportData", tablePart.Table.Name!.Value);
            Assert.Equal("A1:C3", tablePart.Table.GetFirstChild<AutoFilter>()!.Reference!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_HeaderNumberFormatsUseDirectPackageWhenColumnsAreDeferred() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19, 8, 30, 0)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20, 9, 45, 0))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C3", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AutoFitColumns();
                sheet.ColumnStyleByHeader("Score").NumberFormat("#,##0.000");
                sheet.ColumnStyleByHeader("Created").NumberFormat("yyyy-mm-dd hh:mm");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("#,##0.000", GetCellNumberFormatCode(spreadsheet, cells["B2"]));
            Assert.Equal("yyyy-mm-dd hh:mm", GetCellNumberFormatCode(spreadsheet, cells["C2"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_HeaderNumberFormatSurvivesDeferredModelClones() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19, 8, 30, 0)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20, 9, 45, 0))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.ColumnStyleByHeader("Score").NumberFormat("#,##0.000");
                sheet.AddTable("A1:C3", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AutoFitColumns();

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("#,##0.000", GetCellNumberFormatCode(spreadsheet, cells["B2"]));
            Assert.Equal("ReportData", worksheetPart.TableDefinitionParts.Single().Table!.Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_HeaderNumberFormatUsesNormalizedDirectHeader() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows, ("Total   Amount ", row => row.Score));
                sheet.ColumnStyleByHeader("Total Amount").NumberFormat("0.00");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("0.00", GetCellNumberFormatCode(spreadsheet, cells["A2"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_HeaderNumberFormatMatchesDirectSheetInMultiSheetModel() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var north = new DataTable("North");
            north.Columns.Add("Name", typeof(string));
            north.Columns.Add("Score", typeof(int));
            north.Rows.Add("Alpha", 10);
            var south = new DataTable("South");
            south.Columns.Add("Name", typeof(string));
            south.Columns.Add("Score", typeof(int));
            south.Rows.Add("Beta", 20);
            dataSet.Tables.Add(north);
            dataSet.Tables.Add(south);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.InsertDataSet(dataSet);
                var sheet = CreateDeferredSheetHandle(document, "South");
                sheet.ColumnStyleByHeader("Score").NumberFormat("0.00");
                typeof(ExcelDocument).GetField("_directDataSetMetadataSourceSheet", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(document, null);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var workbookSheet = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>()
                .Single(sheet => string.Equals(sheet.Name?.Value, "South", StringComparison.Ordinal));
            var worksheetPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(workbookSheet.Id!);
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Beta", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("0.00", GetCellNumberFormatCode(spreadsheet, cells["B2"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_HeaderNumberFormatIncludeHeaderFormatsHeaderCell() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19, 8, 30, 0)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20, 9, 45, 0))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.ColumnStyleByHeader("Score", includeHeader: true).NumberFormat("#,##0.000");

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("#,##0.000", GetCellNumberFormatCode(spreadsheet, cells["B1"]));
            Assert.Equal("#,##0.000", GetCellNumberFormatCode(spreadsheet, cells["B2"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_HeaderNumberFormatWorksAfterChartPreservesFastSaveModel() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("North", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("South", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("East", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C4", hasHeader: true, name: "ReportData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                var chartData = new ExcelChartData(
                    rows.Select(row => row.Name),
                    new[] { new ExcelChartSeries("Score", rows.Select(row => (double)row.Score)) });
                sheet.AddChart(chartData, row: 2, column: 6, widthPixels: 480, heightPixels: 280, type: ExcelChartType.ColumnClustered, title: "Scores");
                sheet.ColumnStyleByHeader("Score").NumberFormat("#,##0.000");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart != null);
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("#,##0.000", GetCellNumberFormatCode(spreadsheet, cells["B2"]));
            Assert.Equal("#,##0.000", GetCellNumberFormatCode(spreadsheet, cells["B4"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataSet_HeaderNumberFormatPreservesSourceInvalidationSubscription() {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Original", 10);
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.InsertDataSet(dataSet);
                var sheet = document.Sheets.First(item => string.Equals(item.Name, "Items", StringComparison.Ordinal));
                sheet.ColumnStyleByHeader("Score").NumberFormat("0.00");
                table.Rows.Add("Late", 20);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Original", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.False(cells.ContainsKey("A3"));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_TableAutoFilterCriteriaMaterializesAndPersistsCriteria() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C3", hasHeader: true, name: "FilteredReport", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AddAutoFilter("A1:C3", new Dictionary<uint, IEnumerable<string>> {
                    { 0U, new[] { "Alpha" } }
                });

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var tablePart = Assert.Single(worksheetPart.TableDefinitionParts);
            var tableFilter = Assert.Single(tablePart.Table!.Elements<AutoFilter>());
            var filterColumn = Assert.Single(tableFilter.Elements<FilterColumn>());
            Assert.Equal(0U, filterColumn.ColumnId!.Value);
            Assert.Equal("Alpha", filterColumn.GetFirstChild<Filters>()!.Elements<Filter>().Single().Val!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_TableAutoFilterFastPathUpdatesDeferredModel() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C3", hasHeader: true, name: "FilteredReport", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: false);
                sheet.AddAutoFilter("A1:C3");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var tablePart = Assert.Single(worksheetPart.TableDefinitionParts);
            var tableFilter = Assert.Single(tablePart.Table!.Elements<AutoFilter>());
            Assert.Equal("A1:C3", tableFilter.Reference!.Value);
            Assert.Empty(tableFilter.ChildElements);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_MaterializationRestoresCapturedSheetFormatProperties() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                var worksheet = sheet.WorksheetPart.Worksheet;
                worksheet.InsertBefore(new SheetFormatProperties {
                    DefaultRowHeight = 21D,
                    DefaultColumnWidth = 14D,
                    CustomHeight = true
                }, worksheet.GetFirstChild<SheetData>());

                sheet.CellValue(5, 1, "forces materialization");
                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var sheetFormat = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetFormatProperties>()!;
            Assert.Equal(21D, sheetFormat.DefaultRowHeight!.Value);
            Assert.Equal(14D, sheetFormat.DefaultColumnWidth!.Value);
            Assert.True(sheetFormat.CustomHeight!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_MaterializationRestoresCapturedPrintLayoutMetadata() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.SetMargins(0.25D, 0.25D, 0.5D, 0.5D, 0.3D, 0.3D);
                sheet.SetOrientation(ExcelPageOrientation.Landscape);
                sheet.SetPageSetup(fitToWidth: 1U, fitToHeight: 0U);

                sheet.CellValue(5, 1, "forces materialization");
                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
            var margins = worksheet.GetFirstChild<PageMargins>()!;
            var setup = worksheet.GetFirstChild<PageSetup>()!;
            Assert.Equal(0.25D, margins.Left!.Value);
            Assert.Equal(0.5D, margins.Top!.Value);
            Assert.Equal(OrientationValues.Landscape, setup.Orientation!.Value);
            Assert.Equal(1U, setup.FitToWidth!.Value);
            Assert.Equal(0U, setup.FitToHeight!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PivotTableAfterDeferredImportKeepsSourceRowsAndSharedItems() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("Alpha", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddPivotTable(
                    sourceRange: "A1:C4",
                    destinationCell: "E2",
                    name: "ScorePivot",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("30", cells["B4"].CellValue!.Text);

            var pivotPart = Assert.Single(worksheetPart.PivotTableParts);
            Assert.Equal("ScorePivot", pivotPart.PivotTableDefinition!.Name!.Value);
            var cacheDefinition = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!;
            Assert.False(cacheDefinition.SaveData!.Value);
            Assert.True(cacheDefinition.RefreshOnLoad!.Value);
            var cacheRecordsPart = Assert.Single(pivotPart.PivotTableCacheDefinitionPart!.GetPartsOfType<PivotTableCacheRecordsPart>());
            Assert.Equal(0U, cacheRecordsPart.PivotCacheRecords!.Count!.Value);
            var cacheFields = cacheDefinition.CacheFields!.Elements<CacheField>().ToList();
            var nameItems = cacheFields[0].SharedItems!.Elements<StringItem>().Select(item => item.Val!.Value).ToList();
            Assert.Equal(new[] { "Alpha", "Beta" }, nameItems);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PivotTableFallbackAfterDeferredImportKeepsSourceRows() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("Gamma", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddPivotTable(
                    sourceRange: "A1:C4",
                    destinationCell: "E2",
                    name: "ScorePivot",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                var relationship = sheet.WorksheetPart.AddHyperlinkRelationship(new Uri("https://example.org/pivot-fallback"), true);
                var hyperlinks = sheet.WorksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
                if (hyperlinks == null) {
                    hyperlinks = new Hyperlinks();
                    InsertHyperlinksBeforeWorksheetTail(sheet.WorksheetPart.Worksheet, hyperlinks);
                }

                hyperlinks.Append(new Hyperlink { Reference = "D7", Id = relationship.Id });

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("Gamma", GetSpreadsheetCellText(spreadsheet, cells["A4"]));
            Assert.Single(worksheetPart.HyperlinkRelationships);
            var savedHyperlinks = Assert.Single(worksheetPart.Worksheet.Elements<Hyperlinks>());
            Assert.Single(savedHyperlinks.Elements<Hyperlink>());
            Assert.Single(worksheetPart.PivotTableParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PivotCacheRawXmlPreservesNonBmpTextWhenEmbeddingRequired() {
            using var memory = new MemoryStream();
            string emoji = char.ConvertFromUtf32(0x1F600);
            string rocket = char.ConvertFromUtf32(0x1F680);
            var rows = new[] {
                new NullablePerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19), "Emoji " + emoji),
                new NullablePerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20), "Rocket " + rocket)
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Note", row => row.Note));
                sheet.AddPivotTable(
                    sourceRange: "A1:C3",
                    destinationCell: "E2",
                    name: "ScorePivot",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") },
                    calculatedFields: new[] { new ExcelPivotCalculatedField("DoubleScore", "'Score' * 2") },
                    options: new ExcelPivotTableOptions { SaveSourceData = true });

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var pivotPart = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts.First().PivotTableParts);
            var cacheDefinition = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!;
            Assert.True(cacheDefinition.SaveData!.Value);
            Assert.False(cacheDefinition.RefreshOnLoad!.Value);
            var cacheRecordsPart = Assert.Single(pivotPart.PivotTableCacheDefinitionPart!.GetPartsOfType<PivotTableCacheRecordsPart>());
            using var reader = new StreamReader(cacheRecordsPart.GetStream(FileMode.Open, FileAccess.Read));
            string cacheRecordsXml = reader.ReadToEnd();

            Assert.Contains("Emoji " + emoji, cacheRecordsXml);
            Assert.Contains("Rocket " + rocket, cacheRecordsXml);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_LargeSimplePivotOmitsCacheRecords() {
            using var memory = new MemoryStream();
            var rows = Enumerable.Range(1, 5000)
                .Select(index => new PerformanceObjectExportRow(
                    index % 2 == 0 ? "Alpha" : "Beta",
                    index,
                    new DateTime(2026, 5, 1).AddDays(index % 28)))
                .ToArray();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddPivotTable(
                    sourceRange: "A1:C5001",
                    destinationCell: "E2",
                    name: "ScorePivot",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A3"]));
            Assert.Equal("5000", cells["B5001"].CellValue!.Text);

            var pivotPart = Assert.Single(worksheetPart.PivotTableParts);
            var cacheDefinition = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!;
            Assert.False(cacheDefinition.SaveData!.Value);
            Assert.True(cacheDefinition.RefreshOnLoad!.Value);
            var cacheRecordsPart = Assert.Single(pivotPart.PivotTableCacheDefinitionPart!.GetPartsOfType<PivotTableCacheRecordsPart>());
            Assert.Equal(0U, cacheRecordsPart.PivotCacheRecords!.Count!.Value);
            var cacheFields = cacheDefinition.CacheFields!.Elements<CacheField>().ToList();
            var nameItems = cacheFields[0].SharedItems!.Elements<StringItem>().Select(item => item.Val!.Value).ToList();
            Assert.Equal(new[] { "Beta", "Alpha" }, nameItems);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_OverlayNumericValuesPreserveLexicalText() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.CellValue(2, 4, 0);
                sheet.CellValue(2, 5, 0);

                var cells = sheet.WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                cells["D2"].DataType = CellValues.Number;
                cells["D2"].CellValue = new CellValue("1234567890123456789");
                cells["E2"].DataType = null;
                cells["E2"].CellValue = new CellValue("1234567890.123456789");

                sheet.AddPivotTable(
                    sourceRange: "A1:C3",
                    destinationCell: "G2",
                    name: "ScorePivot",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                document.MaterializeDeferredDataSetImport();
                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var savedCells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

            Assert.Equal("1234567890123456789", savedCells["D2"].CellValue!.Text);
            Assert.Equal(CellValues.Number, savedCells["D2"].DataType!.Value);
            Assert.Equal("1234567890.123456789", savedCells["E2"].CellValue!.Text);
            Assert.True(savedCells["E2"].DataType == null || savedCells["E2"].DataType!.Value == CellValues.Number);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PivotTableDeferredSharedItemsPreserveNativeTypes() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("Gamma", 30, new DateTime(2026, 5, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddPivotTable(
                    sourceRange: "A1:C4",
                    destinationCell: "E2",
                    name: "ScorePivot",
                    rowFields: new[] { "Score" },
                    dataFields: new[] { new ExcelPivotDataField("Name", DataConsolidateFunctionValues.Count, "Name Count") });

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var pivotPart = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts.First().PivotTableParts);
            var cacheFields = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ToList();
            var scoreSharedItems = cacheFields[1].SharedItems!;
            Assert.True(scoreSharedItems.ContainsNumber!.Value);
            Assert.Equal(new[] { 10D, 20D, 30D }, scoreSharedItems.Elements<NumberItem>().Select(item => item.Val!.Value).ToArray());
            Assert.Empty(scoreSharedItems.Elements<StringItem>());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PivotTableDeferredDateHierarchyMaterializesBeforeGrouping() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20)),
                new PerformanceObjectExportRow("Gamma", 30, new DateTime(2026, 6, 21))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddPivotTable(
                    sourceRange: "A1:C4",
                    destinationCell: "E2",
                    name: "CreatedPivot",
                    rowFields: new[] { "Created" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") },
                    groupings: new[] { ExcelPivotGrouping.DateHierarchy("Created", GroupByValues.Years, GroupByValues.Months) });

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("30", cells["B4"].CellValue!.Text);

            var pivotPart = Assert.Single(worksheetPart.PivotTableParts);
            var cacheFields = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ToList();
            Assert.Equal("Created Years", cacheFields[3].Name!.Value);
            Assert.Equal("Created Months", cacheFields[4].Name!.Value);
            Assert.Contains(cacheFields[3].FieldGroup!.GetFirstChild<GroupItems>()!.Elements<StringItem>(), item => item.Val!.Value == "2026");
            Assert.Contains(cacheFields[4].FieldGroup!.GetFirstChild<GroupItems>()!.Elements<StringItem>(), item => item.Val!.Value == "May");
            Assert.Contains(cacheFields[4].FieldGroup!.GetFirstChild<GroupItems>()!.Elements<StringItem>(), item => item.Val!.Value == "June");
            Assert.Contains(cacheFields[3].SharedItems!.Elements<StringItem>(), item => item.Val!.Value == "2026");
            Assert.Contains(cacheFields[4].SharedItems!.Elements<StringItem>(), item => item.Val!.Value == "May");
            Assert.Contains(cacheFields[4].SharedItems!.Elements<StringItem>(), item => item.Val!.Value == "June");
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_WorksheetRangePivotTableDateSharedItemsPreserveDateType() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Created");
                sheet.CellValue(1, 2, "Score");
                sheet.CellValue(2, 1, new DateTime(2026, 5, 19));
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(3, 1, new DateTime(2026, 5, 20));
                sheet.CellValue(3, 2, 20);

                sheet.AddPivotTable(
                    sourceRange: "A1:B3",
                    destinationCell: "D2",
                    name: "CreatedPivot",
                    rowFields: new[] { "Created" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var pivotPart = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts.First().PivotTableParts);
            var cacheFields = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ToList();
            var createdSharedItems = cacheFields[0].SharedItems!;
            Assert.True(createdSharedItems.ContainsDate!.Value);
            Assert.Equal(new[] { new DateTime(2026, 5, 19), new DateTime(2026, 5, 20) }, createdSharedItems.Elements<DateTimeItem>().Select(item => item.Val!.Value).ToArray());
            Assert.Empty(createdSharedItems.Elements<NumberItem>());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_PivotTableSkipsUnusedSharedItems() {
            using var memory = new MemoryStream();
            var rows = Enumerable.Range(1, 50)
                .Select(index => new PerformanceObjectExportRow(index % 2 == 0 ? "Alpha" : "Beta", index * 10, new DateTime(2026, 5, 1).AddDays(index)))
                .ToArray();

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddPivotTable(
                    sourceRange: "A1:C51",
                    destinationCell: "E2",
                    name: "ScorePivot",
                    rowFields: new[] { "Name" },
                    dataFields: new[] { new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Sum, "Total Score") });

                document.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var pivotPart = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts.First().PivotTableParts);
            var cacheFields = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ToList();
            Assert.Equal(new[] { "Alpha", "Beta" }, cacheFields[0].SharedItems!.Elements<StringItem>().Select(item => item.Val!.Value).OrderBy(item => item, StringComparer.Ordinal).ToArray());
            Assert.Empty(cacheFields[1].SharedItems!.ChildElements);
            Assert.Empty(cacheFields[2].SharedItems!.ChildElements);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjects_DuplicateExplicitHeadersFallBackBeforeTablePromotion() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Name", row => row.Score));
                sheet.AddTable("A1:B2", hasHeader: true, name: "DuplicateHeaders", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            Assert.Equal("A1:B2", worksheetPart.TableDefinitionParts.Single().Table!.Reference!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertObjectsThenAddTable_ExternalMutationPreservesDirectPackageCandidate() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score),
                    ("Created", row => row.Created));
                sheet.AddTable("A1:C2", hasHeader: true, name: "ObjectSales", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.CellValue(4, 1, "Manual edit");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
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
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Empty(worksheetPart.TableDefinitionParts);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_FluentRowsFrom_NullableSimpleRowsUseDirectPackage() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new NullablePerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19), string.Empty),
                new NullablePerformanceObjectExportRow(null, null, null, "Manual")
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
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
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Equal(string.Empty, GetSpreadsheetCellText(spreadsheet, cells["A3"]));
            Assert.Equal(string.Empty, GetSpreadsheetCellText(spreadsheet, cells["B3"]));
            Assert.Equal(string.Empty, GetSpreadsheetCellText(spreadsheet, cells["C3"]));
            Assert.Equal("Manual", GetSpreadsheetCellText(spreadsheet, cells["D3"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_FluentRowsFrom_RepeatedSimpleTypeUsesDirectPackage() {
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19))
            };

            using var first = new MemoryStream();
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.AsFluent()
                    .Sheet("First", sheet => sheet.RowsFrom(rows))
                    .End()
                    .Save(first);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            using var second = new MemoryStream();
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.AsFluent()
                    .Sheet("Second", sheet => sheet.RowsFrom(rows))
                    .End()
                    .Save(second);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }
        }

        [Fact]
        public void PerformanceReview_FluentRowsFrom_ReadOnlyListUsesDirectPackageWithoutSnapshotEnumeration() {
            var rows = new ThrowOnEnumerateReadOnlyList<PerformanceObjectExportRow>(
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20)));

            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.AsFluent()
                    .Sheet("Data", sheet => sheet.RowsFrom(rows))
                    .End()
                    .Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }
        }

        [Fact]
        public void PerformanceReview_FluentRowsFrom_DirtyWorkbookPreservesSimpleRows() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.AsFluent()
                    .Sheet("Data", sheet => sheet
                        .Cell(5, 5, "Manual edit")
                        .RowsFrom(rows))
                    .End()
                    .Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Score", GetSpreadsheetCellText(spreadsheet, cells["B1"]));
            Assert.Equal("Created", GetSpreadsheetCellText(spreadsheet, cells["C1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.Equal(1U, cells["C2"].StyleIndex!.Value);
            Assert.Equal("Beta", GetSpreadsheetCellText(spreadsheet, cells["A3"]));
            Assert.Equal("Manual edit", GetSpreadsheetCellText(spreadsheet, cells["E5"]));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_FluentRowsFromThenTable_UsesDirectPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var rows = new[] {
                new PerformanceObjectExportRow("Alpha", 10, new DateTime(2026, 5, 19)),
                new PerformanceObjectExportRow("Beta", 20, new DateTime(2026, 5, 20))
            };

            using (var document = ExcelDocument.Create(new MemoryStream())) {
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.AsFluent()
                    .Sheet("Data", sheet => sheet.RowsFrom(rows))
                    .End();
                document.Sheets[0].CellValue(5, 1, "Manual edit");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataTable(table);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Name", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
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
        public void PerformanceReview_InsertDataTable_ObjectColumnDateAndTimeValuesKeepStylesWithDirectPackage() {
            using var memory = new MemoryStream();
            var table = new DataTable("Mixed");
            table.Columns.Add("Kind", typeof(string));
            table.Columns.Add("Value", typeof(object));
            table.Rows.Add("When", new DateTime(2026, 5, 19, 8, 30, 0));
            table.Rows.Add("Duration", TimeSpan.FromMinutes(95));

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataTable(table);
                sheet.CellValue(5, 1, "Manual edit");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
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
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.False(cells.ContainsKey("A3"));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTable_ExternalMutationPreservesDirectPackageCandidate() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataTable(table);
                sheet.CellValue(5, 1, "Manual edit");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(5, 1, out string? text));
            Assert.Equal("Manual edit", text);
        }

        [Fact]
        public void PerformanceReview_InsertDataTable_HiddenSheetSkipsDirectPackageAndPreservesState() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("Alpha");

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
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
        public void PerformanceReview_InsertDataTableAsTable_SourceMutationAfterInsertDoesNotChangeDirectCandidate() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                Assert.Equal("A1:B2", sheet.InsertDataTableAsTable(table, tableName: "Sales Table"));
                table.Rows[0]["Name"] = "Changed";
                table.Rows.Add("Late", 20);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.False(cells.ContainsKey("A3"));
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("A1:B2", tableDefinition.Reference!.Value);
            Assert.Equal("Sales_Table", tableDefinition.Name!.Value);
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                Assert.Equal("A1:C3", sheet.InsertDataTableAsTable(table, tableName: "Sales Table", style: OfficeIMO.Excel.TableStyle.TableStyleMedium9));

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                Assert.Equal("A1:B1", sheet.InsertDataTableAsTable(table, includeHeaders: false, tableName: "HeaderlessSales"));

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
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("0", tableDefinition.HeaderRowCount!.Value.ToString(CultureInfo.InvariantCulture));
            var columns = tableDefinition.TableColumns!.Elements<TableColumn>().ToList();
            Assert.Equal("Column1", columns[0].Name!.Value);
            Assert.Equal("Column2", columns[1].Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataTableAsTable_ExternalMutationPreservesDirectPackageCandidate() {
            using var memory = new MemoryStream();
            var table = new DataTable("Sales");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");
                sheet.CellValue(4, 1, "Manual edit");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
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

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                document.BuiltinDocumentProperties.Title = "Sales Export";
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataTableAsTable(table, tableName: "SalesTable");

                document.Save(memory);

                Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.Equal("Sales Export", loaded.BuiltinDocumentProperties.Title);
            Assert.True(loaded.Sheets[0].TryGetCellText(2, 1, out string? text));
            Assert.Equal("Alpha", text);
        }

        [Fact]
        public void PerformanceReview_InsertDataReader_UsesDirectDataSetPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Columns.Add("Created", typeof(DateTime));
            table.Rows.Add("Alpha", 10, new DateTime(2026, 5, 19));
            table.Rows.Add("Beta", 20, new DateTime(2026, 5, 20));
            table.Rows.Add("Gamma", DBNull.Value, DBNull.Value);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                using IDataReader reader = table.CreateDataReader();
                Assert.Equal("A1:C4", sheet.InsertDataReader(reader, tableName: "Reader Table", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4));

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("10", cells["B2"].CellValue!.Text);
            Assert.True(cells["C2"].StyleIndex?.Value > 0U);
            Assert.Equal("Gamma", GetSpreadsheetCellText(spreadsheet, cells["A4"]));
            Assert.Equal(string.Empty, cells["B4"].CellValue!.Text);
            Assert.Equal(string.Empty, cells["C4"].CellValue!.Text);
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("A1:C4", tableDefinition.Reference!.Value);
            Assert.Equal("Reader_Table", tableDefinition.Name!.Value);
            Assert.Equal("TableStyleMedium4", tableDefinition.TableStyleInfo!.Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataReader_AutoFitUsesDirectDataSetPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Description", typeof(string));
            table.Rows.Add("Alpha", "A longer value for sizing");

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                using IDataReader reader = table.CreateDataReader();
                sheet.InsertDataReader(reader, tableName: "ReaderAutoFit", autoFit: true);

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
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
        public void PerformanceReview_InsertDataReader_ExternalMutationPreservesDirectPackageCandidate() {
            using var memory = new MemoryStream();
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                using IDataReader reader = table.CreateDataReader();
                sheet.InsertDataReader(reader, tableName: "ReaderTable");
                sheet.CellValue(5, 1, "Manual edit");

                document.Save(memory);

                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var loaded = ExcelDocument.Load(memory, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.True(loaded.Sheets[0].TryGetCellText(5, 1, out string? text));
            Assert.Equal("Manual edit", text);
        }

        [Fact]
        public void PerformanceReview_InsertDataReader_HeaderlessTableUsesDirectDataSetPackageWhenWorkbookIsClean() {
            using var memory = new MemoryStream();
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                using IDataReader reader = table.CreateDataReader();
                Assert.Equal("A1:B1", sheet.InsertDataReader(reader, includeHeaders: false, tableName: "HeaderlessReader"));

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
            var tableDefinition = worksheetPart.TableDefinitionParts.Single().Table!;
            Assert.Equal("0", tableDefinition.HeaderRowCount!.Value.ToString(CultureInfo.InvariantCulture));
            var columns = tableDefinition.TableColumns!.Elements<TableColumn>().ToList();
            Assert.Equal("Column1", columns[0].Name!.Value);
            Assert.Equal("Column2", columns[1].Name!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void PerformanceReview_InsertDataReader_UsesBulkValueReadsWhenAvailable() {
            using var memory = new MemoryStream();
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Columns.Add("Note", typeof(string));
            table.Rows.Add("Alpha", 10, "Ready");
            table.Rows.Add("Beta", 20, DBNull.Value);

            using var reader = new CountingDataReader(table.CreateDataReader());
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataReader(reader, tableName: "ReaderTable");
                document.Save(memory);
            }

            Assert.Equal(2, reader.GetValuesCalls);
            Assert.Equal(0, reader.GetValueCalls);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal(string.Empty, cells["C3"].CellValue!.Text);
        }

        [Fact]
        public void PerformanceReview_InsertDataReader_FallsBackWhenBulkValueReadsAreUnsupported() {
            using var memory = new MemoryStream();
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);
            table.Rows.Add("Beta", 20);

            using var reader = new CountingDataReader(table.CreateDataReader(), throwOnGetValues: true);
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                var sheet = document.AddWorksheet("Data");
                sheet.InsertDataReader(reader, tableName: "ReaderTable");
                document.Save(memory);
            }

            Assert.Equal(1, reader.GetValuesCalls);
            Assert.Equal(4, reader.GetValueCalls);

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Beta", GetSpreadsheetCellText(spreadsheet, cells["A3"]));
            Assert.Equal("20", cells["B3"].CellValue!.Text);
        }

        [Fact]
        public void PerformanceReview_InsertDataReader_FailedReadDoesNotPartiallyAppendRows() {
            var table = new DataTable("ReaderData");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("Alpha", 10);
            table.Rows.Add("Beta", 20);

            using var reader = new CountingDataReader(table.CreateDataReader(), throwOnReadAfterRows: 1);
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddWorksheet("Data");

            var exception = Assert.Throws<InvalidOperationException>(() => sheet.InsertDataReader(reader, tableName: "ReaderTable"));
            Assert.Contains("Simulated reader failure", exception.Message);

            var sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            Assert.Empty(sheetData.Elements<Row>());
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

        private sealed class NullablePerformanceObjectExportRow {
            public NullablePerformanceObjectExportRow(string? name, int? score, DateTime? created, string note) {
                Name = name;
                Score = score;
                Created = created;
                Note = note;
            }

            public string? Name { get; }

            public int? Score { get; }

            public DateTime? Created { get; }

            public string Note { get; }
        }

        private sealed class MutablePerformanceObjectExportRow {
            public string Name { get; set; } = string.Empty;

            public int Score { get; set; }
        }

#if NET6_0_OR_GREATER
        private sealed class ExtendedPerformanceObjectExportRow {
            public ExtendedPerformanceObjectExportRow(string name, PerformanceExportStatus status, TimeSpan duration, DateOnly localDate, TimeOnly localTime) {
                Name = name;
                Status = status;
                Duration = duration;
                LocalDate = localDate;
                LocalTime = localTime;
            }

            public string Name { get; }

            public PerformanceExportStatus Status { get; }

            public TimeSpan Duration { get; }

            public DateOnly LocalDate { get; }

            public TimeOnly LocalTime { get; }
        }

        private enum PerformanceExportStatus {
            Ready,
            Waiting
        }
#endif

        private sealed class ThrowOnEnumerateReadOnlyList<T> : IReadOnlyList<T> {
            private readonly T[] _items;

            internal ThrowOnEnumerateReadOnlyList(params T[] items) {
                _items = items;
            }

            public int Count => _items.Length;

            public T this[int index] => _items[index];

            public IEnumerator<T> GetEnumerator() => throw new InvalidOperationException("RowsFrom direct-save path should use IReadOnlyList<T> indexing without snapshot enumeration.");

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }

        private sealed class CountingDataReader : IDataReader {
            private readonly IDataReader _inner;
            private readonly bool _throwOnGetValues;
            private readonly int _throwOnReadAfterRows;
            private int _successfulReads;

            internal CountingDataReader(IDataReader inner, bool throwOnGetValues = false, int throwOnReadAfterRows = -1) {
                _inner = inner;
                _throwOnGetValues = throwOnGetValues;
                _throwOnReadAfterRows = throwOnReadAfterRows;
            }

            internal int GetValuesCalls { get; private set; }

            internal int GetValueCalls { get; private set; }

            public object this[int i] => _inner[i];

            public object this[string name] => _inner[name];

            public int Depth => _inner.Depth;

            public bool IsClosed => _inner.IsClosed;

            public int RecordsAffected => _inner.RecordsAffected;

            public int FieldCount => _inner.FieldCount;

            public void Close() => _inner.Close();

            public void Dispose() => _inner.Dispose();

            public bool GetBoolean(int i) => _inner.GetBoolean(i);

            public byte GetByte(int i) => _inner.GetByte(i);

            public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) => _inner.GetBytes(i, fieldOffset, buffer, bufferoffset, length);

            public char GetChar(int i) => _inner.GetChar(i);

            public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length) => _inner.GetChars(i, fieldoffset, buffer, bufferoffset, length);

            public IDataReader GetData(int i) => _inner.GetData(i);

            public string GetDataTypeName(int i) => _inner.GetDataTypeName(i);

            public DateTime GetDateTime(int i) => _inner.GetDateTime(i);

            public decimal GetDecimal(int i) => _inner.GetDecimal(i);

            public double GetDouble(int i) => _inner.GetDouble(i);

            public Type GetFieldType(int i) => _inner.GetFieldType(i);

            public float GetFloat(int i) => _inner.GetFloat(i);

            public Guid GetGuid(int i) => _inner.GetGuid(i);

            public short GetInt16(int i) => _inner.GetInt16(i);

            public int GetInt32(int i) => _inner.GetInt32(i);

            public long GetInt64(int i) => _inner.GetInt64(i);

            public string GetName(int i) => _inner.GetName(i);

            public int GetOrdinal(string name) => _inner.GetOrdinal(name);

            public DataTable? GetSchemaTable() => _inner.GetSchemaTable();

            public string GetString(int i) => _inner.GetString(i);

            public object GetValue(int i) {
                GetValueCalls++;
                return _inner.GetValue(i);
            }

            public int GetValues(object[] values) {
                GetValuesCalls++;
                if (_throwOnGetValues) {
                    throw new NotSupportedException();
                }

                return _inner.GetValues(values);
            }

            public bool IsDBNull(int i) => _inner.IsDBNull(i);

            public bool NextResult() => _inner.NextResult();

            public bool Read() {
                if (_throwOnReadAfterRows >= 0 && _successfulReads >= _throwOnReadAfterRows) {
                    throw new InvalidOperationException("Simulated reader failure.");
                }

                bool read = _inner.Read();
                if (read) {
                    _successfulReads++;
                }

                return read;
            }
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

        private static DataTable CreateSingleColumnDataTable(string name, params string[] values) {
            var table = new DataTable(name);
            table.Columns.Add("Name", typeof(string));
            foreach (string value in values) {
                table.Rows.Add(value);
            }

            return table;
        }

        private static ExcelSheet CreateDeferredSheetHandle(ExcelDocument document, string sheetName) {
            var workbookPart = document._spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null.");
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            string relationshipId = workbookPart.GetIdOfPart(worksheetPart);
            return new ExcelSheet(
                document,
                document._spreadSheetDocument,
                new Sheet { Name = sheetName, Id = relationshipId, SheetId = 9999U });
        }

        private static uint AddBoldCellStyle(SpreadsheetDocument document) {
            var stylesPart = document.WorkbookPart!.WorkbookStylesPart ?? document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet ??= new Stylesheet(
                new Fonts(new Font()),
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(new Border()),
                new CellFormats(new CellFormat()));

            var stylesheet = stylesPart.Stylesheet;
            stylesheet.Fonts ??= new Fonts(new Font());
            stylesheet.Fills ??= new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
            stylesheet.Borders ??= new Borders(new Border());
            stylesheet.CellFormats ??= new CellFormats(new CellFormat());

            uint fontId = (uint)stylesheet.Fonts.Elements<Font>().Count();
            stylesheet.Fonts.Append(new Font(new Bold()));
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Elements<Font>().Count();

            uint styleIndex = (uint)stylesheet.CellFormats.Elements<CellFormat>().Count();
            stylesheet.CellFormats.Append(new CellFormat {
                FontId = fontId,
                FillId = 0U,
                BorderId = 0U,
                ApplyFont = true
            });
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Elements<CellFormat>().Count();
            stylesheet.Save();
            return styleIndex;
        }

        private static void AssertRichOverlayStylePersisted(MemoryStream memory) {
            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts
                .Single(part => part.Worksheet.Descendants<Cell>().Any(cell => string.Equals(cell.CellReference?.Value, "D2", StringComparison.Ordinal)));
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("North", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("Styled side note", GetSpreadsheetCellText(spreadsheet, cells["D2"]));
            Assert.NotNull(cells["D2"].StyleIndex);
            var stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            var cellFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)cells["D2"].StyleIndex!.Value);
            Assert.NotNull(cellFormat.FontId);
            Assert.NotEqual(0U, cellFormat.FontId!.Value);
            var font = stylesheet.Fonts!.Elements<Font>().ElementAt((int)cellFormat.FontId.Value);
            Assert.NotNull(font.Bold);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        private static void InsertHyperlinksBeforeWorksheetTail(Worksheet worksheet, Hyperlinks hyperlinks) {
            OpenXmlElement? insertBefore =
                worksheet.Elements<PrintOptions>().Cast<OpenXmlElement>()
                    .Concat(worksheet.Elements<PageMargins>())
                    .Concat(worksheet.Elements<PageSetup>())
                    .Concat(worksheet.Elements<HeaderFooter>())
                    .Concat(worksheet.Elements<RowBreaks>())
                    .Concat(worksheet.Elements<ColumnBreaks>())
                    .Concat(worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Drawing>())
                    .Concat(worksheet.Elements<LegacyDrawing>())
                    .Concat(worksheet.Elements<TableParts>())
                    .FirstOrDefault();

            if (insertBefore == null) {
                worksheet.Append(hyperlinks);
            } else {
                worksheet.InsertBefore(hyperlinks, insertBefore);
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
            string? customFormat = stylesheet.NumberingFormats?.Elements<NumberingFormat>()
                .FirstOrDefault(format => format.NumberFormatId?.Value == numberFormatId)
                ?.FormatCode
                ?.Value;
            return customFormat ?? numberFormatId switch {
                1U => "0",
                2U => "0.00",
                3U => "#,##0",
                4U => "#,##0.00",
                9U => "0%",
                10U => "0.00%",
                11U => "0.00E+00",
                12U => "# ?/?",
                13U => "# ??/??",
                14U => "mm-dd-yy",
                15U => "d-mmm-yy",
                16U => "d-mmm",
                17U => "mmm-yy",
                18U => "h:mm AM/PM",
                19U => "h:mm:ss AM/PM",
                20U => "h:mm",
                21U => "h:mm:ss",
                22U => "m/d/yy h:mm",
                37U => "#,##0 ;(#,##0)",
                38U => "#,##0 ;[Red](#,##0)",
                39U => "#,##0.00;(#,##0.00)",
                40U => "#,##0.00;[Red](#,##0.00)",
                45U => "mm:ss",
                46U => "[h]:mm:ss",
                47U => "mmss.0",
                48U => "##0.0E+0",
                49U => "@",
                _ => null
            };
        }

        private static string? GetSpreadsheetCellText(SpreadsheetDocument spreadsheet, Cell cell) {
            string? value = cell.CellValue?.Text;
            if (cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                return cell.InlineString?.InnerText ?? string.Empty;
            }

            if (cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sharedStringIndex)) {
                return spreadsheet.WorkbookPart?.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sharedStringIndex)?.InnerText;
            }

            return value;
        }

        private sealed class NonSeekableWriteStream : Stream {
            private readonly MemoryStream _inner = new();

            public override bool CanRead => false;

            public override bool CanSeek => false;

            public override bool CanWrite => true;

            public override long Length => throw new NotSupportedException();

            public override long Position {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public byte[] ToArray() => _inner.ToArray();

            public override void Flush() => _inner.Flush();

            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
        }
    }
}

namespace System.Management.Automation {
    public sealed class PSObject {
        public PSObject(params PSPropertyInfo[] properties) {
            Properties = properties;
        }

        public IReadOnlyList<PSPropertyInfo> Properties { get; }
    }

    public sealed class PSPropertyInfo {
        public PSPropertyInfo(string name, object? value, bool isGettable = true) {
            Name = name;
            Value = value;
            IsGettable = isGettable;
        }

        public string Name { get; }

        public object? Value { get; }

        public bool IsGettable { get; }
    }
}
