using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using System;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Reader_ReadRangeAsDataReader_ExposesTypedSchemaAndValues() {
            var expectedDate = new DateTime(2026, 7, 8);
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(1, 3, "Created");
                sheet.CellValue(1, 4, "Active");
                sheet.CellValue(1, 5, "Note");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, 12.5d);
                sheet.CellValue(2, 3, expectedDate);
                sheet.CellValue(2, 4, true);
                sheet.CellValue(2, 5, "Alpha");
                sheet.CellValue(3, 1, 2);
                sheet.CellValue(3, 4, false);
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:E3", chunkRows: 1, schemaSampleRows: 2);

            Assert.True(Assert.IsAssignableFrom<DbDataReader>(dataReader).HasRows);
            Assert.Equal(5, dataReader.FieldCount);
            Assert.Equal(typeof(double), dataReader.GetFieldType(dataReader.GetOrdinal("Id")));
            Assert.Equal(typeof(double), dataReader.GetFieldType(dataReader.GetOrdinal("Amount")));
            Assert.Equal(typeof(DateTime), dataReader.GetFieldType(dataReader.GetOrdinal("Created")));
            Assert.Equal(typeof(bool), dataReader.GetFieldType(dataReader.GetOrdinal("Active")));
            Assert.Equal(typeof(string), dataReader.GetFieldType(dataReader.GetOrdinal("Note")));

            Assert.True(dataReader.Read());
            Assert.Equal(1d, dataReader.GetDouble(dataReader.GetOrdinal("Id")));
            Assert.Equal(12.5d, dataReader.GetDouble(dataReader.GetOrdinal("Amount")));
            Assert.Equal(expectedDate, dataReader.GetDateTime(dataReader.GetOrdinal("Created")));
            Assert.True(dataReader.GetBoolean(dataReader.GetOrdinal("Active")));
            Assert.Equal("Alpha", dataReader.GetString(dataReader.GetOrdinal("Note")));

            Assert.True(dataReader.Read());
            Assert.Equal(2d, dataReader.GetDouble(dataReader.GetOrdinal("Id")));
            Assert.True(dataReader.IsDBNull(dataReader.GetOrdinal("Amount")));
            Assert.False(dataReader.GetBoolean(dataReader.GetOrdinal("Active")));
            Assert.False(dataReader.Read());
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_LoadsDataTableWithDisambiguatedHeaders() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(1, 2, "Status");
                sheet.CellValue(1, 3, "");
                sheet.CellValue(2, 1, "OK");
                sheet.CellValue(2, 2, "Warning");
                sheet.CellValue(2, 3, "Generated");
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:C2");
            var table = new DataTable();
            table.Load(dataReader);

            Assert.Equal(new[] { "Status", "Status_2", "Column3" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
            DataRow row = Assert.Single(table.Rows.Cast<DataRow>());
            Assert.Equal("OK", row["Status"]);
            Assert.Equal("Warning", row["Status_2"]);
            Assert.Equal("Generated", row["Column3"]);
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_RejectsRangesBeyondBufferedCellBudget() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                document.AddWorksheet("Data").CellValue(1, 1, "Header");
            }
            using var reader = ExcelDocumentReader.Open(memory.ToArray(), new ExcelReadOptions {
                MaxDataReaderBufferedCells = 2
            });

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                reader.GetSheet("Data").ReadRangeAsDataReader("A1:C2"));

            Assert.Contains(nameof(ExcelReadOptions.MaxDataReaderBufferedCells), exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutHeadersPreservesBlankRowsInsideRange() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Alpha");
                sheet.CellValue(1, 3, 10);
                sheet.CellValue(3, 2, "Beta");
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:C3", headersInFirstRow: false, chunkRows: 1, schemaSampleRows: 1);

            Assert.Equal("Column1", dataReader.GetName(0));
            Assert.Equal("Column2", dataReader.GetName(1));
            Assert.Equal("Column3", dataReader.GetName(2));

            Assert.True(dataReader.Read());
            Assert.Equal("Alpha", dataReader.GetString(0));
            Assert.True(dataReader.IsDBNull(1));
            Assert.Equal(10d, dataReader.GetDouble(2));

            Assert.True(dataReader.Read());
            Assert.True(dataReader.IsDBNull(0));
            Assert.True(dataReader.IsDBNull(1));
            Assert.True(dataReader.IsDBNull(2));

            Assert.True(dataReader.Read());
            Assert.True(dataReader.IsDBNull(0));
            Assert.Equal("Beta", dataReader.GetString(1));
            Assert.True(dataReader.IsDBNull(2));
            Assert.False(dataReader.Read());
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_ReturnsSampledAndRemainingRows() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Alpha");
                sheet.CellValue(3, 1, 2);
                sheet.CellValue(3, 2, "Beta");
                sheet.CellValue(4, 1, 3);
                sheet.CellValue(4, 2, "Gamma");
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:B4", chunkRows: 1, schemaSampleRows: 2);

            Assert.True(dataReader.Read());
            Assert.Equal(1d, dataReader.GetDouble(0));
            Assert.Equal("Alpha", dataReader.GetString(1));

            Assert.True(dataReader.Read());
            Assert.Equal(2d, dataReader.GetDouble(0));
            Assert.Equal("Beta", dataReader.GetString(1));

            Assert.True(dataReader.Read());
            Assert.Equal(3d, dataReader.GetDouble(0));
            Assert.Equal("Gamma", dataReader.GetString(1));
            Assert.False(dataReader.Read());
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutSchemaSamples_StreamsRows() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Alpha");
                sheet.CellValue(4098, 1, 4097);
                sheet.CellValue(4098, 2, "Gamma");
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:B4098", schemaSampleRows: 0);
            var values = new object[dataReader.FieldCount];

            Assert.Equal(typeof(object), dataReader.GetFieldType(0));
            Assert.Throws<InvalidOperationException>(() => dataReader.GetValue(0));
            Assert.True(dataReader.Read());
            Assert.Equal(2, dataReader.GetValues(values));
            Assert.Equal(1, dataReader.GetInt32(0));
            Assert.Equal(1d, values[0]);
            Assert.Equal("Alpha", values[1]);

            Assert.True(dataReader.Read());
            Assert.Equal(DBNull.Value, dataReader.GetValue(0));
            Assert.Equal(DBNull.Value, dataReader.GetValue(1));

            int rowsRead = 2;
            object? lastId = null;
            object? lastName = null;
            int lastTypedId = 0;
            while (dataReader.Read()) {
                rowsRead++;
                dataReader.GetValues(values);
                lastId = values[0];
                lastName = values[1];
                if (lastId != DBNull.Value) {
                    lastTypedId = dataReader.GetInt32(0);
                }
            }

            Assert.Equal(4097, rowsRead);
            Assert.Equal(4097d, lastId);
            Assert.Equal("Gamma", lastName);
            Assert.Equal(4097, lastTypedId);
            Assert.False(dataReader.Read());
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutSchemaSamples_PreservesValuesAfterTypedAccess() {
            var expectedDate = new DateTime(2026, 7, 9);
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Created");
                sheet.CellValue(1, 3, "Active");
                sheet.CellValue(2, 1, 7);
                sheet.CellValue(2, 2, expectedDate);
                sheet.CellValue(2, 3, true);
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:C2", schemaSampleRows: 0);
            var values = new object[dataReader.FieldCount];

            Assert.True(dataReader.Read());
            Assert.Equal(7, dataReader.GetInt32(0));
            Assert.Equal(expectedDate, dataReader.GetDateTime(1));
            Assert.True(dataReader.GetBoolean(2));

            Assert.Equal(3, dataReader.GetValues(values));
            Assert.Equal(7d, values[0]);
            Assert.Equal(expectedDate, values[1]);
            Assert.Equal(true, values[2]);
            Assert.False(dataReader.Read());
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutSchemaSamples_PreservesUtf8FastPathCellKinds() {
            var expectedDate = new DateTime(2026, 7, 10, 12, 30, 0);
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                string[] headers = { "Id", "Direct", "Shared", "Created", "Formula", "Active", "Error", "Lines" };
                for (int column = 0; column < headers.Length; column++) {
                    sheet.CellValue(1, column + 1, headers[column]);
                }

                sheet.CellValue(2, 1, 7);
                sheet.CellValue(2, 2, "A & B < C");
                sheet.CellValue(2, 3, "Padded shared");
                sheet.CellValue(2, 4, expectedDate);
                sheet.CellValue(2, 5, 3);
                sheet.CellValue(2, 6, true);
                sheet.CellValue(2, 7, "#DIV/0!");
                sheet.CellValue(2, 8, "Line 1\r\nLine 2");
                sheet.CellValue(4098, 1, 4097);
            }

            memory.Position = 0;
            using (var spreadsheet = SpreadsheetDocument.Open(memory, true)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name == "Data");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);

                var direct = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "B2");
                direct.DataType = CellValues.String;
                direct.CellValue = new CellValue("A & B < C");

                var shared = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "C2");
                Assert.Equal(CellValues.SharedString, shared.DataType!.Value);
                shared.CellValue = new CellValue(" +" + shared.CellValue!.Text + " ");

                var formula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "E2");
                formula.DataType = CellValues.Number;
                formula.CellFormula = new CellFormula("1+2");
                formula.CellValue = new CellValue("3");

                var error = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "G2");
                error.DataType = CellValues.Error;
                error.CellValue = new CellValue("#DIV/0!");

                var lines = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "H2");
                lines.DataType = CellValues.String;
                lines.CellValue = new CellValue("Line 1\r\nLine 2");
                worksheetPart.Worksheet.Save();
            }

            byte[] workbook = memory.ToArray();
            using (var reader = ExcelDocumentReader.Open(workbook))
            using (var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:H4098", schemaSampleRows: 0)) {
                Assert.True(dataReader.Read());
                Assert.Equal(7, dataReader.GetInt32(0));
                Assert.Equal("A & B < C", dataReader.GetString(1));
                Assert.Equal("Padded shared", dataReader.GetString(2));
                Assert.Equal(expectedDate, dataReader.GetDateTime(3));
                Assert.Equal(3, dataReader.GetInt32(4));
                Assert.True(dataReader.GetBoolean(5));
                Assert.Equal("#DIV/0!", dataReader.GetString(6));
                Assert.Equal("Line 1\nLine 2", dataReader.GetString(7));
            }

            using (var reader = ExcelDocumentReader.Open(workbook, new ExcelReadOptions { UseCachedFormulaResult = false }))
            using (var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:H4098", schemaSampleRows: 0)) {
                Assert.True(dataReader.Read());
                Assert.Equal("1+2", dataReader.GetString(4));
            }
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutSchemaSamples_DecodesEntitiesBeforeTypedParsing() {
            var expectedDate = new DateTime(2026, 7, 10);
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Shared");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(1, 3, "Active");
                sheet.CellValue(1, 4, "Created");
                sheet.CellValue(2, 1, "Expected shared value");
                sheet.CellValue(2, 2, 1.5d);
                sheet.CellValue(2, 3, true);
                sheet.CellValue(2, 4, expectedDate);
                sheet.CellValue(4098, 1, "Last");
            }

            byte[] workbook = RewriteWorksheetCellValue(
                memory.ToArray(),
                "A2",
                value => string.Concat(value.Select(character => $"&#{(int)character};")));
            workbook = RewriteWorksheetCellValue(workbook, "B2", value => value.Replace(".", "&#46;"));
            workbook = RewriteWorksheetCellValue(workbook, "C2", _ => "&#49;");
            workbook = RewriteWorksheetCellValue(
                workbook,
                "D2",
                value => string.Concat(value.Select(character => $"&#{(int)character};")));

            using var reader = ExcelDocumentReader.Open(workbook);
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:D4098", schemaSampleRows: 0);

            Assert.True(dataReader.Read());
            Assert.Equal("Expected shared value", dataReader.GetString(0));
            Assert.Equal(1.5d, dataReader.GetDouble(1));
            Assert.True(dataReader.GetBoolean(2));
            Assert.Equal(expectedDate, dataReader.GetDateTime(3));
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutSchemaSamples_SupportsSimpleInlineStrings() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Placeholder");
                sheet.CellValue(4098, 1, 4097);
            }

            memory.Position = 0;
            using (var spreadsheet = SpreadsheetDocument.Open(memory, true)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name == "Data");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                var inline = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "B2");
                inline.DataType = CellValues.InlineString;
                inline.CellValue = null;
                inline.InlineString = new InlineString(new Text(" Inline & <value> ") { Space = SpaceProcessingModeValues.Preserve });
                worksheetPart.Worksheet.Save();
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:B4098", schemaSampleRows: 0);

            Assert.True(dataReader.Read());
            Assert.Equal(1, dataReader.GetInt32(0));
            Assert.Equal(" Inline & <value> ", dataReader.GetString(1));
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutSchemaSamples_FallsBackForRichInlineStrings() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Placeholder");
                sheet.CellValue(4098, 1, 4097);
            }

            memory.Position = 0;
            using (var spreadsheet = SpreadsheetDocument.Open(memory, true)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name == "Data");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                var inline = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "B2");
                inline.DataType = CellValues.InlineString;
                inline.CellValue = null;
                inline.InlineString = new InlineString(
                    new Run(new Text("Rich ")),
                    new Run(new RunProperties(new Bold()), new Text("inline")));
                worksheetPart.Worksheet.Save();
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:B4098", schemaSampleRows: 0);

            Assert.True(dataReader.Read());
            Assert.Equal(1, dataReader.GetInt32(0));
            Assert.Equal("Rich inline", dataReader.GetString(1));
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutSchemaSamples_ResolvesPaddedSharedStringIndexes() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Padded");
            }

            memory.Position = 0;
            using (var spreadsheet = SpreadsheetDocument.Open(memory, true)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(s => s.Name == "Data");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                var cell = worksheetPart.Worksheet.Descendants<Cell>().Single(c => c.CellReference == "A2");
                Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
                cell.CellValue = new CellValue(" " + cell.CellValue!.Text + " ");
                worksheetPart.Worksheet.Save();
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:A2", schemaSampleRows: 0);

            Assert.True(dataReader.Read());
            Assert.Equal("Padded", dataReader.GetString(0));
            Assert.False(dataReader.Read());
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutSchemaSamples_AllowsSelectiveColumnAccess() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(1, 3, "Note");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Alpha");
                sheet.CellValue(2, 3, "First note");
                sheet.CellValue(4098, 1, 4097);
                sheet.CellValue(4098, 2, "Gamma");
                sheet.CellValue(4098, 3, "Last note");
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:C4098", schemaSampleRows: 0);

            Assert.True(dataReader.Read());
            Assert.Equal(1, dataReader.GetInt32(0));
            Assert.Equal("First note", dataReader.GetString(2));

            Assert.True(dataReader.Read());
            Assert.True(dataReader.IsDBNull(0));

            int rowsRead = 2;
            int lastId = 0;
            while (dataReader.Read()) {
                rowsRead++;
                if (!dataReader.IsDBNull(0)) {
                    lastId = dataReader.GetInt32(0);
                }
            }

            Assert.Equal(4097, rowsRead);
            Assert.Equal(4097, lastId);
            Assert.False(dataReader.Read());
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutSchemaSamples_PreservesOutOfOrderCellsAcrossAccessOrders() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(1, 3, "Note");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Alpha");
                sheet.CellValue(2, 3, "First note");
                sheet.CellValue(3, 1, 2);
                sheet.CellValue(3, 2, "Beta");
                sheet.CellValue(3, 3, "Second note");
                sheet.CellValue(4098, 1, 4097);
            }

            memory.Position = 0;
            using (var spreadsheet = SpreadsheetDocument.Open(memory, true)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name == "Data");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                foreach (uint rowIndex in new[] { 2U, 3U }) {
                    var row = worksheetPart.Worksheet.Descendants<Row>().Single(item => item.RowIndex == rowIndex);
                    var cell = row.Elements<Cell>().Single(item => item.CellReference == "B" + rowIndex);
                    cell.Remove();
                    row.Append(cell);
                }

                worksheetPart.Worksheet.Save();
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:C4098", schemaSampleRows: 0);

            Assert.True(dataReader.Read());
            Assert.Equal("First note", dataReader.GetString(2));
            Assert.Equal("Alpha", dataReader.GetString(1));

            Assert.True(dataReader.Read());
            Assert.Equal("Beta", dataReader.GetString(1));
            Assert.Equal("Second note", dataReader.GetString(2));
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutSchemaSamples_PreservesLargeOutOfOrderRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataReaderLargeOutOfOrderRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorksheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(2, 1, "First");
                    sheet.CellValue(2049, 1, "Middle");
                    sheet.CellValue(4097, 1, "Last");
                    document.Save();
                }

                MoveWorksheetRowToEnd(filePath, 2049U);

                using var reader = ExcelDocumentReader.Open(filePath);
                using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:A4097", schemaSampleRows: 0, chunkRows: 512);

                Assert.True(dataReader.Read());
                Assert.Equal("First", dataReader.GetString(0));

                string? middle = null;
                string? last = null;
                int rowsRead = 1;
                while (dataReader.Read()) {
                    rowsRead++;
                    if (!dataReader.IsDBNull(0)) {
                        string value = dataReader.GetString(0);
                        if (rowsRead == 2048) {
                            middle = value;
                        }

                        if (rowsRead == 4096) {
                            last = value;
                        }
                    }
                }

                Assert.Equal(4096, rowsRead);
                Assert.Equal("Middle", middle);
                Assert.Equal("Last", last);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutSchemaSamples_PreservesInRangeRowsAfterOutOfRangeRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDataReaderOutOfRangeBeforeInRange.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorksheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(2, 1, "First");
                    sheet.CellValue(4097, 1, "Last in range");
                    sheet.CellValue(5000, 1, "Outside range");
                    document.Save();
                }

                MoveWorksheetRowToEnd(filePath, 4097U);

                using var reader = ExcelDocumentReader.Open(filePath);
                using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:A4097", schemaSampleRows: 0, chunkRows: 512);

                Assert.True(dataReader.Read());
                Assert.Equal("First", dataReader.GetString(0));

                string? last = null;
                int rowsRead = 1;
                while (dataReader.Read()) {
                    rowsRead++;
                    if (!dataReader.IsDBNull(0)) {
                        last = dataReader.GetString(0);
                    }
                }

                Assert.Equal(4096, rowsRead);
                Assert.Equal("Last in range", last);
                Assert.False(dataReader.Read());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static byte[] RewriteWorksheetCellValue(byte[] workbook, string cellReference, Func<string, string> rewrite) {
            using var memory = new MemoryStream();
            memory.Write(workbook, 0, workbook.Length);
            memory.Position = 0;

            using (var spreadsheet = SpreadsheetDocument.Open(memory, true)) {
                var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
                string xml;
                using (var stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read))
                using (var reader = new StreamReader(stream, Encoding.UTF8, true, 1024, leaveOpen: false)) {
                    xml = reader.ReadToEnd();
                }

                int cellStart = xml.IndexOf($"r=\"{cellReference}\"", StringComparison.Ordinal);
                int cellEnd = cellStart < 0 ? -1 : xml.IndexOf("</c>", cellStart, StringComparison.Ordinal);
                int valueStart = cellStart < 0 ? -1 : xml.IndexOf("<v>", cellStart, StringComparison.Ordinal);
                int valueEnd = valueStart < 0 ? -1 : xml.IndexOf("</v>", valueStart, StringComparison.Ordinal);
                Assert.True(cellStart >= 0 && cellEnd >= 0 && valueStart >= 0 && valueEnd > valueStart && valueEnd < cellEnd);

                valueStart += 3;
                string value = xml.Substring(valueStart, valueEnd - valueStart);
                string updated = xml.Substring(0, valueStart) + rewrite(value) + xml.Substring(valueEnd);
                using var output = worksheetPart.GetStream(FileMode.Create, FileAccess.Write);
                using var writer = new StreamWriter(output, new UTF8Encoding(false), 1024, leaveOpen: false);
                writer.Write(updated);
            }

            return memory.ToArray();
        }
    }
}
