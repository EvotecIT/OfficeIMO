using System.Data;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void WriteDataReader_WritesPackageAndLeavesReaderOpen() {
            var table = new DataTable("ReaderData");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add(1, "Alpha");
            table.Rows.Add(2, "Beta");
            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();

            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(output, reader);

            Assert.False(reader.IsClosed);
            Assert.Equal("A1:B3", result.Range);
            Assert.Equal(2, result.RowCount);
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet
                .Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Alpha", GetSpreadsheetCellText(spreadsheet, cells["B2"]));
            Assert.Equal("2", cells["A3"].CellValue!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void DirectTabularWriters_WritePackagesToNonSeekableStreams() {
            using (var rowOutput = new NonSeekableReadWriteBuffer(Array.Empty<byte>())) {
                ExcelDataSetImportResult result = ExcelDocument.WriteRows(
                    rowOutput,
                    new[] { new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true) },
                    ["Id", "Name"],
                    static (writer, row) => writer.Write(row.Id).Write(row.Name));

                Assert.Equal(1, result.RowCount);
                using var spreadsheet = SpreadsheetDocument.Open(
                    new MemoryStream(rowOutput.ToArray(), writable: false),
                    false);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            }

            var table = new DataTable("ReaderData");
            table.Columns.Add("Id", typeof(int));
            table.Rows.Add(1);
            using var reader = table.CreateDataReader();
            using var readerOutput = new NonSeekableReadWriteBuffer(Array.Empty<byte>());

            ExcelDataSetImportResult readerResult = ExcelDocument.WriteDataReader(readerOutput, reader);

            Assert.Equal(1, readerResult.RowCount);
            using var readerSpreadsheet = SpreadsheetDocument.Open(
                new MemoryStream(readerOutput.ToArray(), writable: false),
                false);
            Assert.Empty(new OpenXmlValidator().Validate(readerSpreadsheet));
        }

        [Fact]
        public void WriteDataReader_CompactPackageStreamsAndRoundTrips() {
            var table = new DataTable("ReaderData");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Region", typeof(string));
            table.Columns.Add("Owner", typeof(string));
            table.Columns.Add("CreatedOn", typeof(DateTime));
            table.Columns.Add("Amount", typeof(double));
            table.Columns.Add("Units", typeof(int));
            table.Columns.Add("Active", typeof(bool));
            table.Columns.Add("Notes", typeof(string));
            table.Rows.Add(1, "North", "Ava", new DateTime(2026, 7, 10), 123.45, 2, true, "Alpha");
            table.Rows.Add(2, "South", "Noah", new DateTime(2026, 7, 11), 678.90, 4, false, "Beta");
            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();

            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(
                output,
                reader,
                new ExcelTabularWriteOptions {
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });

            Assert.False(reader.IsClosed);
            Assert.Equal("A1:H3", result.Range);
            Assert.Equal(2, result.RowCount);
            using (var spreadsheet = SpreadsheetDocument.Open(output, false)) {
                var savedRows = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet
                    .Descendants<Row>()
                    .ToArray();
                Assert.All(savedRows[0].Elements<Cell>(), cell => Assert.NotNull(cell.CellReference));
                Assert.All(savedRows.Skip(1).SelectMany(row => row.Elements<Cell>()), cell => Assert.Null(cell.CellReference));
                Assert.Null(spreadsheet.WorkbookPart.SharedStringTablePart);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            }

            using var workbookReader = ExcelDocumentReader.Open(output);
            object?[,] values = workbookReader.GetSheet("Data").ReadRange("A1:H3");
            Assert.Equal("Id", values[0, 0]);
            Assert.Equal("Ava", values[1, 2]);
            Assert.Equal(678.90, Convert.ToDouble(values[2, 4], CultureInfo.InvariantCulture));
            Assert.Equal(false, values[2, 6]);
        }

        [Fact]
        public void WriteDataReader_CompactPackagePreservesDoubleValuesExactly() {
            double[] expected = [
                0D,
                -0.5D,
                123.4D,
                123.45D,
                -9876.5D,
                NextRepresentableDouble(123.45D),
                Math.PI,
                90_071_992_547_409.9D
            ];
            var table = new DataTable("Doubles");
            table.Columns.Add("Value", typeof(double));
            foreach (double value in expected) {
                table.Rows.Add(value);
            }

            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();
            ExcelDocument.WriteDataReader(
                output,
                reader,
                new ExcelTabularWriteOptions {
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });

            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            Cell[] cells = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet
                .Descendants<Cell>()
                .Skip(1)
                .ToArray();
            Assert.Equal(expected.Length, cells.Length);
            for (int index = 0; index < expected.Length; index++) {
                string rawValue = cells[index].CellValue!.Text;
                double actual = double.Parse(rawValue, CultureInfo.InvariantCulture);
                Assert.Equal(BitConverter.DoubleToInt64Bits(expected[index]), BitConverter.DoubleToInt64Bits(actual));
            }

            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Theory]
        [InlineData(ExcelDateSystem.NineteenHundred)]
        [InlineData(ExcelDateSystem.NineteenFour)]
        public void WriteRows_DateSerialFastPathPreservesOaDateRoundTrip(ExcelDateSystem dateSystem) {
            DateTime[] expected = [
                new DateTime(1900, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified),
                new DateTime(1900, 3, 1, 23, 59, 59, 999, DateTimeKind.Unspecified),
                new DateTime(1904, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified),
                new DateTime(2026, 8, 9, 8, 17, 42, 123, DateTimeKind.Unspecified).AddTicks(4567),
                new DateTime(9999, 12, 31, 23, 59, 59, 999, DateTimeKind.Unspecified)
            ];
            using var output = new MemoryStream();

            ExcelDocument.WriteRows(
                output,
                expected,
                ["When"],
                static (writer, value) => writer.Write(value),
                new ExcelTabularWriteOptions {
                    DateSystem = dateSystem,
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });

            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            Cell[] cells = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet
                .Descendants<Cell>()
                .Skip(1)
                .ToArray();
            Assert.Equal(expected.Length, cells.Length);
            for (int index = 0; index < expected.Length; index++) {
                double serial = double.Parse(cells[index].CellValue!.Text, CultureInfo.InvariantCulture);
                DateTime expectedRoundTrip = ExcelDateSystemConverter.FromSerial(
                    ExcelDateSystemConverter.ToSerial(expected[index], dateSystem),
                    dateSystem);
                Assert.Equal(expectedRoundTrip, ExcelDateSystemConverter.FromSerial(serial, dateSystem));
            }

            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void WriteDataReader_CompiledSchemaLanePreservesMixedColumnsAndNulls() {
            var table = new DataTable("MixedReaderData");
            table.Columns.Add("When", typeof(DateTime));
            table.Columns.Add("Active", typeof(bool));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(double));
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Amount", typeof(decimal));
            table.Columns.Add("Large", typeof(long));
            table.Rows.Add(new DateTime(2026, 8, 9, 8, 30, 0), true, "First", 1.25, 7, 12.50m, 9_000_000_000L);
            table.Rows.Add(new DateTime(2026, 8, 10, 9, 45, 0), false, DBNull.Value, 2.5, 8, 15.75m, 9_000_000_001L);
            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();

            ExcelDocument.WriteDataReader(
                output,
                reader,
                new ExcelTabularWriteOptions {
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });

            using var workbookReader = ExcelDocumentReader.Open(output);
            object?[,] values = workbookReader.GetSheet("Data").ReadRange("A1:G3");
            Assert.Equal(new DateTime(2026, 8, 9, 8, 30, 0), values[1, 0]);
            Assert.Equal(true, values[1, 1]);
            Assert.Equal("First", values[1, 2]);
            Assert.Equal(1.25, values[1, 3]);
            Assert.Equal(7d, values[1, 4]);
            Assert.Equal(12.5, values[1, 5]);
            Assert.Equal(9_000_000_000d, values[1, 6]);
            Assert.Equal(string.Empty, values[2, 2]);
        }

        [Fact]
        public void WriteDataReader_CompiledSchemaLaneSupportsProviderBackedReaders() {
            var sourceTable = new DataTable("ProviderSource");
            sourceTable.Columns.Add("Id", typeof(int));
            sourceTable.Columns.Add("Name", typeof(string));
            sourceTable.Columns.Add("When", typeof(DateTime));
            sourceTable.Columns.Add("Active", typeof(bool));
            sourceTable.Rows.Add(7, "First", new DateTime(2026, 8, 9, 8, 30, 0), true);
            sourceTable.Rows.Add(8, DBNull.Value, new DateTime(2026, 8, 10, 9, 45, 0), false);

            using var sourcePackage = new MemoryStream();
            using (var sourceReader = sourceTable.CreateDataReader()) {
                ExcelDocument.WriteDataReader(
                    sourcePackage,
                    sourceReader,
                    new ExcelTabularWriteOptions {
                        IncludeCellReferences = false,
                        UseSharedStrings = false
                    });
            }
            sourcePackage.Position = 0;

            using ExcelWorkbookDataReader providerReader = ExcelDocument.OpenDataReader(
                sourcePackage,
                new ExcelReadOptions {
                    SheetName = "Data",
                    InferSchema = true
                });
            Assert.All(
                Enumerable.Range(0, providerReader.FieldCount),
                ordinal => Assert.NotEqual(typeof(object), providerReader.GetFieldType(ordinal)));
            using var countingReader = new CountingDataReader(providerReader);
            using var output = new MemoryStream();
            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(
                output,
                countingReader,
                new ExcelTabularWriteOptions {
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });

            Assert.Equal(2, result.RowCount);
            Assert.Equal(0, countingReader.GetValuesCalls);
            Assert.Equal(0, countingReader.GetValueCalls);
            using var workbookReader = ExcelDocumentReader.Open(output);
            object?[,] values = workbookReader.GetSheet("Data").ReadRange("A1:D3");
            Assert.Equal(7d, values[1, 0]);
            Assert.Equal("First", values[1, 1]);
            Assert.Equal(new DateTime(2026, 8, 9, 8, 30, 0), values[1, 2]);
            Assert.Equal(true, values[1, 3]);
            Assert.Equal(string.Empty, values[2, 1]);
            Assert.Equal(false, values[2, 3]);
        }

        [Fact]
        public void WriteDataReader_CompiledSchemaLaneUsesNonNullableProviderMetadata() {
            var table = new DataTable("NonNullableReaderData");
            table.Columns.Add("Id", typeof(int)).AllowDBNull = false;
            table.Columns.Add("Name", typeof(string)).AllowDBNull = false;
            table.Rows.Add(7, "First");
            table.Rows.Add(8, "Second");
            using var countingReader = new CountingDataReader(table.CreateDataReader());
            using var output = new MemoryStream();

            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(
                output,
                countingReader,
                new ExcelTabularWriteOptions {
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });

            Assert.Equal(2, result.RowCount);
            Assert.Equal(0, countingReader.GetValuesCalls);
            Assert.Equal(0, countingReader.GetValueCalls);
            Assert.Equal(0, countingReader.IsDBNullCalls);
            using var workbookReader = ExcelDocumentReader.Open(output);
            object?[,] values = workbookReader.GetSheet("Data").ReadRange("A1:B3");
            Assert.Equal(7d, values[1, 0]);
            Assert.Equal("Second", values[2, 1]);
        }

        [Fact]
        public void WriteDataReader_HeaderlessEmptyReaderWritesValidEmptySheet() {
            var table = new DataTable("Empty");
            table.Columns.Add("Id", typeof(int));
            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();

            ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(
                output,
                reader,
                new ExcelTabularWriteOptions { IncludeHeaders = false });

            Assert.Equal(string.Empty, result.Range);
            Assert.Equal(0, result.RowCount);
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            Assert.Empty(spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void WriteRows_StreamsTypedValuesAndRoundTrips() {
            using var output = new MemoryStream();
            var rows = new[] {
                new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10, 8, 30, 0), true),
                new TabularWriteRow(2, "Beta", new DateTime(2026, 7, 11, 9, 45, 0), false)
            };

            ExcelDataSetImportResult result = ExcelDocument.WriteRows(
                output,
                rows,
                ["Id", "Name", "Created", "Active"],
                static (writer, row) => writer
                    .Write(row.Id)
                    .Write(row.Name)
                    .Write(row.Created)
                    .Write(row.Active),
                new ExcelTabularWriteOptions { IncludeCellReferences = false, UseSharedStrings = false });

            Assert.Equal("A1:D3", result.Range);
            using var reader = ExcelDocumentReader.Open(new MemoryStream(output.ToArray(), writable: false));
            object?[,] values = reader.GetSheet("Data").ReadRange("A1:D3");
            Assert.Equal("Beta", values[2, 1]);
            Assert.Equal(new DateTime(2026, 7, 11, 9, 45, 0), values[2, 2]);
            Assert.Equal(false, values[2, 3]);
        }

        [Fact]
        public void WriteRows_ConsumesGeneratorWhileWriting() {
            using var output = new MemoryStream();
            bool enumerationActive = false;

            ExcelDataSetImportResult result = ExcelDocument.WriteRows(
                output,
                StreamRows(),
                ["Id", "Name"],
                (writer, row) => {
                    Assert.True(enumerationActive);
                    writer.Write(row.Id).Write(row.Name);
                });

            Assert.False(enumerationActive);
            Assert.Equal(2, result.RowCount);
            Assert.Equal("A1:B3", result.Range);

            IEnumerable<TabularWriteRow> StreamRows() {
                enumerationActive = true;
                try {
                    yield return new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true);
                    yield return new TabularWriteRow(2, "Beta", new DateTime(2026, 7, 11), false);
                } finally {
                    enumerationActive = false;
                }
            }
        }

        [Fact]
        public void WriteRows_CancellationStopsBeforeAdvancingGenerator() {
            using var output = new MemoryStream();
            using var cts = new CancellationTokenSource();
            bool enumerationStarted = false;
            cts.Cancel();

            Assert.Throws<OperationCanceledException>(() => ExcelDocument.WriteRows(
                output,
                StreamRows(),
                ["Id"],
                static (writer, row) => writer.Write(row.Id),
                ct: cts.Token));

            Assert.False(enumerationStarted);

            IEnumerable<TabularWriteRow> StreamRows() {
                enumerationStarted = true;
                yield return new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true);
            }
        }

        [Fact]
        public void WriteRows_TableModeCancellationStopsWhileBufferingGenerator() {
            using var output = new MemoryStream();
            using var cts = new CancellationTokenSource();
            int advances = 0;

            Assert.Throws<OperationCanceledException>(() => ExcelDocument.WriteRows(
                output,
                StreamRows(),
                ["Id"],
                static (writer, row) => writer.Write(row),
                new ExcelTabularWriteOptions { CreateTable = true },
                cts.Token));

            Assert.Equal(1, advances);
            Assert.Equal(0, output.Length);

            IEnumerable<int> StreamRows() {
                advances++;
                cts.Cancel();
                yield return 1;
                advances++;
                yield return 2;
            }
        }

        [Fact]
        public void WriteRows_TableModeStopsUnknownSequenceAtWorksheetLimit() {
            using var output = new MemoryStream();
            int advances = 0;

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => ExcelDocument.WriteRows(
                output,
                StreamRows(),
                ["Value"],
                static (writer, value) => writer.Write(value),
                new ExcelTabularWriteOptions { CreateTable = true }));

            Assert.Contains("maximum worksheet row count", exception.Message, StringComparison.Ordinal);
            Assert.Equal(1_048_576, advances);
            Assert.Equal(0, output.Length);

            IEnumerable<byte> StreamRows() {
                for (int index = 0; index < 1_048_576; index++) {
                    advances++;
                    yield return 0;
                }
            }
        }

        [Fact]
        public async Task WriteRowsAsync_AwaitsAndWritesRowsSinglePass() {
            using var output = new MemoryStream();
            bool enumerationActive = false;

            ExcelDataSetImportResult result = await ExcelDocument.WriteRowsAsync(
                output,
                StreamRows(),
                ["Id", "Name", "Created", "Active"],
                (writer, row) => {
                    Assert.True(enumerationActive);
                    writer
                        .Write(row.Id)
                        .Write(row.Name)
                        .Write(row.Created)
                        .Write(row.Active);
                },
                new ExcelTabularWriteOptions {
                    SheetName = "Async Rows",
                    IncludeCellReferences = false
                });

            Assert.False(enumerationActive);
            Assert.Equal("Async Rows", result.SheetName);
            Assert.Equal("A1:D3", result.Range);
            Assert.Equal(2, result.RowCount);

            byte[] package = output.ToArray();
            using (var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false)) {
                Row[] savedRows = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet
                    .Descendants<Row>()
                    .ToArray();
                Assert.All(savedRows[0].Elements<Cell>(), cell => Assert.NotNull(cell.CellReference));
                Assert.All(savedRows.Skip(1).SelectMany(row => row.Elements<Cell>()), cell => Assert.Null(cell.CellReference));
                Assert.Null(spreadsheet.WorkbookPart.SharedStringTablePart);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            }

            using var reader = ExcelDocumentReader.Open(new MemoryStream(package, writable: false));
            object?[,] values = reader.GetSheet("Async Rows").ReadRange("A1:D3");
            Assert.Equal("Beta", values[2, 1]);
            Assert.Equal(new DateTime(2026, 7, 11), values[2, 2]);
            Assert.Equal(false, values[2, 3]);

            async IAsyncEnumerable<TabularWriteRow> StreamRows() {
                enumerationActive = true;
                try {
                    await Task.Yield();
                    yield return new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true);
                    await Task.Yield();
                    yield return new TabularWriteRow(2, "Beta", new DateTime(2026, 7, 11), false);
                } finally {
                    enumerationActive = false;
                }
            }
        }

        [Fact]
        public async Task WriteRowsAsync_CancellationStopsBeforeAdvancingAgainAndDisposesSource() {
            using var output = new MemoryStream();
            using var cts = new CancellationTokenSource();
            int advances = 0;
            bool disposed = false;

            await Assert.ThrowsAsync<OperationCanceledException>(() => ExcelDocument.WriteRowsAsync(
                output,
                StreamRows(),
                ["Id"],
                (writer, row) => {
                    writer.Write(row.Id);
                    cts.Cancel();
                },
                ct: cts.Token));

            Assert.Equal(1, advances);
            Assert.True(disposed);

            async IAsyncEnumerable<TabularWriteRow> StreamRows() {
                try {
                    advances++;
                    await Task.Yield();
                    yield return new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true);
                    advances++;
                    yield return new TabularWriteRow(2, "Beta", new DateTime(2026, 7, 11), false);
                } finally {
                    disposed = true;
                }
            }
        }

        [Fact]
        public async Task WriteRowsAsync_RejectsOptionsThatRequireBuffering() {
            using var output = new MemoryStream();

            await Assert.ThrowsAsync<ArgumentException>(() => ExcelDocument.WriteRowsAsync(
                output,
                EmptyRows(),
                ["Id"],
                static (writer, row) => writer.Write(row.Id),
                new ExcelTabularWriteOptions { CreateTable = true }));

            static async IAsyncEnumerable<TabularWriteRow> EmptyRows() {
                await Task.CompletedTask;
                yield break;
            }
        }

        [Fact]
        public async Task WriteRowsAsync_WritesHeaderlessEmptyPackageToNonSeekableStream() {
            using var output = new NonSeekableReadWriteBuffer(Array.Empty<byte>());

            ExcelDataSetImportResult result = await ExcelDocument.WriteRowsAsync(
                output,
                EmptyRows(),
                ["Id"],
                static (writer, row) => writer.Write(row.Id),
                new ExcelTabularWriteOptions { IncludeHeaders = false });

            Assert.Equal(string.Empty, result.Range);
            Assert.Equal(0, result.RowCount);
            using var spreadsheet = SpreadsheetDocument.Open(
                new MemoryStream(output.ToArray(), writable: false),
                false);
            Assert.Empty(spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));

            static async IAsyncEnumerable<TabularWriteRow> EmptyRows() {
                await Task.CompletedTask;
                yield break;
            }
        }

        [Fact]
        public void WriteRows_DefaultOptionsObjectUsesInlineStringsAndPreservesSettings() {
            using var output = new MemoryStream();
            var options = new ExcelTabularWriteOptions {
                SheetName = "Configured Rows",
                CreateTable = true,
                TableName = "ConfiguredRows"
            };

            ExcelDataSetImportResult result = ExcelDocument.WriteRows(
                output,
                new[] { new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true) },
                ["Id", "Name", "Created", "Active"],
                static (writer, row) => writer
                    .Write(row.Id)
                    .Write(row.Name)
                    .Write(row.Created)
                    .Write(row.Active),
                options);

            Assert.True(options.UseSharedStrings);
            Assert.Equal("Configured Rows", result.SheetName);
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            Assert.Null(spreadsheet.WorkbookPart!.SharedStringTablePart);
            var table = spreadsheet.WorkbookPart.WorksheetParts.Single().TableDefinitionParts.Single().Table;
            Assert.Equal("ConfiguredRows", table!.Name?.Value);
        }

        [Fact]
        public void WriteRows_RejectsRowsWithTheWrongCellCount() {
            using var output = new MemoryStream();

            Assert.Throws<InvalidOperationException>(() => ExcelDocument.WriteRows(
                output,
                new[] { new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true) },
                ["Id", "Name"],
                static (writer, row) => writer.Write(row.Id),
                new ExcelTabularWriteOptions { UseSharedStrings = false }));
        }

        [Fact]
        public void WriteRows_RejectsRowsWithTooManyCells() {
            using var output = new MemoryStream();

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => ExcelDocument.WriteRows(
                output,
                new[] { new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true) },
                ["Id"],
                static (writer, row) => writer.Write(row.Id).Write(row.Name),
                new ExcelTabularWriteOptions { UseSharedStrings = false }));

            Assert.Contains("more than 1 cells", exception.Message, StringComparison.Ordinal);
        }

        private static double NextRepresentableDouble(double value) {
            long bits = BitConverter.DoubleToInt64Bits(value);
            return BitConverter.Int64BitsToDouble(bits + 1);
        }

        private sealed record TabularWriteRow(int Id, string? Name, DateTime Created, bool Active);
    }
}
