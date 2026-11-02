using System.Data;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.SharedSource.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(true, false, false, true)]
        [InlineData(false, true, false, true)]
        [InlineData(false, false, true, false)]
        public void WriteDataReader_RequireStreamingRejectsBufferingBeforeReading(bool sharedStrings, bool autoFit, bool table, bool headers) {
            var source = new DataTable();
            source.Columns.Add("Value", typeof(int));
            source.Rows.Add(42);
            using var reader = source.CreateDataReader();
            using var output = new MemoryStream();
            Assert.Throws<ArgumentException>(() => {
                var destination1 = output;
                var writeOptions1 = new ExcelWriterOptions {
                    RequireStreaming = true,
                    UseSharedStrings = sharedStrings,
                    AutoFit = autoFit,
                    CreateTable = table,
                    IncludeHeaders = headers
                };
                using (var workbookWriter1 = ExcelDocument.CreateWriter(destination1, writeOptions1, leaveOpen: true)) {
                    return workbookWriter1.WriteTable(writeOptions1.SheetName, reader);
                }
            });
            Assert.Equal(0, output.Length);
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
        }

        [Fact]
        public void PooledUtf8TextWriter_PreservesSurrogatePairsAcrossFlushes() {
            using var output = new MemoryStream();
            using (var writer = new PooledUtf8TextWriter(output, new UTF8Encoding(false), bufferSize: 16, leaveOpen: true)) {
                writer.Write("prefix ");
                writer.Write('\ud83d');
                writer.Flush();
                writer.Write('\ude80');
                writer.Write(" & <suffix>");
            }

            Assert.True(output.CanWrite);
            Assert.Equal("prefix \ud83d\ude80 & <suffix>", Encoding.UTF8.GetString(output.ToArray()));
        }

        [Fact]
        public void WriteDataReader_WritesPackageAndLeavesReaderOpen() {
            var table = new DataTable("ReaderData");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add(1, "Alpha");
            table.Rows.Add(2, "Beta");
            using var reader = table.CreateDataReader();
            using var output = new MemoryStream();

            ExcelDataSetImportResult result;

            var destination3 = output;

            var writeOptions3 = new ExcelWriterOptions {
                UseSharedStrings = false,
                CreateTable = false
            };

            using (var workbookWriter3 = ExcelDocument.CreateWriter(destination3, writeOptions3, leaveOpen: true)) {
                result = workbookWriter3.WriteTable(writeOptions3.SheetName, reader);
            }

            if (destination3.CanSeek)
                destination3.Position = 0;

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
                ExcelDataSetImportResult result;

                var destination4 = rowOutput;

                var writeOptions4 = new ExcelWriterOptions {
                    UseSharedStrings = false,
                    CreateTable = false
                };

                using (var workbookWriter4 = ExcelDocument.CreateWriter(destination4, writeOptions4, leaveOpen: true)) {
                    result = workbookWriter4.WriteTable(writeOptions4.SheetName, new[] { new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true) }, ["Id", "Name"], static (writer, row) => writer.Write(row.Id).Write(row.Name));
                }

                if (destination4.CanSeek)
                    destination4.Position = 0;

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

            ExcelDataSetImportResult readerResult;

            var destination5 = readerOutput;

            var writeOptions5 = new ExcelWriterOptions {
                UseSharedStrings = false,
                CreateTable = false
            };

            using (var workbookWriter5 = ExcelDocument.CreateWriter(destination5, writeOptions5, leaveOpen: true)) {
                readerResult = workbookWriter5.WriteTable(writeOptions5.SheetName, reader);
            }

            if (destination5.CanSeek)
                destination5.Position = 0;

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

            ExcelDataSetImportResult result;

            var destination6 = output;

            var writeOptions6 = new ExcelWriterOptions {
                IncludeCellReferences = false,
                UseSharedStrings = false
            };

            using (var workbookWriter6 = ExcelDocument.CreateWriter(destination6, writeOptions6, leaveOpen: true)) {
                result = workbookWriter6.WriteTable(writeOptions6.SheetName, reader);
            }

            if (destination6.CanSeek)
                destination6.Position = 0;

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
            var destination7 = output;

            var writeOptions7 = new ExcelWriterOptions {
                IncludeCellReferences = false,
                UseSharedStrings = false
            };

            using (var workbookWriter7 = ExcelDocument.CreateWriter(destination7, writeOptions7, leaveOpen: true)) {
                workbookWriter7.WriteTable(writeOptions7.SheetName, reader);
            }

            if (destination7.CanSeek)
                destination7.Position = 0;

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

            var destination8 = output;

            var writeOptions8 = new ExcelWriterOptions {
                DateSystem = dateSystem,
                IncludeCellReferences = false,
                UseSharedStrings = false
            };

            using (var workbookWriter8 = ExcelDocument.CreateWriter(destination8, writeOptions8, leaveOpen: true)) {
                workbookWriter8.WriteTable(writeOptions8.SheetName, expected, ["When"], static (writer, value) => writer.Write(value));
            }

            if (destination8.CanSeek)
                destination8.Position = 0;

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

            var destination9 = output;

            var writeOptions9 = new ExcelWriterOptions {
                IncludeCellReferences = false,
                UseSharedStrings = false
            };

            using (var workbookWriter9 = ExcelDocument.CreateWriter(destination9, writeOptions9, leaveOpen: true)) {
                workbookWriter9.WriteTable(writeOptions9.SheetName, reader);
            }

            if (destination9.CanSeek)
                destination9.Position = 0;

            using var workbookReader = ExcelDocumentReader.Open(output);
            object?[,] values = workbookReader.GetSheet("Data").ReadRange("A1:G3");
            Assert.Equal(new DateTime(2026, 8, 9, 8, 30, 0), values[1, 0]);
            Assert.Equal(true, values[1, 1]);
            Assert.Equal("First", values[1, 2]);
            Assert.Equal(1.25, values[1, 3]);
            Assert.Equal(7d, values[1, 4]);
            Assert.Equal(12.5, values[1, 5]);
            Assert.Equal(9_000_000_000d, values[1, 6]);
            Assert.Null(values[2, 2]);
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
                var destination10 = sourcePackage;

                var writeOptions10 = new ExcelWriterOptions {
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                };

                using (var workbookWriter10 = ExcelDocument.CreateWriter(destination10, writeOptions10, leaveOpen: true)) {
                    workbookWriter10.WriteTable(writeOptions10.SheetName, sourceReader);
                }

                if (destination10.CanSeek)
                    destination10.Position = 0;
            }
            sourcePackage.Position = 0;

            using ExcelWorkbookDataReader providerReader = ExcelDocument.OpenReader(
                sourcePackage,
                new ExcelReaderOptions {
                    SheetName = "Data",
                    InferSchema = true
                });
            Assert.All(
                Enumerable.Range(0, providerReader.FieldCount),
                ordinal => Assert.NotEqual(typeof(object), providerReader.GetFieldType(ordinal)));
            using var countingReader = new CountingDataReader(providerReader);
            using var output = new MemoryStream();
            ExcelDataSetImportResult result;

            var destination11 = output;

            var writeOptions11 = new ExcelWriterOptions {
                IncludeCellReferences = false,
                UseSharedStrings = false
            };

            using (var workbookWriter11 = ExcelDocument.CreateWriter(destination11, writeOptions11, leaveOpen: true)) {
                result = workbookWriter11.WriteTable(writeOptions11.SheetName, countingReader);
            }

            if (destination11.CanSeek)
                destination11.Position = 0;

            Assert.Equal(2, result.RowCount);
            using var workbookReader = ExcelDocumentReader.Open(output);
            object?[,] values = workbookReader.GetSheet("Data").ReadRange("A1:D3");
            Assert.Equal(7d, values[1, 0]);
            Assert.Equal("First", values[1, 1]);
            Assert.Equal(new DateTime(2026, 8, 9, 8, 30, 0), values[1, 2]);
            Assert.Equal(true, values[1, 3]);
            Assert.Null(values[2, 1]);
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

            ExcelDataSetImportResult result;

            var destination12 = output;

            var writeOptions12 = new ExcelWriterOptions {
                IncludeCellReferences = false,
                UseSharedStrings = false
            };

            using (var workbookWriter12 = ExcelDocument.CreateWriter(destination12, writeOptions12, leaveOpen: true)) {
                result = workbookWriter12.WriteTable(writeOptions12.SheetName, countingReader);
            }

            if (destination12.CanSeek)
                destination12.Position = 0;

            Assert.Equal(2, result.RowCount);
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

            ExcelDataSetImportResult result;

            var destination13 = output;

            var writeOptions13 = new ExcelWriterOptions {
                IncludeHeaders = false
            };

            using (var workbookWriter13 = ExcelDocument.CreateWriter(destination13, writeOptions13, leaveOpen: true)) {
                result = workbookWriter13.WriteTable(writeOptions13.SheetName, reader);
            }

            if (destination13.CanSeek)
                destination13.Position = 0;

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

            ExcelDataSetImportResult result;

            var destination14 = output;

            var writeOptions14 = new ExcelWriterOptions {
                IncludeCellReferences = false,
                UseSharedStrings = false
            };

            using (var workbookWriter14 = ExcelDocument.CreateWriter(destination14, writeOptions14, leaveOpen: true)) {
                result = workbookWriter14.WriteTable(writeOptions14.SheetName, rows, ["Id", "Name", "Created", "Active"], static (writer, row) => writer.Write(row.Id).Write(row.Name).Write(row.Created).Write(row.Active));
            }

            if (destination14.CanSeek)
                destination14.Position = 0;

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

            ExcelDataSetImportResult result;

            var destination15 = output;

            var writeOptions15 = new ExcelWriterOptions {
                UseSharedStrings = false,
                CreateTable = false
            };

            using (var workbookWriter15 = ExcelDocument.CreateWriter(destination15, writeOptions15, leaveOpen: true)) {
                result = workbookWriter15.WriteTable(writeOptions15.SheetName, StreamRows(), ["Id", "Name"], (writer, row) => {
                    Assert.True(enumerationActive);
                    writer.Write(row.Id).Write(row.Name);
                });
            }

            if (destination15.CanSeek)
                destination15.Position = 0;

            Assert.False(enumerationActive);
            Assert.Equal(2, result.RowCount);
            Assert.Equal("A1:B3", result.Range);

            IEnumerable<TabularWriteRow> StreamRows() {
                enumerationActive = true;
                try {
                    yield return new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true);
                    yield return new TabularWriteRow(2, "Beta", new DateTime(2026, 7, 11), false);
                }
                finally {
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

            Assert.Throws<OperationCanceledException>(() => {
                var destination16 = output;
                var writeOptions16 = new ExcelWriterOptions {
                    UseSharedStrings = false,
                    CreateTable = false
                };
                using (var workbookWriter16 = ExcelDocument.CreateWriter(destination16, writeOptions16, leaveOpen: true)) {
                    return workbookWriter16.WriteTable(writeOptions16.SheetName, StreamRows(), ["Id"], static (writer, row) => writer.Write(row.Id), cts.Token);
                }
            });

            Assert.False(enumerationStarted);

            IEnumerable<TabularWriteRow> StreamRows() {
                enumerationStarted = true;
                yield return new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true);
            }
        }

        [Fact]
        public void WriteTable_CancellationStopsSource() {
            using var output = new MemoryStream();
            using var cts = new CancellationTokenSource();
            int advances = 0;

            Assert.Throws<OperationCanceledException>(() => {
                var destination17 = output;
                var writeOptions17 = new ExcelWriterOptions {
                    CreateTable = true
                };
                using (var workbookWriter17 = ExcelDocument.CreateWriter(destination17, writeOptions17, leaveOpen: true)) {
                    return workbookWriter17.WriteTable(writeOptions17.SheetName, StreamRows(), ["Id"], static (writer, row) => writer.Write(row), cts.Token);
                }
            });

            Assert.Equal(1, advances);
            Assert.True(output.CanWrite);

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

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => {
                var destination18 = output;
                var writeOptions18 = new ExcelWriterOptions {
                    CreateTable = true
                };
                using (var workbookWriter18 = ExcelDocument.CreateWriter(destination18, writeOptions18, leaveOpen: true)) {
                    return workbookWriter18.WriteTable(writeOptions18.SheetName, StreamRows(), ["Value"], static (writer, value) => writer.Write(value));
                }
            });

            Assert.Contains("maximum worksheet row count", exception.Message, StringComparison.Ordinal);
            Assert.Equal(1_048_576, advances);
            Assert.True(output.CanWrite);

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

            ExcelDataSetImportResult result;

            var destination19 = output;

            var writeOptions19 = new ExcelWriterOptions {
                SheetName = "Async Rows",
                IncludeCellReferences = false,
                UseSharedStrings = false
            };

            using (var workbookWriter19 = ExcelDocument.CreateWriter(destination19, writeOptions19, leaveOpen: true)) {
                result = await workbookWriter19.WriteTable(writeOptions19.SheetName, StreamRows(), ["Id", "Name", "Created", "Active"], (writer, row) => {
                    Assert.True(enumerationActive);
                    writer.Write(row.Id).Write(row.Name).Write(row.Created).Write(row.Active);
                });
            }

            if (destination19.CanSeek)
                destination19.Position = 0;

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
                }
                finally {
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

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => {
                var destination20 = output;
                var writeOptions20 = new ExcelWriterOptions {
                    UseSharedStrings = false,
                    CreateTable = false
                };
                using (var workbookWriter20 = ExcelDocument.CreateWriter(destination20, writeOptions20, leaveOpen: true)) {
                    await workbookWriter20.WriteTable(writeOptions20.SheetName, StreamRows(), ["Id"], (writer, row) => {
                        writer.Write(row.Id);
                        cts.Cancel();
                    }, cts.Token);
                }
            });

            Assert.Equal(1, advances);
            Assert.True(disposed);

            async IAsyncEnumerable<TabularWriteRow> StreamRows() {
                try {
                    advances++;
                    await Task.Yield();
                    yield return new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true);
                    advances++;
                    yield return new TabularWriteRow(2, "Beta", new DateTime(2026, 7, 11), false);
                }
                finally {
                    disposed = true;
                }
            }
        }

        [Fact]
        public async Task WriteRowsAsync_WritesHeaderlessEmptyPackageToNonSeekableStream() {
            using var output = new NonSeekableReadWriteBuffer(Array.Empty<byte>());

            ExcelDataSetImportResult result;

            var destination22 = output;

            var writeOptions22 = new ExcelWriterOptions {
                IncludeHeaders = false
            };

            using (var workbookWriter22 = ExcelDocument.CreateWriter(destination22, writeOptions22, leaveOpen: true)) {
                result = await workbookWriter22.WriteTable(writeOptions22.SheetName, EmptyRows(), ["Id"], static (writer, row) => writer.Write(row.Id));
            }

            if (destination22.CanSeek)
                destination22.Position = 0;

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
        public void StreamingWriter_ConfiguredOptionsPreserveTableSettings() {
            using var output = new MemoryStream();
            var options = new ExcelWriterOptions {
                SheetName = "Configured Rows",
                CreateTable = true,
                TableName = "ConfiguredRows"
            };

            ExcelDataSetImportResult result;

            var destination23 = output;

            var writeOptions23 = options ?? new ExcelWriterOptions {
                UseSharedStrings = false,
                CreateTable = false
            };

            using (var workbookWriter23 = ExcelDocument.CreateWriter(destination23, writeOptions23, leaveOpen: true)) {
                result = workbookWriter23.WriteTable(writeOptions23.SheetName, new[] { new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true) }, ["Id", "Name", "Created", "Active"], static (writer, row) => writer.Write(row.Id).Write(row.Name).Write(row.Created).Write(row.Active));
            }

            if (destination23.CanSeek)
                destination23.Position = 0;

            Assert.False(options.UseSharedStrings);
            Assert.Equal("Configured Rows", result.SheetName);
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            Assert.Null(spreadsheet.WorkbookPart!.SharedStringTablePart);
            var table = spreadsheet.WorkbookPart.WorksheetParts.Single().TableDefinitionParts.Single().Table;
            Assert.Equal("ConfiguredRows", table!.Name?.Value);
        }

        [Fact]
        public void WriteRows_RejectsRowsWithTheWrongCellCount() {
            using var output = new MemoryStream();

            Assert.Throws<InvalidOperationException>(() => {
                var destination24 = output;
                var writeOptions24 = new ExcelWriterOptions {
                    UseSharedStrings = false
                };
                using (var workbookWriter24 = ExcelDocument.CreateWriter(destination24, writeOptions24, leaveOpen: true)) {
                    return workbookWriter24.WriteTable(writeOptions24.SheetName, new[] { new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true) }, ["Id", "Name"], static (writer, row) => writer.Write(row.Id));
                }
            });
        }

        [Fact]
        public void WriteRows_RejectsRowsWithTooManyCells() {
            using var output = new MemoryStream();

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => {
                var destination25 = output;
                var writeOptions25 = new ExcelWriterOptions {
                    UseSharedStrings = false
                };
                using (var workbookWriter25 = ExcelDocument.CreateWriter(destination25, writeOptions25, leaveOpen: true)) {
                    return workbookWriter25.WriteTable(writeOptions25.SheetName, new[] { new TabularWriteRow(1, "Alpha", new DateTime(2026, 7, 10), true) }, ["Id"], static (writer, row) => writer.Write(row.Id).Write(row.Name));
                }
            });

            Assert.Contains("more than 1 cells", exception.Message, StringComparison.Ordinal);
        }

        private static double NextRepresentableDouble(double value) {
            long bits = BitConverter.DoubleToInt64Bits(value);
            return BitConverter.Int64BitsToDouble(bits + 1);
        }

        private sealed record TabularWriteRow(int Id, string? Name, DateTime Created, bool Active);
    }
}
