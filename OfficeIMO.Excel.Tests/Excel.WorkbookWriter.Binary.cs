using OfficeIMO.Excel;
using OfficeIMO.SharedSource.IO;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    /// <summary>Checks saved shared strings across sequential, overlapping and aborted writer lifetimes.</summary>
    [Fact]
    public void WorkbookWriter_XlsKeepsSharedStringsIsolatedAcrossWorkbookLifetimes() {
        string[] warm = { "warm one", "warm two", "warm three", "warm four" };
        AssertStrings(WriteStrings(warm), warm);

        string[] first = { "first one", "first two" };
        string[] second = { "second one", "second two" };
        using var firstOutput = new MemoryStream();
        using var secondOutput = new MemoryStream();
        using (var firstWriter = ExcelDocument.CreateWriter(ExcelFileFormat.Xls, firstOutput, leaveOpen: true)) {
            firstWriter.WriteTable("Data", first, new[] { "Name" }, static (row, value) => row.Write(value));
            using (var secondWriter = ExcelDocument.CreateWriter(ExcelFileFormat.Xls, secondOutput, leaveOpen: true)) {
                secondWriter.WriteTable("Data", second, new[] { "Name" }, static (row, value) => row.Write(value));
            }
        }
        AssertStrings(firstOutput.ToArray(), first);
        AssertStrings(secondOutput.ToArray(), second);

        using (var failedOutput = new MemoryStream())
        using (var failedWriter = ExcelDocument.CreateWriter(ExcelFileFormat.Xls, failedOutput, leaveOpen: true)) {
            var failure = new InvalidOperationException("Synthetic row callback failure.");
            Assert.Same(failure, Assert.Throws<InvalidOperationException>(() => failedWriter.WriteTable(
                "Data", new[] { "aborted private text" }, new[] { "Name" }, (row, value) => {
                    row.Write(value);
                    throw failure;
                })));
        }

        string[] fresh = { "fresh one", "fresh two" };
        AssertStrings(WriteStrings(fresh), fresh);

        static byte[] WriteStrings(string[] values) {
            using var output = new MemoryStream();
            using (var writer = ExcelDocument.CreateWriter(ExcelFileFormat.Xls, output, leaveOpen: true)) {
                writer.WriteTable("Data", values, new[] { "Name" }, static (row, value) => row.Write(value));
            }
            return output.ToArray();
        }

        static void AssertStrings(byte[] package, string[] expected) {

            byte[] workbook = ReadCompoundStream(package, "Workbook");
            uint uniqueStrings = 0;
            for (int offset = 0; offset + 12 <= workbook.Length; offset += 4 + ReadUInt16(workbook, offset + 2)) {
                if (ReadUInt16(workbook, offset) == 0x00fc) {
                    uniqueStrings = ReadUInt32(workbook, offset + 8);
                    break;
                }
            }
            Assert.Equal((uint)(expected.Length + 1), uniqueStrings);

            using (var reader = ExcelDocument.OpenReader(package)) {
                Assert.Equal("Name", reader.GetName(0));
                foreach (string value in expected) {
                    Assert.True(reader.Read());
                    Assert.Equal(value, reader.GetString(0));
                }
                Assert.False(reader.Read());
                Assert.False(reader.NextResult());
            }

        }
    }

    [Fact]
    public void WorkbookWriter_XlsPreservesFullWidthNumericRowsAcrossRowBlocks() {
        const int columns = 256;
        const int rows = 33;
        string[] headers = Enumerable.Range(0, columns).Select(column => "Column" + column).ToArray();
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(ExcelFileFormat.Xls, output, leaveOpen: true)) {
            var result = writer.WriteTable("FullWidth", Enumerable.Range(0, rows), headers, static (row, index) => {
                for (int column = 0; column < columns; column++) row.Write(index * columns + column);
            });
            Assert.Equal("A1:IV34", result.Range);
        }
        output.Position = 0;
        using var reader = ExcelDocument.OpenReader(output);
        Assert.Equal(columns, reader.FieldCount);
        Assert.Equal(headers[columns - 1], reader.GetName(columns - 1));
        for (int index = 0; index < rows; index++) {
            Assert.True(reader.Read());
            for (int column = 0; column < columns; column++) Assert.Equal(index * columns + column, reader.GetInt32(column));
        }
        Assert.False(reader.Read());
        Assert.False(reader.NextResult());
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsx, ExcelDateSystem.NineteenHundred, false)]
    [InlineData(ExcelFileFormat.Xlsx, ExcelDateSystem.NineteenHundred, true)]
    [InlineData(ExcelFileFormat.Xlsx, ExcelDateSystem.NineteenFour, false)]
    [InlineData(ExcelFileFormat.Xlsb, ExcelDateSystem.NineteenHundred, false)]
    [InlineData(ExcelFileFormat.Xlsb, ExcelDateSystem.NineteenHundred, true)]
    [InlineData(ExcelFileFormat.Xlsb, ExcelDateSystem.NineteenFour, false)]
    [InlineData(ExcelFileFormat.Xls, ExcelDateSystem.NineteenHundred, false)]
    [InlineData(ExcelFileFormat.Xls, ExcelDateSystem.NineteenFour, false)]
    public async System.Threading.Tasks.Task WorkbookWriter_FormatsPreserveSinglePassValuesAcrossSheets(ExcelFileFormat format, ExcelDateSystem dateSystem, bool asynchronous) {
        int caller = System.Threading.Thread.CurrentThread.ManagedThreadId;
        int rowCount = format == ExcelFileFormat.Xls ? 33 : 16_384;
        using var output = new NonSeekableReadWriteBuffer(Array.Empty<byte>());
        var rows = new SinglePassRows<int>(Enumerable.Range(0, rowCount));
        var date = new DateTime(2026, 8, 6, 14, 35, 12);
        var offset = new DateTimeOffset(date, TimeSpan.Zero);
        string[] headers = { "Text", "Id", "Date", "Flag", "Blank", "Duration", "Decimal", "Offset" };
        int callbacks = 0;
        ExcelDocument.ExcelTabularRowWriter? retained = null;
        using (var writer = ExcelDocument.CreateWriter(format, output, new ExcelWriterOptions { DateSystem = dateSystem }, leaveOpen: true)) {
            void WriteRow(ExcelDocument.ExcelTabularRowWriter row, int index) {
                if (!asynchronous) Assert.Equal(caller, System.Threading.Thread.CurrentThread.ManagedThreadId);
                Assert.Equal(callbacks++, index);
                retained = row;
                row.Write("Żółć 🚀\r\n&").Write(index).Write(date).Write(index % 2 == 0).WriteBlank()
                    .Write(TimeSpan.FromHours(30)).Write(index + 0.25m).Write((object)offset);
            }
            var result = asynchronous
                ? await writer.WriteTable("Events", AsyncRows(), headers, WriteRow)
                : writer.WriteTable("Events", rows, headers, WriteRow);
            Assert.Equal(("Events", (string?)null, "A1:H" + (rowCount + 1), rowCount, 8),
                (result.SheetName, result.TableName, result.Range, result.RowCount, result.ColumnCount));
            Assert.Throws<InvalidOperationException>(() => retained!.Write(1));
            var empty = writer.WriteTable("Empty", Array.Empty<int>(), new[] { "Value" }, (row, value) => row.Write(value));
            Assert.Equal(0, empty.RowCount);
        }
        Assert.Equal(1, rows.Enumerations);
        Assert.Equal(rowCount, callbacks);
        Assert.True(output.CanWrite);
        using var reader = ExcelDocument.OpenReader(output.ToArray());
        Assert.Equal(8, reader.FieldCount);
        for (int column = 0; column < headers.Length; column++) Assert.Equal(headers[column], reader.GetName(column));
        for (int index = 0; index < rowCount; index++) {
            Assert.True(reader.Read());
            Assert.Equal("Żółć 🚀\r\n&", reader.GetString(0));
            Assert.Equal(index, reader.GetInt32(1));
            Assert.Equal(date, reader.GetDateTime(2));
            Assert.Equal(index % 2 == 0, reader.GetBoolean(3));
            Assert.True(reader.IsDBNull(4));
            Assert.Equal(1.25, reader.GetDouble(5));
            Assert.Equal(index + 0.25, reader.GetDouble(6));
            Assert.Equal(offset.LocalDateTime, reader.GetDateTime(7));
        }
        Assert.False(reader.Read());
        Assert.True(reader.NextResult());
        Assert.Equal("Value", reader.GetName(0));
        Assert.False(reader.Read());
        Assert.False(reader.NextResult());

        async IAsyncEnumerable<int> AsyncRows() {
            foreach (int index in rows) {
                if (index % 1024 == 0) await System.Threading.Tasks.Task.Yield();
                yield return index;
            }
        }

    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsx)]
    [InlineData(ExcelFileFormat.Xlsb)]
    [InlineData(ExcelFileFormat.Xls)]
    public void WorkbookWriter_ReleasesTheRowWhenItsCallbackThrows(ExcelFileFormat format) {
        int rowCount = format == ExcelFileFormat.Xls ? 1 : 8192;
        int callbacks = 0;
        string text = new string('X', 128);
        using var output = new MemoryStream();
        var failure = new InvalidDataException("callback failed");
        ExcelDocument.ExcelTabularRowWriter? retained = null;
        using (var writer = ExcelDocument.CreateWriter(format, output, leaveOpen: true)) {
            Assert.Same(failure, Assert.Throws<InvalidDataException>(() => writer.WriteTable("Data", Enumerable.Range(1, rowCount),
                new[] { "First", "Second" }, (row, value) => {
                    retained = row;
                    row.Write(value);
                    if (format != ExcelFileFormat.Xls) row.Write(text);
                    if (++callbacks == rowCount) throw failure;
                })));
        }
        Assert.Equal(rowCount, callbacks);
        Assert.Throws<InvalidOperationException>(() => retained!.Write(2));
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsb, nameof(ExcelWriterOptions.CreateTable))]
    [InlineData(ExcelFileFormat.Xls, nameof(ExcelWriterOptions.CreateTable))]
    [InlineData(ExcelFileFormat.Xlsb, nameof(ExcelWriterOptions.FreezeHeader))]
    [InlineData(ExcelFileFormat.Xls, nameof(ExcelWriterOptions.FreezeHeader))]
    [InlineData(ExcelFileFormat.Xlsb, nameof(ExcelWriterOptions.FreezeFirstColumn))]
    [InlineData(ExcelFileFormat.Xls, nameof(ExcelWriterOptions.FreezeFirstColumn))]
    public void WorkbookWriter_BinaryOptionsRejectUnsupportedRequestsBeforeWriting(ExcelFileFormat format, string option) {
        using var output = new MemoryStream();
        var options = option switch {
            nameof(ExcelWriterOptions.CreateTable) => new ExcelWriterOptions { CreateTable = true },
            nameof(ExcelWriterOptions.FreezeHeader) => new ExcelWriterOptions { FreezeHeader = true },
            nameof(ExcelWriterOptions.FreezeFirstColumn) => new ExcelWriterOptions { FreezeFirstColumn = true },
            _ => throw new ArgumentOutOfRangeException(nameof(option))
        };
        Assert.Throws<NotSupportedException>(() => ExcelDocument.CreateWriter(format, output,
            options, leaveOpen: true));
        if (format == ExcelFileFormat.Xls && option == nameof(ExcelWriterOptions.CreateTable)) {
            Assert.Throws<NotSupportedException>(() => ExcelDocument.CreateWriter(format, output,
                new ExcelWriterOptions { RequireStreaming = true }, leaveOpen: true));
        }
        Assert.Equal(0, output.Length);
    }


    [Theory]
    [InlineData(ExcelFileFormat.Xlsx, "_x")]
    [InlineData(ExcelFileFormat.Xlsb, "_x")]
    [InlineData(ExcelFileFormat.Xls, "_x")]
    [InlineData(ExcelFileFormat.Xlsx, "'x")]
    [InlineData(ExcelFileFormat.Xlsb, "'x")]
    [InlineData(ExcelFileFormat.Xls, "'x")]
    [InlineData(ExcelFileFormat.Xlsx, " x")]
    [InlineData(ExcelFileFormat.Xlsb, " x")]
    [InlineData(ExcelFileFormat.Xls, " x")]
    [InlineData(ExcelFileFormat.Xlsx, "_'_")]
    [InlineData(ExcelFileFormat.Xlsb, "_'_")]
    [InlineData(ExcelFileFormat.Xls, "_'_")]
    [InlineData(ExcelFileFormat.Xlsx, "'   '")]
    [InlineData(ExcelFileFormat.Xlsb, "'   '")]
    [InlineData(ExcelFileFormat.Xls, "'   '")]
    public void WorkbookWriter_PreservesResolvedNamesInResultsAndPackages(ExcelFileFormat format, string suffix) {
        string stem = suffix is "_'_" or "'   '" ? "" : new string('A', 30);
        string expected = suffix == "_'_" ? "_'_" : suffix == "'   '" ? "Sheet1" : suffix == "_x" ? stem + "_" : stem;
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(format, output, new ExcelWriterOptions(), leaveOpen: true)) {
            var first = writer.WriteTable(stem + suffix, new[] { 41 }, new[] { "Id" }, static (row, value) => row.Write(value));
            Assert.Equal(expected, first.SheetName);
            if (suffix == "_x") {
                var second = writer.WriteTable(stem, new[] { 42 }, new[] { "Id" }, static (row, value) => row.Write(value));
                Assert.Equal(stem, second.SheetName);
            }
        }
        byte[] bytes = output.ToArray();
        using var reader = ExcelDocument.OpenReader(bytes, new ExcelReaderOptions { SheetName = expected });
        Assert.True(reader.Read());
        Assert.Equal(41, reader.GetInt32(0));
        Assert.False(reader.Read());
        if (suffix == "_x") {
            using var second = ExcelDocument.OpenReader(bytes, new ExcelReaderOptions { SheetName = stem });
            Assert.True(second.Read());
            Assert.Equal(42, second.GetInt32(0));
            Assert.False(second.Read());
        }
    }


    [Theory]
    [InlineData(ExcelFileFormat.Xlsb)]
    [InlineData(ExcelFileFormat.Xls)]
    public void WorkbookWriter_BinaryIncrementalTableUsesTheFinalUnionSchema(ExcelFileFormat format) {
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(format, output, leaveOpen: true)) {
            using var table = writer.BeginTable("Growing");
            table.WriteRow(new[] { "Id" }, new object?[] { 1 });
            table.WriteRow(new[] { "Name", "Id" }, new object?[] { "later", 2 });
            var result = table.Complete();
            Assert.Equal(("A1:B3", 2, 2), (result.Range, result.RowCount, result.ColumnCount));
        }
        using var reader = ExcelDocument.OpenReader(output.ToArray());
        Assert.Equal("Id", reader.GetName(0));
        Assert.Equal("Name", reader.GetName(1));
        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.True(reader.IsDBNull(1));
        Assert.True(reader.Read());
        Assert.Equal(2, reader.GetInt32(0));
        Assert.Equal("later", reader.GetString(1));
        Assert.False(reader.Read());
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsx)]
    [InlineData(ExcelFileFormat.Xlsb)]
    [InlineData(ExcelFileFormat.Xls)]
    public void WorkbookWriter_PreservesCallbackFailureWhenDestinationCleanupAlsoFails(ExcelFileFormat format) {
        using var output = new OfficeIMO.Tests.TestStreams.ThrowAfterBytesWriteStream(0);
        var failure = new InvalidDataException("Synthetic callback failure.");
        var writer = ExcelDocument.CreateWriter(format, output, leaveOpen: true);
        int callbacks = 0;
        ExcelDocument.ExcelTabularRowWriter? retained = null;
        try {
            Assert.Same(failure, Assert.Throws<InvalidDataException>(() => writer.WriteTable(
                "Data", new[] { 1 }, new[] { "Value" }, (row, value) => {
                    callbacks++;
                    retained = row;
                    row.Write(value);
                    throw failure;
                })));
            Assert.Equal(1, callbacks);
            Assert.Throws<InvalidOperationException>(() => retained!.Write(2));
        } finally {
            try { writer.Dispose(); }
            catch (IOException) { /* The failed destination may also reject the ZIP trailer. */ }
        }
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsx)]
    [InlineData(ExcelFileFormat.Xlsb)]
    public void WorkbookWriter_LargeDestinationFailureStopsTheSinglePassSource(ExcelFileFormat format) {
        const int rowCount = 256;
        using var output = new OfficeIMO.Tests.TestStreams.ThrowAfterBytesWriteStream(2 * 1024 * 1024);
        var random = new Random(42);
        bool disposed = false;
        int advances = 0;
        var writer = ExcelDocument.CreateWriter(format, output, leaveOpen: true);
        try {
            Assert.Throws<IOException>(() => writer.WriteTable("Data", Rows(), new[] { "Text" },
                static (row, value) => row.Write(value)));
            Assert.True(disposed);
            Assert.InRange(advances, 1, rowCount - 1);
            Assert.InRange(output.Length, 1024L * 1024, 2L * 1024 * 1024);
            Assert.Throws<InvalidOperationException>(() => writer.WriteTable("Later", new[] { "unused" },
                new[] { "Text" }, static (row, value) => row.Write(value)));
        } finally {
            try { writer.Dispose(); }
            catch (IOException) { /* The failed destination may also reject the ZIP trailer. */ }
        }

        IEnumerable<string> Rows() {
            try {
                for (int index = 0; index < rowCount; index++) {
                    var characters = new char[16_384];
                    for (int character = 0; character < characters.Length; character++)
                        characters[character] = (char)random.Next(32, 0xD800);
                    advances++;
                    yield return new string(characters);
                }
            } finally { disposed = true; }
        }
    }

}
