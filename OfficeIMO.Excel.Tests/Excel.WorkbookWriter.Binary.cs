using OfficeIMO.Excel;
using OfficeIMO.SharedSource.IO;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(ExcelFileFormat.Xlsb, ExcelDateSystem.NineteenHundred)]
    [InlineData(ExcelFileFormat.Xlsb, ExcelDateSystem.NineteenFour)]
    [InlineData(ExcelFileFormat.Xls, ExcelDateSystem.NineteenHundred)]
    [InlineData(ExcelFileFormat.Xls, ExcelDateSystem.NineteenFour)]
    public void WorkbookWriter_BinaryFormatsPreserveSinglePassValuesAcrossRowBlocksAndSheets(ExcelFileFormat format, ExcelDateSystem dateSystem) {
        using var output = new NonSeekableReadWriteBuffer(Array.Empty<byte>());
        var rows = new SinglePassRows<int>(Enumerable.Range(0, 33));
        var date = new DateTime(2026, 8, 6, 14, 35, 12);
        var offset = new DateTimeOffset(date, TimeSpan.Zero);
        string[] headers = { "Text", "Id", "Date", "Flag", "Blank", "Duration", "Decimal", "Offset" };
        int callbacks = 0;
        ExcelDocument.ExcelTabularRowWriter? retained = null;
        using (var writer = ExcelDocument.CreateWriter(format, output, new ExcelWriterOptions { DateSystem = dateSystem }, leaveOpen: true)) {
            var result = writer.WriteTable("Events", rows, headers, (row, index) => {
                Assert.Equal(callbacks++, index);
                retained = row;
                row.Write("Żółć 🚀\r\n&").Write(index).Write(date).Write(index % 2 == 0).WriteBlank()
                    .Write(TimeSpan.FromHours(30)).Write(index + 0.25m).Write((object)offset);
            });
            Assert.Equal(("Events", (string?)null, "A1:H34", 33, 8),
                (result.SheetName, result.TableName, result.Range, result.RowCount, result.ColumnCount));
            Assert.Throws<InvalidOperationException>(() => retained!.Write(1));
            var empty = writer.WriteTable("Empty", Array.Empty<int>(), new[] { "Value" }, (row, value) => row.Write(value));
            Assert.Equal(0, empty.RowCount);
        }
        Assert.Equal(1, rows.Enumerations);
        Assert.Equal(33, callbacks);
        Assert.True(output.CanWrite);
        using var reader = ExcelDocument.OpenReader(output.ToArray());
        Assert.Equal(8, reader.FieldCount);
        for (int column = 0; column < headers.Length; column++) Assert.Equal(headers[column], reader.GetName(column));
        for (int index = 0; index < 33; index++) {
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
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsx)]
    [InlineData(ExcelFileFormat.Xlsb)]
    [InlineData(ExcelFileFormat.Xls)]
    public void WorkbookWriter_ReleasesTheRowWhenItsCallbackThrows(ExcelFileFormat format) {
        using var output = new MemoryStream();
        var failure = new InvalidDataException("callback failed");
        ExcelDocument.ExcelTabularRowWriter? retained = null;
        using (var writer = ExcelDocument.CreateWriter(format, output, leaveOpen: true)) {
            Assert.Same(failure, Assert.Throws<InvalidDataException>(() => writer.WriteTable("Data", new[] { 1 },
                new[] { "First", "Second" }, (row, value) => {
                    retained = row;
                    row.Write(value);
                    throw failure;
                })));
        }
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

}
