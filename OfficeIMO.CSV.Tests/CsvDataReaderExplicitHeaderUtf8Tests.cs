#if NET8_0_OR_GREATER
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvDataReaderExplicitHeaderUtf8Tests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ExplicitHeaderKeepsFirstRecordPreambleCursorAndCallerOwnership(bool hasHeaderRow, bool preamble) {
        byte[] prefix = Encoding.UTF8.GetBytes("ignored prefix");
        byte[] payload = (preamble ? Encoding.UTF8.GetPreamble() : Array.Empty<byte>())
            .Concat(Encoding.UTF8.GetBytes("Zażółć 😀,1\nnext,2\n")).ToArray();
        using var stream = new MemoryStream(prefix.Concat(payload).ToArray());
        stream.Position = prefix.Length;
        var options = new CsvLoadOptions {
            Header = new[] { "Name", "Number" }, HasHeaderRow = hasHeaderRow,
            MaxInputBytes = payload.Length,
        };
        using (var reader = CsvDocument.OpenDataReader(stream, options)) {
            options.Header[0] = "changed after opening";
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("Name", reader.GetName(0));
            Assert.Equal(1, reader.GetOrdinal("Number"));
            Assert.True(reader.HasRows);
            Assert.True(reader.HasRows);
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
            Assert.True(reader.Read());
            Assert.True(reader.TryGetUtf8Text(0, out var first));
            Assert.Equal("Zażółć 😀", Encoding.UTF8.GetString(first));
            Assert.Equal("Zażółć 😀", reader.GetString(0));
            Assert.Equal(1, reader.GetInt32(1));
            Assert.True(reader.Read());
            Assert.True(reader.TryGetUtf8Text(0, out var next));
            Assert.Equal("next", Encoding.UTF8.GetString(next));
            Assert.Equal(2, reader.GetInt32(1));
            Assert.False(reader.Read());
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        }
        Assert.True(stream.CanRead);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void ExplicitWidthIgnoresExtraFieldsAndDistinguishesMissingFromEmpty(int headerWidth) {
        string[] header = Enumerable.Range(1, headerWidth).Select(index => "Field" + index).ToArray();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("first,second\nnext,\nlast\n"));
        using var reader = CsvDocument.OpenDataReader(stream,
            new CsvLoadOptions { Header = header, HasHeaderRow = false });
        Assert.Equal(headerWidth, reader.FieldCount);
        string[][] rows = { new[] { "first", "second" }, new[] { "next", "" }, new[] { "last" } };
        foreach (string[] row in rows) {
            Assert.True(reader.Read());
            for (int ordinal = 0; ordinal < headerWidth; ordinal++) {
                Assert.Equal(header[ordinal], reader.GetName(ordinal));
                Assert.Equal(ordinal < row.Length ? row[ordinal] : string.Empty, reader.GetString(ordinal));
                Assert.False(reader.IsDBNull(ordinal));
                bool borrowed = reader.TryGetUtf8Text(ordinal, out var text);
                Assert.Equal(ordinal < row.Length, borrowed);
                Assert.Equal(ordinal < row.Length ? row[ordinal] : string.Empty, Encoding.UTF8.GetString(text));
            }
            Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(headerWidth, out _));
        }
        Assert.False(reader.Read());
    }

    [Fact]
    public void EmptyInputKeepsExplicitColumnMetadata() {
        using var stream = new MemoryStream();
        using var reader = CsvDocument.OpenDataReader(stream,
            new CsvLoadOptions { Header = new[] { "Name", "Amount" }, HasHeaderRow = false });
        Assert.Equal(2, reader.FieldCount);
        Assert.Equal("Name", reader.GetName(0));
        Assert.Equal("Amount", reader.GetName(1));
        Assert.Equal(typeof(string), reader.GetFieldType(1));
        Assert.False(reader.HasRows);
        Assert.False(reader.Read());
    }

    [Fact]
    public void ExplicitHeaderFileBorrowsTheFirstRecordAndReleasesTheOwnedStream() {
        string path = Path.Combine(Path.GetTempPath(), "csv-explicit-header-" + Guid.NewGuid().ToString("N") + ".csv");
        try {
            File.WriteAllText(path, "first,1\nnext,2\n", new UTF8Encoding(false));
            using (var reader = CsvDocument.OpenDataReader(path,
                       new CsvLoadOptions { Header = new[] { "Name", "Number" }, HasHeaderRow = false })) {
                Assert.Equal("Name", reader.GetName(0));
                Assert.True(reader.Read());
                Assert.True(reader.TryGetUtf8Text(0, out var first));
                Assert.Equal("first", Encoding.UTF8.GetString(first));
                Assert.Equal(1, reader.GetInt32(1));
                Assert.True(reader.Read());
                Assert.Equal("next", reader.GetString(0));
                Assert.False(reader.Read());
            }
            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.CanRead);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitNamesUseExistingMissingAndDuplicateHeaderPolicy(bool preserve) {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("1,2,3,4,5,6\n"));
        using var reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions {
            Header = new[] { null!, "", "Value", "value", "Value_2", "H1" },
            GenerateMissingHeaderNames = !preserve,
            DuplicateHeaderBehavior = preserve ? CsvDuplicateHeaderBehavior.Preserve : CsvDuplicateHeaderBehavior.Rename,
        });
        string[] expected = preserve
            ? new[] { "", "", "Value", "value", "Value_2", "H1" }
            : new[] { "H2", "H3", "Value", "value_3", "Value_2", "H1" };
        Assert.Equal(expected, Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToArray());
        Assert.Equal(2, reader.GetOrdinal("Value"));
        Assert.True(reader.Read());
        for (int ordinal = 0; ordinal < expected.Length; ordinal++) {
            Assert.True(reader.TryGetUtf8Text(ordinal, out var text));
            Assert.Equal((ordinal + 1).ToString(), Encoding.UTF8.GetString(text));
        }
        Assert.False(reader.Read());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InvalidExplicitHeadersStillRejectAndLeaveCallerStreamOpen(bool empty) {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("first,second\n"));
        var options = new CsvLoadOptions {
            Header = empty ? Array.Empty<string>() : new[] { "Value", "value" },
            DuplicateHeaderBehavior = CsvDuplicateHeaderBehavior.Throw,
        };
        if (empty) Assert.Throws<ArgumentException>(() => CsvDocument.OpenDataReader(stream, options));
        else Assert.Throws<CsvException>(() => CsvDocument.OpenDataReader(stream, options));
        Assert.True(stream.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitHeadersKeepCommentAndW3CRecordsUnderTheCommentPolicy(bool skipComments) {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("#comment,first\n#Fields: raw,second\nlast,third\n"));
        using var reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions {
            Header = new[] { "Name", "Value" }, HasHeaderRow = true,
            SkipCommentRowsBeforeHeader = true, SkipCommentRows = skipComments,
            RecognizeW3CFieldsHeader = true,
        });
        Assert.Equal("Name", reader.GetName(0));
        Assert.Equal("Value", reader.GetName(1));
        string[][] expected = skipComments
            ? new[] { new[] { "last", "third" } }
            : new[] { new[] { "#comment", "first" }, new[] { "#Fields: raw", "second" }, new[] { "last", "third" } };
        foreach (string[] row in expected) {
            Assert.True(reader.Read());
            Assert.Equal(row[0], reader.GetString(0));
            Assert.Equal(row[1], reader.GetString(1));
        }
        Assert.False(reader.Read());
    }

    [Fact]
    public void QuotedFirstRecordUsesCanonicalNormalizationWithoutLosingTheRecord() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("\"a\"\"b\nend\",1\nlast,2\n"));
        using var reader = CsvDocument.OpenDataReader(stream,
            new CsvLoadOptions { Header = new[] { "Name", "Number" }, HasHeaderRow = true });
        Assert.True(reader.HasRows);
        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(0, out _));
        Assert.Equal("a\"b\nend", reader.GetString(0));
        Assert.Equal(1, reader.GetInt32(1));
        Assert.True(reader.Read());
        Assert.Equal("last", reader.GetString(0));
        Assert.Equal(2, reader.GetInt32(1));
        Assert.False(reader.Read());
    }

    [Theory]
    [InlineData("first\n", false)]
    [InlineData("first,1,extra\n", false)]
    [InlineData("first,1\nnext\n", true)]
    [InlineData("first,1\nnext,2,extra\n", true)]
    public void StrictWidthChecksFirstAndSubsequentDataRecords(string text, bool validFirst) {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(text));
        using var reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions {
            Header = new[] { "Name", "Number" }, ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict,
        });
        if (validFirst) {
            Assert.True(reader.Read());
            Assert.Equal("first", reader.GetString(0));
        }
        Assert.Throws<CsvException>(() => reader.Read());
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void ExplicitHeadersKeepNullMarkersAndTrimmedTextOutOfBorrowedBytes() {
        using var nullStream = new MemoryStream(Encoding.UTF8.GetBytes("NULL,,text\nnext\n"));
        using var nullReader = CsvDocument.OpenDataReader(nullStream,
            new CsvLoadOptions { Header = new[] { "A", "B", "C" }, NullValue = "NULL" });
        Assert.True(nullReader.Read());
        Assert.True(nullReader.IsDBNull(0));
        Assert.False(nullReader.IsDBNull(1));
        Assert.Equal(string.Empty, nullReader.GetString(1));
        Assert.False(nullReader.TryGetUtf8Text(0, out _));
        Assert.False(nullReader.TryGetUtf8Text(2, out _));
        Assert.True(nullReader.Read());
        Assert.Equal(string.Empty, nullReader.GetString(1));
        Assert.False(nullReader.IsDBNull(1));
        using var trimmedStream = new MemoryStream(Encoding.UTF8.GetBytes("  Ada  \n"));
        using var trimmed = CsvDocument.OpenDataReader(trimmedStream,
            new CsvLoadOptions { Header = new[] { "Name" }, TrimWhitespace = true });
        Assert.True(trimmed.Read());
        Assert.Equal("Ada", trimmed.GetString(0));
        Assert.False(trimmed.TryGetUtf8Text(0, out _));
    }

    [Fact]
    public void SchemaConvertersAndSkippedRecordsRemainOnTheirExistingProjectionPaths() {
        int calls = 0;
        var schema = new CsvSchemaBuilder().Column("Name").AsString()
            .ConvertUsing(value => { calls++; return "mapped:" + value; }).Done().Build();
        using var schemaStream = new MemoryStream(Encoding.UTF8.GetBytes("Ada\n"));
        using var converted = CsvDocument.OpenDataReader(schemaStream,
            new CsvLoadOptions { Header = new[] { "Name" } }, new CsvDataReaderOptions { Schema = schema });
        Assert.True(converted.Read());
        int before = calls;
        Assert.False(converted.TryGetUtf8Text(0, out _));
        Assert.Equal(before, calls);
        Assert.Equal("mapped:Ada", converted.GetString(0));
        using var skippedStream = new MemoryStream(Encoding.UTF8.GetBytes("ignored\nkept\n"));
        using var skipped = CsvDocument.OpenDataReader(skippedStream,
            new CsvLoadOptions { Header = new[] { "Name" }, SkipInitialRecords = 1 });
        Assert.True(skipped.Read());
        Assert.Equal("kept", skipped.GetString(0));
        Assert.False(skipped.TryGetUtf8Text(0, out _));
        Assert.False(skipped.Read());
    }

    [Fact]
    public void ExplicitHeaderStreamingRetainsInputBoundsFieldLimitsAndCancellation() {
        using var boundedStream = new MemoryStream(Encoding.UTF8.GetBytes("first,1\n"));
        Assert.Throws<InvalidDataException>(() => CsvDocument.OpenDataReader(boundedStream,
            new CsvLoadOptions { Header = new[] { "Name", "Number" }, MaxInputBytes = 2 }));
        Assert.True(boundedStream.CanRead);
        using var fieldStream = new MemoryStream(Encoding.UTF8.GetBytes("longvalue\n"));
        using var fieldReader = CsvDocument.OpenDataReader(fieldStream,
            new CsvLoadOptions { Header = new[] { "Name" }, MaxFieldLength = 4 });
        Assert.Throws<CsvParseException>(() => fieldReader.Read());
        using var cancellation = new CancellationTokenSource();
        using var canceledStream = new MemoryStream(Encoding.UTF8.GetBytes("first,1\nnext,2\n"));
        using var canceled = CsvDocument.OpenDataReader(canceledStream,
            new CsvLoadOptions { Header = new[] { "Name", "Number" }, CancellationToken = cancellation.Token });
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => canceled.Read());
        Assert.True(canceledStream.CanRead);
    }
}
#endif
