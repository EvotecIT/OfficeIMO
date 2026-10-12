#if NET8_0_OR_GREATER
using System;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvDataReaderUtf8TextTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BorrowedTextHonorsPreambleCurrentPositionAndCallerOwnership(bool preamble) {
        byte[] prefix = Encoding.UTF8.GetBytes("ignored prefix");
        byte[] payload = (preamble ? Encoding.UTF8.GetPreamble() : Array.Empty<byte>())
            .Concat(Encoding.UTF8.GetBytes("Name,Empty\nZażółć 😀,\nnext,last\n")).ToArray();
        using var stream = new MemoryStream(prefix.Concat(payload).ToArray());
        stream.Position = prefix.Length;
        using (DbDataReader reader = CsvDocument.OpenDataReader(stream,
                   new CsvLoadOptions { MaxInputBytes = payload.Length })) {
            Assert.Equal("Name", reader.GetName(0));
            Assert.True(reader.Read());
            Assert.True(reader.TryGetUtf8Text(0, out var text));
            Assert.Equal("Zażółć 😀", Encoding.UTF8.GetString(text));
            Assert.Equal("Zażółć 😀", reader.GetString(0));
            Assert.True(reader.TryGetUtf8Text(1, out var empty));
            Assert.True(empty.IsEmpty);
            Assert.True(reader.Read());
            Assert.True(reader.TryGetUtf8Text(0, out var next));
            Assert.Equal("next", Encoding.UTF8.GetString(next));
            Assert.False(reader.Read());
        }
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void BorrowedTextValidatesCursorAndDistinguishesMissingFields() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("A,B\nfirst\n"));
        using var reader = CsvDocument.OpenDataReader(stream);
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        Assert.True(reader.Read());
        Assert.True(reader.TryGetUtf8Text(0, out var first));
        Assert.Equal("first", Encoding.UTF8.GetString(first));
        Assert.False(reader.TryGetUtf8Text(1, out var missing));
        Assert.True(missing.IsEmpty);
        Assert.Equal(string.Empty, reader.GetString(1));
        Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(-1, out _));
        Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(2, out _));
        Assert.False(reader.Read());
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        reader.Close();
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
    }

    [Fact]
    public void MalformedUtf8RemainsDecodedTextAndStrictDecoderStillRejectsIt() {
        byte[] data = Encoding.UTF8.GetBytes("Name\n").Concat(new byte[] { 0xC3, 0x28, 0x0A }).ToArray();
        using var stream = new MemoryStream(data);
        using var reader = CsvDocument.OpenDataReader(stream);
        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(0, out var text));
        Assert.True(text.IsEmpty);
        Assert.Equal("�(", reader.GetString(0));
        using var strictStream = new MemoryStream(data);
        using var strictReader = CsvDocument.OpenDataReader(strictStream,
            new CsvLoadOptions { Encoding = new UTF8Encoding(false, true) });
        Assert.Throws<DecoderFallbackException>(() => strictReader.Read());
    }

    [Fact]
    public void QuotedFallbackKeepsNormalizedEscapedAndMultilineText() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Name\nplain\n\"a\"\"b\nend\"\nlast\n"));
        using var reader = CsvDocument.OpenDataReader(stream);
        Assert.True(reader.Read());
        Assert.True(reader.TryGetUtf8Text(0, out _));
        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(0, out _));
        Assert.Equal("a\"b\nend", reader.GetString(0));
        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(0, out _));
        Assert.Equal("last", reader.GetString(0));
    }

    [Fact]
    public void Utf16PreambleAndTrimUseCanonicalTextFallback() {
        byte[] data = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("Name\nZażółć\n")).ToArray();
        using var stream = new MemoryStream(data);
        using var reader = CsvDocument.OpenDataReader(stream);
        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(0, out _));
        Assert.Equal("Zażółć", reader.GetString(0));
        using var trimmedStream = new MemoryStream(Encoding.UTF8.GetBytes("Name\n  Ada  \n"));
        using var trimmed = CsvDocument.OpenDataReader(trimmedStream, new CsvLoadOptions { TrimWhitespace = true });
        Assert.True(trimmed.Read());
        Assert.False(trimmed.TryGetUtf8Text(0, out _));
        Assert.Equal("Ada", trimmed.GetString(0));
    }

    [Fact]
    public void SchemaConverterAndNullMarkerDoNotExposeSourceBytes() {
        int calls = 0;
        CsvSchema schema = new CsvSchemaBuilder().Column("Name").AsString()
            .ConvertUsing(value => { calls++; return "mapped:" + value; }).Done().Build();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Name\nAda\n"));
        using var reader = CsvDocument.OpenDataReader(stream, readerOptions: new CsvDataReaderOptions { Schema = schema });
        Assert.True(reader.Read());
        int before = calls;
        Assert.False(reader.TryGetUtf8Text(0, out _));
        Assert.Equal(before, calls);
        Assert.Equal("mapped:Ada", reader.GetString(0));
        using var nullStream = new MemoryStream(Encoding.UTF8.GetBytes("Name\nNULL\n"));
        using var nullReader = CsvDocument.OpenDataReader(nullStream, new CsvLoadOptions { NullValue = "NULL" });
        Assert.True(nullReader.Read());
        Assert.False(nullReader.TryGetUtf8Text(0, out _));
        Assert.True(nullReader.IsDBNull(0));
    }

    [Fact]
    public void HeaderlessStreamRetainsFirstRowWidthAndCursorState() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("first,1\nnext,2\n"));
        using var reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions { HasHeaderRow = false });
        Assert.Equal(2, reader.FieldCount);
        Assert.Equal("Column1", reader.GetName(0));
        Assert.True(reader.HasRows);
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        Assert.True(reader.Read());
        Assert.True(reader.TryGetUtf8Text(0, out var first));
        Assert.Equal("first", Encoding.UTF8.GetString(first));
        Assert.Equal(1, reader.GetInt32(1));
        Assert.True(reader.Read());
        Assert.Equal("next", reader.GetString(0));
        Assert.Equal(2, reader.GetInt32(1));
        Assert.False(reader.Read());
        using var emptyStream = new MemoryStream();
        using var empty = CsvDocument.OpenDataReader(emptyStream, new CsvLoadOptions { HasHeaderRow = false });
        Assert.Equal(0, empty.FieldCount);
        Assert.False(empty.HasRows);
        Assert.False(empty.Read());
    }

    [Fact]
    public void HeaderlessStreamingKeepsWidthLimitsAndCancellation() {
        using var strictStream = new MemoryStream(Encoding.UTF8.GetBytes("first,1\nnext\n"));
        using var strict = CsvDocument.OpenDataReader(strictStream, new CsvLoadOptions {
            HasHeaderRow = false, ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict,
        });
        Assert.True(strict.Read());
        Assert.Throws<CsvException>(() => strict.Read());
        using var boundedStream = new MemoryStream(Encoding.UTF8.GetBytes("first,1\n"));
        Assert.Throws<InvalidDataException>(() => CsvDocument.OpenDataReader(boundedStream,
            new CsvLoadOptions { HasHeaderRow = false, MaxInputBytes = 2 }));
        Assert.True(boundedStream.CanRead);
        using var fieldStream = new MemoryStream(Encoding.UTF8.GetBytes("Name\nlongvalue\n"));
        using var fieldReader = CsvDocument.OpenDataReader(fieldStream, new CsvLoadOptions { MaxFieldLength = 4 });
        Assert.Throws<CsvParseException>(() => fieldReader.Read());
        using var cancellation = new CancellationTokenSource();
        using var canceledStream = new MemoryStream(Encoding.UTF8.GetBytes("first,1\nnext,2\n"));
        using var canceled = CsvDocument.OpenDataReader(canceledStream,
            new CsvLoadOptions { HasHeaderRow = false, CancellationToken = cancellation.Token });
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => canceled.Read());
        Assert.True(canceledStream.CanRead);
    }
}
#endif
