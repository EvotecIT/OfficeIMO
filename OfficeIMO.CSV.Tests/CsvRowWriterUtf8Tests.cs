#if NET8_0_OR_GREATER
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvRowWriterUtf8Tests
{
    public static object[][] SerializationOptions => new object[][]
    {
        new object[] { new CsvSaveOptions { NewLine = "\n" } },
        new object[] { new CsvSaveOptions { NewLine = "\r\n", QuoteMode = CsvQuoteMode.Always } },
        new object[] { new CsvSaveOptions { QuoteMode = CsvQuoteMode.Never, QuoteFields = new[] { "Missing", "Empty" } } },
        new object[] { new CsvSaveOptions { QuoteFields = new[] { "Missing", "Empty", "Text" } } },
        new object[] { new CsvSaveOptions { DelimiterText = "|λ|", NewLine = "\n", IncludeHeader = false } },
        new object[] { new CsvSaveOptions { Delimiter = ';', NullValue = "=NULL;", FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape } },
        new object[] { new CsvSaveOptions { Delimiter = '\'', FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape } },
        new object[] { new CsvSaveOptions { DelimiterText = "'=", FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape } },
        new object[] { new CsvSaveOptions { Encoding = Encoding.Unicode, NullValue = "NULL" } },
        new object[] { new CsvSaveOptions { Encoding = new UTF8Encoding(true), QuoteMode = CsvQuoteMode.Always } },
        new object[] { new CsvSaveOptions { CompressionType = CsvCompressionType.GZip } },
        new object[] { new CsvSaveOptions { CompressionType = CsvCompressionType.Deflate } },
        new object[] { new CsvSaveOptions { CompressionType = CsvCompressionType.Brotli } },
        new object[] { new CsvSaveOptions { CompressionType = CsvCompressionType.ZLib } }
    };

    [Theory]
    [MemberData(nameof(SerializationOptions))]
    public void Utf8Rows_MatchTextRowsAcrossSerializationOptions(CsvSaveOptions options)
    {
        string[] columns = { "Missing", "Empty", "Text", "Formula", "Unicode", "Long" };
        string?[][] rows = {
            new string?[] { null, "", "a,;|λ|b\"\"\"\"\r\nc\0", "  =SUM(A1)", "Zażółć 🚀 λ \t", new string('x', 33000) + "🚀" },
            new string?[] { "missing", null, "\"quoted\"", "\t=1", "\u00a0=preserved", "after large value" }
        };
        byte[] expected = WriteText(columns, rows, options);
        using var output = new MemoryStream();
        using (var writer = CsvRowWriter.CreateStream(output, options, leaveOpen: true, bufferSize: 16))
        {
            writer.WriteUtf8Row(columns, Encode(rows[0]));
            var second = Encode(rows[1]);
            writer.WriteUtf8Row(second.Length, second, static (values, i) => values[i]);
        }
        Assert.Equal(DecodeCompression(expected, options.CompressionType), DecodeCompression(output.ToArray(), options.CompressionType));
    }

    [Fact]
    public void Utf8Rows_InvalidEncodingDoesNotWriteHeaderOrPartialRow()
    {
        string[] columns = { "A", "B" };
        ReadOnlyMemory<byte>?[] invalid = { Encoding.UTF8.GetBytes("would be written"), new byte[] { 0xF0, 0x80, 0x80, 0x80 } };
        using var output = new MemoryStream();
        using (var writer = CsvRowWriter.CreateStream(output, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true))
        {
            Assert.Throws<DecoderFallbackException>(() => writer.WriteUtf8Row(columns, invalid));
            Assert.Equal(0, output.Length);
            writer.WriteUtf8Row(columns, Encode(new[] { "first", "row" }));
            Assert.Throws<DecoderFallbackException>(() => writer.WriteUtf8Row(invalid));
            Assert.Throws<CsvException>(() => writer.WriteUtf8Row(columns, new ReadOnlyMemory<byte>?[1]));
            Assert.Throws<CsvException>(() => writer.WriteUtf8Row(new[] { "B", "A" }, Encode(new[] { "wrong", "order" })));
            writer.WriteUtf8Row(Encode(new[] { "second", "row" }));
        }
        Assert.Equal("A,B\nfirst,row\nsecond,row\n", Encoding.UTF8.GetString(output.ToArray()));
    }

    [Fact]
    public void Utf8Rows_AccessorIsConsumedOnceAndDoesNotRetainBorrowedMemory()
    {
        using var output = new MemoryStream();
        byte[] borrowed = Encoding.UTF8.GetBytes("original");
        int calls = 0;
        using (var writer = CsvRowWriter.CreateStream(output, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true, bufferSize: 16))
        {
            Assert.Throws<DecoderFallbackException>(() => writer.WriteUtf8Row(new[] { "Bad" }, 1, 0, static (_, _) => new byte[] { 0xFF }));
            writer.WriteUtf8Row(new[] { "A", "B" }, 2, borrowed, (value, index) => { calls++; return index == 0 ? value : ReadOnlyMemory<byte>.Empty; });
            Array.Fill(borrowed, (byte)'X');
            writer.WriteTextRow(new[] { "text", "Zażółć 🚀" });
            writer.WriteUtf8Row(2, 0, static (_, index) => index == 0 ? Encoding.UTF8.GetBytes("bytes") : null);
        }
        Assert.Equal(2, calls);
        Assert.Equal("A,B\noriginal,\ntext,Zażółć 🚀\nbytes,\n", Encoding.UTF8.GetString(output.ToArray()));
    }

    [Fact]
    public void Utf8Rows_TextWriterDestinationMatchesExistingTextApi()
    {
        var options = new CsvSaveOptions { DelimiterText = "λ|", NullValue = "NULL", QuoteMode = CsvQuoteMode.Always };
        string[] columns = { "A", "B" };
        string?[] values = { "🚀 \"value\"", null };
        using var text = new StringWriter();
        using var utf8 = new StringWriter();
        using (var writer = new CsvRowWriter(text, options, leaveOpen: true)) writer.WriteTextRow(columns, values);
        using (var writer = new CsvRowWriter(utf8, options, leaveOpen: true))
        {
            Assert.Throws<DecoderFallbackException>(() => writer.WriteUtf8Row(columns, new ReadOnlyMemory<byte>?[] { new byte[] { 0x80 }, null }));
            Assert.Equal(string.Empty, utf8.ToString());
            writer.WriteUtf8Row(columns, Encode(values));
        }
        Assert.Equal(text.ToString(), utf8.ToString());
    }

    [Theory]
    [InlineData(CsvCompressionType.None, false)]
    [InlineData(CsvCompressionType.None, true)]
    [InlineData(CsvCompressionType.GZip, false)]
    [InlineData(CsvCompressionType.GZip, true)]
    public void StreamFactory_CompletesCompressionAndHonorsOwnership(CsvCompressionType compression, bool leaveOpen)
    {
        using var output = new FlushTrackingStream();
        var options = new CsvSaveOptions { CompressionType = compression, NewLine = "\n" };
        using (var writer = CsvRowWriter.CreateStream(output, options, leaveOpen, bufferSize: 16))
            writer.WriteUtf8Row(new[] { "Name" }, Encode(new[] { "🚀" }));
        Assert.Equal(leaveOpen, output.CanWrite);
        Assert.True(output.WasFlushed);
        byte[] bytes = output.ToArray();
        if (compression == CsvCompressionType.GZip)
        {
            using var gzip = new GZipStream(new MemoryStream(bytes), CompressionMode.Decompress);
            using var decoded = new MemoryStream();
            gzip.CopyTo(decoded);
            bytes = decoded.ToArray();
        }
        Assert.Equal("Name\n🚀\n", Encoding.UTF8.GetString(bytes));
    }

    [Fact]
    public void StreamFactory_WritesAtCurrentPositionAndRejectsClosedWriter()
    {
        using var output = new MemoryStream();
        output.Write(Encoding.UTF8.GetBytes("prefix\n"));
        var writer = CsvRowWriter.CreateStream(output, new CsvSaveOptions { IncludeHeader = false, NewLine = "\n" }, leaveOpen: true);
        writer.WriteUtf8Row(new[] { "A" }, Encode(new[] { "value" }));
        writer.Dispose();
        Assert.Equal("prefix\nvalue\n", Encoding.UTF8.GetString(output.ToArray()));
        Assert.Throws<ObjectDisposedException>(() => writer.WriteUtf8Row(Encode(new[] { "closed" })));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Utf8Rows_FileFactoryRetainsEncodingAndAppendBoundary(bool unicode)
    {
        string path = Path.Combine(Path.GetTempPath(), "csv-utf8-" + Guid.NewGuid().ToString("N") + ".csv");
        var options = new CsvSaveOptions { NewLine = "\n", Encoding = unicode ? Encoding.Unicode : new UTF8Encoding(false) };
        try
        {
            using (var writer = CsvRowWriter.CreateFile(path, options, bufferSize: 16))
                writer.WriteUtf8Row(new[] { "Name" }, Encode(new[] { "Zażółć 🚀" }));
            byte[] expected = WriteText(new[] { "Name" }, new[] { new string?[] { "Zażółć 🚀" } }, options);
            Assert.Equal(expected, File.ReadAllBytes(path));
            // Exercise the append separator with a pre-existing file that has no newline.
            File.WriteAllText(path, "Name\nfirst", options.Encoding!);
            options.IncludeHeader = false;
            using (var writer = CsvRowWriter.CreateFile(path, options, append: true, bufferSize: 16))
                writer.WriteUtf8Row(new[] { "Name" }, Encode(new[] { "second 🚀" }));
            byte[] preamble = options.Encoding!.GetPreamble();
            byte[] appended = preamble.Concat(options.Encoding.GetBytes("Name\nfirst\nsecond 🚀\n")).ToArray();
            Assert.Equal(appended, File.ReadAllBytes(path));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    private sealed class FlushTrackingStream : MemoryStream
    {
        internal bool WasFlushed { get; private set; }
        public override void Flush() { WasFlushed = true; base.Flush(); }
    }

    private static ReadOnlyMemory<byte>?[] Encode(string?[] values) =>
        values.Select(value => value is null ? (ReadOnlyMemory<byte>?)null : Encoding.UTF8.GetBytes(value)).ToArray();

    private static byte[] DecodeCompression(byte[] bytes, CsvCompressionType compression)
    {
        if (compression is CsvCompressionType.Auto or CsvCompressionType.None) return bytes;
        using var input = new MemoryStream(bytes);
        using Stream decoded = compression switch {
            CsvCompressionType.GZip => new GZipStream(input, CompressionMode.Decompress),
            CsvCompressionType.Deflate => new DeflateStream(input, CompressionMode.Decompress),
            CsvCompressionType.Brotli => new BrotliStream(input, CompressionMode.Decompress),
            CsvCompressionType.ZLib => new ZLibStream(input, CompressionMode.Decompress),
            _ => throw new ArgumentOutOfRangeException(nameof(compression))
        };
        using var output = new MemoryStream();
        decoded.CopyTo(output);
        return output.ToArray();
    }

    private static byte[] WriteText(string[] columns, string?[][] rows, CsvSaveOptions options)
    {
        using var output = new MemoryStream();
        Stream destination = options.CompressionType switch {
            CsvCompressionType.GZip => new GZipStream(output, options.CompressionLevel, true),
            CsvCompressionType.Deflate => new DeflateStream(output, options.CompressionLevel, true),
            CsvCompressionType.Brotli => new BrotliStream(output, options.CompressionLevel, true),
            CsvCompressionType.ZLib => new ZLibStream(output, options.CompressionLevel, true),
            _ => output
        };
        using (var text = new StreamWriter(destination, options.Encoding ?? new UTF8Encoding(false), 16, leaveOpen: destination == output))
        using (var writer = new CsvRowWriter(text, options))
        {
            foreach (string?[] row in rows) writer.WriteTextRow(columns, row);
        }
        return output.ToArray();
    }
}
#endif
