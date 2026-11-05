using System;
using System.IO;
using System.Text;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvDataReaderInterningTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReaderSharesNormalizedFieldsAcrossQuotedFallbackWithinOneRead(bool trim)
    {
        const string csv = "Name,Other\nshared-value,Name\n\"shared-value\",\"Łódź\"\nshared-value,Łódź\n";
        var options = new CsvReaderOptions {
            DetectDelimiter = false, InternStrings = true, TrimWhitespace = trim
        };
        string retained;
        using (var input = new MemoryStream(Encoding.UTF8.GetBytes(csv)))
        using (var reader = CsvDocument.OpenReader(input, options))
        {
            Assert.True(reader.Read());
            retained = reader.GetString(0);
            Assert.Equal("shared-value", retained);
            Assert.Same(reader.GetName(0), reader.GetValue(1));
            Assert.True(reader.Read());
            Assert.Same(retained, reader.GetValue(0));
            string unicode = reader.GetString(1);
            Assert.Equal("Łódź", unicode);
            var values = new object[2];
            Assert.Equal(2, reader.GetValues(values));
            Assert.Same(retained, values[0]);
            Assert.Same(unicode, values[1]);
            Assert.True(reader.Read());
            Assert.Same(retained, reader.GetString(0));
            Assert.Same(unicode, reader.GetValue(1));
            Assert.False(reader.Read());
        }
        using (var input = new MemoryStream(Encoding.UTF8.GetBytes(csv)))
        using (var reader = CsvDocument.OpenReader(input, options))
        {
            Assert.True(reader.Read());
            Assert.Equal(retained, reader.GetString(0));
            Assert.NotSame(retained, reader.GetString(0));
        }
        Assert.Equal("shared-value", retained);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InternedReaderPreservesProgressFailureAndCancellationOrdering(bool cancel)
    {
        using var cancellation = new System.Threading.CancellationTokenSource();
        var failure = new InvalidOperationException("progress failed");
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("V\none\ntwo\n"));
        using var reader = CsvDocument.OpenReader(input, new CsvReaderOptions {
            DetectDelimiter = false,
            InternStrings = true,
            CancellationToken = cancellation.Token,
            ProgressReportInterval = cancel ? 2 : 1,
            ProgressCallback = _ => { if (cancel) cancellation.Cancel(); else throw failure; }
        });
        if (cancel)
        {
            Assert.True(reader.Read());
            Assert.Equal("one", reader.GetString(0));
            Assert.ThrowsAny<OperationCanceledException>(() => reader.Read());
        }
        else
        {
            Assert.Same(failure, Assert.Throws<InvalidOperationException>(() => reader.Read()));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReaderSkipsRecordsEmptiedByItsDecoder(bool intern)
    {
        var encoding = (Encoding)new UTF8Encoding(false).Clone();
        encoding.DecoderFallback = new DecoderReplacementFallback("");
        byte[] csv = { (byte)'V', 10, 0xFF, 10, (byte)'k', (byte)'e', (byte)'p', (byte)'t', 10 };
        using var input = new MemoryStream(csv);
        using var reader = CsvDocument.OpenReader(input, new CsvReaderOptions {
            Encoding = encoding, DetectDelimiter = false, InternStrings = intern
        });
        Assert.True(reader.Read());
        Assert.Equal("kept", reader.GetString(0));
        Assert.False(reader.Read());
    }

}
