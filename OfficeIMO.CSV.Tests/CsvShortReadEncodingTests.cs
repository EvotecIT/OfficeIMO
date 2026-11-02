using System;
using System.IO;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvShortReadEncodingTests
{
    [Theory]
    [InlineData("utf8", CsvCompressionType.None)]
    [InlineData("utf16", CsvCompressionType.None)]
    [InlineData("utf32", CsvCompressionType.None)]
    [InlineData("utf8", CsvCompressionType.GZip)]
    [InlineData("utf16", CsvCompressionType.GZip)]
    [InlineData("utf32", CsvCompressionType.Deflate)]
    public void Load_Detects_Bom_When_Source_Returns_One_Byte(string encodingName, CsvCompressionType compression)
    {
        Encoding encoding = encodingName switch
        {
            "utf16" => new UnicodeEncoding(false, true),
            "utf32" => new UTF32Encoding(false, true),
            _ => new UTF8Encoding(true)
        };
        using var bytes = new MemoryStream();
        CsvDocument.Parse("Name,Notes\nAlpha,Zażółć\n").Save(bytes, new CsvSaveOptions { Encoding = encoding, CompressionType = compression });
        using var source = new ShortReadStream(bytes.ToArray());
        var document = CsvDocument.Load(source, new CsvLoadOptions { CompressionType = compression });
        Assert.Equal($"Name,Notes{Environment.NewLine}Alpha,Zażółć{Environment.NewLine}", document.ToString());
        Assert.True(source.CanRead);
    }

    private sealed class ShortReadStream : Stream
    {
        private readonly MemoryStream _inner;
        internal ShortReadStream(byte[] bytes) => _inner = new MemoryStream(bytes);
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, Math.Min(count, 1));
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}
