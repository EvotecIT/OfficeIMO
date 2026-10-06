using System;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvSynchronousStreamSaveTests
{
    [Theory]
    [InlineData(CsvCompressionType.None)]
    [InlineData(CsvCompressionType.GZip)]
    [InlineData(CsvCompressionType.Deflate)]
#if NET6_0_OR_GREATER
    [InlineData(CsvCompressionType.Brotli)]
    [InlineData(CsvCompressionType.ZLib)]
#endif
    public void Save_Replaces_Seekable_Stream_And_Rewinds(CsvCompressionType compression)
    {
        var document = new CsvDocument().WithHeader("Name", "Notes")
            .AddRow("Alpha", "Łódź, line1\nline2");
        var options = new CsvSaveOptions { NewLine = "\n", CompressionType = compression };
        using var output = new MemoryStream();
        byte[] previous = Encoding.UTF8.GetBytes(new string('x', 2048));
        output.Write(previous, 0, previous.Length);
        output.Position = 5;

        document.Save(output, options);

        Assert.True(output.CanWrite);
        Assert.Equal(0, output.Position);
        Assert.Equal(document.ToBytes(options), output.ToArray());
        Assert.Equal("Name,Notes\nAlpha,\"Łódź, line1\nline2\"\n",
            CsvDocument.Load(output, new CsvLoadOptions { CompressionType = compression }).ToString(options));
    }

    [Theory]
    [InlineData(CsvCompressionType.None, false)]
    [InlineData(CsvCompressionType.None, true)]
    [InlineData(CsvCompressionType.GZip, false)]
    [InlineData(CsvCompressionType.GZip, true)]
    public void Save_Formatting_Failure_Preserves_Contents_And_Position(CsvCompressionType compression, bool useUtf16)
    {
        var document = new CsvDocument().WithHeader("Value")
            .AddRow(new string('x', 600_000))
            .AddRow(new ThrowingValue());
        var options = new CsvSaveOptions {
            CompressionType = compression,
            Encoding = useUtf16 ? Encoding.Unicode : new UTF8Encoding(false)
        };
        byte[] previous = Encoding.UTF8.GetBytes("existing complete artifact");
        using var output = new MemoryStream();
        output.Write(previous, 0, previous.Length);
        output.Position = 5;

        var exception = Assert.Throws<InvalidOperationException>(() => document.Save(output, options));

        Assert.Equal("Value formatting failed.", exception.Message);
        Assert.Equal(previous, output.ToArray());
        Assert.Equal(5, output.Position);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public void Save_Rejects_A_Fixed_Length_Mapped_View_Without_Changing_The_File()
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.MappedSave." + Guid.NewGuid().ToString("N"));
        byte[] previous = Encoding.UTF8.GetBytes(new string('x', 2048));
        try
        {
            File.WriteAllBytes(path, previous);
            using (var mapping = MemoryMappedFile.CreateFromFile(path, FileMode.Open))
            using (var output = mapping.CreateViewStream(0, previous.Length, MemoryMappedFileAccess.ReadWrite))
            {
                output.Position = 5;
                Assert.Throws<NotSupportedException>(() =>
                    new CsvDocument().WithHeader("Name").AddRow("Alpha").Save(output));
                Assert.True(output.CanWrite);
                output.Flush();
            }
            Assert.Equal(previous, File.ReadAllBytes(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Save_Appends_To_Forward_Only_Stream_And_Leaves_It_Open()
    {
        using var output = new ForwardOnlyOutput();
        byte[] prefix = Encoding.UTF8.GetBytes("prefix\n");
        output.Write(prefix, 0, prefix.Length);

        new CsvDocument().WithHeader("Name").AddRow("Alpha")
            .Save(output, new CsvSaveOptions { NewLine = "\n" });

        Assert.True(output.CanWrite);
        Assert.Equal("prefix\nName\nAlpha\n", Encoding.UTF8.GetString(output.ToArray()));
    }

    private sealed class ThrowingValue
    {
        public override string ToString() => throw new InvalidOperationException("Value formatting failed.");
    }

    private sealed class ForwardOnlyOutput : MemoryStream
    {
        public override bool CanSeek => false;
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override long Seek(long offset, SeekOrigin loc) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
