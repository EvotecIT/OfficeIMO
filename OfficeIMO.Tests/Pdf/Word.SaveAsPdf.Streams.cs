using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf_ToStream_NotWritable_Throws() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfStreamReadOnly.docx");

        using var document = WordDocument.Create(docPath);
        document.AddParagraph("Hello World");
        document.Save();

        using var stream = new MemoryStream(new byte[1], 0, 1, writable: false);
        Assert.Throws<ArgumentException>(() => document.SaveAsPdf(stream));
    }

    [Fact]
    public async Task Test_WordDocument_SaveAsPdfAsync_ToStream_NotWritable_Throws() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfStreamReadOnlyAsync.docx");

        using var document = WordDocument.Create(docPath);
        document.AddParagraph("Hello World");
        document.Save();

        using var stream = new MemoryStream(new byte[1], 0, 1, writable: false);
        await Assert.ThrowsAsync<ArgumentException>(() => document.SaveAsPdfAsync(stream, cancellationToken: CancellationToken.None));
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_ToStream_Rewinds() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfStreamRewind.docx");

        using var document = WordDocument.Create(docPath);
        document.AddParagraph("Hello World");
        document.Save();

        using var stream = new MemoryStream();
        document.SaveAsPdf(stream);
        Assert.Equal(0, stream.Position);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public async Task Test_WordDocument_SaveAsPdfAsync_ToStream_Rewinds() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfStreamRewindAsync.docx");

        using var document = WordDocument.Create(docPath);
        document.AddParagraph("Hello World");
        document.Save();

        using var stream = new MemoryStream();
        await document.SaveAsPdfAsync(stream, cancellationToken: CancellationToken.None);
        Assert.Equal(0, stream.Position);
        Assert.True(stream.Length > 0);
    }
}