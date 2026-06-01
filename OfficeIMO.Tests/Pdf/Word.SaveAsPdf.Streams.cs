using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using UglyToad.PdfPig;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

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

    [Fact]
    public void Test_WordDocument_SaveAsPdf_OfficeIMOEngine_ToStream_Rewinds() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfNativeStreamRewind.docx");

        using var document = WordDocument.Create(docPath);
        document.AddParagraph("Hello native stream");
        document.Save();

        using var stream = new MemoryStream();
        document.SaveAsPdf(stream, new PdfSaveOptions());
        Assert.Equal(0, stream.Position);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public async Task Test_WordDocument_SaveAsPdfAsync_OfficeIMOEngine_ToStream_Rewinds() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfNativeStreamRewindAsync.docx");

        using var document = WordDocument.Create(docPath);
        document.AddParagraph("Hello native async stream");
        document.Save();

        using var stream = new MemoryStream();
        await document.SaveAsPdfAsync(stream, new PdfSaveOptions(), CancellationToken.None);
        Assert.Equal(0, stream.Position);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public async Task Test_WordDocument_SaveAsPdfAsync_OfficeIMOEngine_ToBytes_UsesNativeEngine() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfNativeAsyncBytes.docx");

        using var document = WordDocument.Create(docPath);
        document.AddParagraph("Hello native async bytes");
        document.Save();

        byte[] bytes = await document.SaveAsPdfAsync(new PdfSaveOptions {
            IncludePageNumbers = false,
            PageSize = new PdfCore.PageSize(240, 320),
            Margins = PdfCore.PageMargins.Uniform(36)
        }, CancellationToken.None);

        Assert.True(bytes.Length > 0);
        PdfCore.PdfPageInfo pageInfo = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).Pages);
        Assert.Equal(240, pageInfo.Width, 1);
        Assert.Equal(320, pageInfo.Height, 1);

        using PdfDocument pdf = PdfDocument.Open(bytes);
        Assert.Contains("Hello native async bytes", pdf.GetPage(1).Text);
    }
}
