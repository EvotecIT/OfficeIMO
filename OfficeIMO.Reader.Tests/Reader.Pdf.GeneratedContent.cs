using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPdfGeneratedContentTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void NonTextPdfChunksAreMarkedAsGeneratedContent(bool perPage, bool imageOnly) {
        byte[] png = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(20, 30, OfficeColor.White), OfficeImageExportFormat.Png);
        byte[] pdf = PdfDocument.Create(document => document.Page(page => {
            page.Size(200, 300);
            if (imageOnly) page.Canvas(canvas => canvas.Image(png, 0, 0, 200, 300));
        })).ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);
        var chunks = PdfReaderAdapter.Read(stream, "empty.pdf", pdfOptions: new ReaderPdfOptions { ChunkByPage = perPage }).ToArray();
        Assert.NotEmpty(chunks);
        Assert.All(chunks, chunk => Assert.Contains(chunk.Location.SourceBlockKind, new[] { "warning", "visual" }));
        if (imageOnly) Assert.Contains(chunks, chunk => chunk.Visuals?.Count > 0);
    }
}
