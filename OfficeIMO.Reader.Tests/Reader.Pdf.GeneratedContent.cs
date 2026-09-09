using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPdfGeneratedContentTests {
    [Fact]
    public void DetectedTableKeepsItsGeometryAndCellsWithoutInventingSourceText() {
        byte[] pdf = PdfDocument.Create(builder => builder.Content(content => content.Table(new[] {
            new[] { "Code", "Name", "Qty" }, new[] { "A-100", "Alpha", "2" }, new[] { "B-200", "Beta", "14" }
        }, style: new PdfTableStyle { HeaderRowCount = 1 }))).ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);
        var document = PdfReaderAdapter.ReadDocument(stream, "table.pdf");
        var block = Assert.Single(document.Pages.SelectMany(page => page.Blocks), block => block.Kind == "table");
        Assert.Empty(block.Text);
        Assert.False(string.IsNullOrWhiteSpace(block.Id));
        Assert.Equal(1, block.Location!.Page);
        Assert.NotNull(block.Region);
        Assert.True(block.Region!.Width > 0 && block.Region.Height > 0);
        Assert.Contains(document.Pages.SelectMany(page => page.Tables).SelectMany(table => table.Rows), row => row.Contains("A-100") && row.Contains("Alpha"));
    }

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
