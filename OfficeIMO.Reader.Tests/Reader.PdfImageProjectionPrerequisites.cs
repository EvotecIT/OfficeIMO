using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Image;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPdfImageProjectionPrerequisitesTests {
    [Fact]
    public void PdfTextAndMetadataAvoidsImageExtractionAndRetainsDecodingOutcomesAndSourceFormCounts() {
        byte[] bytes = PdfInspectionProjectionFixture.Build(32);
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(new ReaderPdfOptions {
            LoadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxContentOperands = 20 } }
        }).Build();
        var options = new ReaderOptions { Projection = ReaderProjection.TextAndMetadata, Inspection = new ReaderInspectionOptions() };
        OfficeDocumentReadResult result = reader.ReadDocument(bytes, "lean.pdf", options);
        Assert.Null(result.Markdown); Assert.Empty(result.Assets); Assert.Empty(result.Visuals); Assert.Empty(result.Forms); Assert.Empty(result.OcrCandidates);
        Assert.Contains(result.Chunks, chunk => chunk.Text!.Contains("Native page text"));
        Assert.Contains(result.Chunks, chunk => chunk.Text!.Contains("Nested form text"));
        Assert.All(result.Chunks, chunk => Assert.Null(chunk.Markdown));
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.text-decoding");
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.page-outcome" && entry.Value == "Read");
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.form" && entry.Name == "SourceWidgetCount" && entry.Value == "1");
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.source.form-widget" && entry.Value == "Entry");
        Assert.Throws<PdfReadLimitException>(() => reader.ReadDocument(bytes, "full.pdf", new ReaderOptions { Inspection = new ReaderInspectionOptions() }));
        OfficeDocumentReadResult full = new OfficeDocumentReaderBuilder().AddPdfHandler().Build().ReadDocument(bytes, "full.pdf");
        Assert.NotNull(full.Markdown); Assert.NotEmpty(full.Assets); Assert.NotEmpty(full.Forms);
        OfficeDocumentReadResult documentChunks = new OfficeDocumentReaderBuilder().AddPdfHandler(new ReaderPdfOptions { ChunkByPage = false })
            .Build().ReadDocument(bytes, "document.pdf", new ReaderOptions { Projection = ReaderProjection.TextAndMetadata });
        Assert.All(documentChunks.Chunks, chunk => Assert.Null(chunk.Location.Page));
        Assert.Contains(documentChunks.Chunks, chunk => chunk.Text!.Contains("Nested form text"));
    }

    [Fact]
    public void ImageTextAndMetadataAvoidsAssetAndOcrProjectionBeforeAssetBudgetAdmission() {
        byte[] bytes = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue), OfficeImageExportFormat.Png);
        var reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();
        var options = new ReaderOptions { Projection = ReaderProjection.TextAndMetadata, Inspection = new ReaderInspectionOptions(),
            ResourceLimits = new ReaderResourceLimits { MaxAssets = 0 } };
        OfficeDocumentReadResult result = reader.ReadDocument(bytes, "photo.png", options);
        Assert.Null(result.Markdown); Assert.NotEmpty(result.Chunks); Assert.Empty(result.Assets); Assert.Empty(result.Visuals); Assert.Empty(result.OcrCandidates);
        Assert.All(result.Chunks, chunk => Assert.Null(chunk.Markdown));
        Assert.Contains(result.Metadata, entry => entry.Category == "image.resolution");
        Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(bytes, "photo.png",
            new ReaderOptions { ResourceLimits = new ReaderResourceLimits { MaxAssets = 0 } }));
        Assert.NotEmpty(reader.ReadDocument(bytes, "photo.png").Assets);
    }
}
