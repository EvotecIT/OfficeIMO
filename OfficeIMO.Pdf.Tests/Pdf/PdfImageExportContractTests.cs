using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfImageExportContractTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void LoadedPdfPageUsesTheSharedFiveFormatContract(OfficeImageExportFormat format) {
        PdfReadDocument document = LoadTwoPageDocument();

        OfficeImageExportResult result = document.Pages[0].ExportImage(format);

        Assert.Equal(format, result.Format);
        Assert.Equal(format.GetMimeType(), OfficeImageReader.Identify(result.Bytes).MimeType);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
        Assert.NotEmpty(result.Bytes);
    }

    [Fact]
    public void PdfBatchBuilderPreservesSelectionAndSharedRasterSafety() {
        PdfReadDocument document = LoadTwoPageDocument();

        IReadOnlyList<OfficeImageExportResult> results = document
            .ToImages(new PdfImageExportOptions {
                Scale = 10D,
                MaximumRasterPixels = 20_000L
            })
            .Pages("2,1")
            .AsWebp()
            .Export();

        Assert.Equal(new[] { "Page 2", "Page 1" }, results.Select(result => result.Name));
        Assert.All(results, result => {
            Assert.Equal(OfficeImageExportFormat.Webp, result.Format);
            Assert.True((long)result.Width * result.Height <= 20_000L);
            Assert.Contains(result.Diagnostics, diagnostic =>
                diagnostic.Code == OfficeImageExportDiagnosticCodes.RasterScaleReduced);
        });
    }

    [Fact]
    public void PdfPageBuilderSupportsDpiThumbnailAndConfiguredEncoding() {
        PdfReadPage page = LoadTwoPageDocument().Pages[0];

        OfficeImageExportResult result = page
            .ToImage()
            .AtDpi(288D)
            .AsThumbnail(96)
            .WithRasterEncoding(encoding => encoding.Jpeg.Quality = 70)
            .AsJpeg()
            .Export();

        Assert.Equal(OfficeImageExportFormat.Jpeg, result.Format);
        Assert.True(Math.Max(result.Width, result.Height) <= 96);
        Assert.InRange(result.DpiX, 1D, 287D);
        Assert.Equal(result.DpiX, result.DpiY);
        Assert.Equal("image/jpeg", OfficeImageReader.Identify(result.Bytes).MimeType);
    }

    [Fact]
    public void PdfSvgUsesTheSharedConfiguredBackground() {
        PdfReadPage page = LoadTwoPageDocument().Pages[0];

        OfficeImageExportResult result = page
            .ToImage()
            .WithBackground("#123456")
            .AsSvg()
            .Export();
        string svg = System.Text.Encoding.UTF8.GetString(result.Bytes);

        Assert.Contains("fill=\"#123456\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void ConversionResultIsTheSinglePagedImageAdapterAndPreservesDiagnostics() {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Markdown.Pdf",
            "MARKDOWN_APPROXIMATION",
            "table 1",
            "A source feature was approximated."));
        var conversion = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Projected source")),
            report);

        OfficeImageExportResult result = Assert.Single(conversion
            .ToImages()
            .AsPng()
            .Export());

        OfficeImageExportDiagnostic diagnostic = Assert.Single(
            result.Diagnostics,
            item => item.Code == "MARKDOWN_APPROXIMATION");
        Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
        Assert.Equal("table 1", diagnostic.Source);
    }

    private static PdfReadDocument LoadTwoPageDocument() =>
        PdfReadDocument.Load(PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page two"))
            .ToBytes());
}
