using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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
    public void PdfBatchBuilderUsesTheSharedOutputCountBudget() {
        PdfReadDocument document = LoadTwoPageDocument();

        OfficeImageExportBatchLimitException exception =
            Assert.Throws<OfficeImageExportBatchLimitException>(() =>
                document
                    .ToImages()
                    .WithMaximumPages(1)
                    .AsPng()
                    .Export());

        Assert.Equal(nameof(OfficeImageExportOptions.MaximumOutputCount), exception.LimitName);
        Assert.Equal(2, exception.Actual);
        Assert.Equal(1, exception.Maximum);
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
    public void FluentDocumentReaderUsesTheSharedImageExportContract() {
        PdfDocument document = PdfDocument.Open(PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Fluent image export"))
            .ToBytes());

        OfficeImageExportResult result = Assert.Single(document.Read.ExportImages(
            OfficeImageExportFormat.Svg));

        Assert.Equal(OfficeImageExportFormat.Svg, result.Format);
        Assert.Equal(result.Width, OfficeImageReader.Identify(result.Bytes).Width);
        Assert.Equal(result.Height, OfficeImageReader.Identify(result.Bytes).Height);
    }

    [Fact]
    public void AuthoredDocumentExportsImagesWithoutAReadFacade() {
        PdfDocument document = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Authored page"));

        OfficeImageExportResult direct = Assert.Single(document.ExportImages(OfficeImageExportFormat.Webp));
        OfficeImageExportResult fluent = Assert.Single(document.ToImages().AsPng().Export());

        Assert.Equal(OfficeImageExportFormat.Webp, direct.Format);
        Assert.Equal("image/webp", OfficeImageReader.Identify(direct.Bytes).MimeType);
        Assert.Equal(OfficeImageExportFormat.Png, fluent.Format);
        Assert.Equal("image/png", OfficeImageReader.Identify(fluent.Bytes).MimeType);
    }

    [Fact]
    public async Task AuthoredDocumentExportDefersSnapshotAndHonorsPreCanceledToken() {
        bool materialized = false;
        PdfDocument document = PdfDocument.Create()
            .Deferred(_ => {
                materialized = true;
                return item => item.Paragraph(paragraph => paragraph.Text("Deferred page"));
            });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        PdfDocumentImageExportBuilder builder = document.ToImages();

        Assert.False(materialized);
        Assert.ThrowsAny<OperationCanceledException>(() =>
            document.ExportImages(OfficeImageExportFormat.Png, cancellationToken: cancellation.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => builder.ExportAsync(cancellation.Token));
        Assert.False(materialized);
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

    [Fact]
    public void PdfImageExport_StrictOmissionPolicyRejectsUnsupportedOperators() {
        PdfReadDocument document = PdfReadDocument.Open(
            BuildSinglePagePdf("1 2 FuturePaint"));
        var options = new PdfImageExportOptions {
            Policy = new OfficeImageExportPolicy {
                RequireNoOmissions = true
            }
        };

        OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(
            () => document.Pages[0].ExportImage(OfficeImageExportFormat.Png, options));

        Assert.Contains(
            exception.Diagnostics,
            diagnostic =>
                diagnostic.Code == "render.operator.unsupported" &&
                diagnostic.LossKind == OfficeImageExportLossKind.Omission);
    }

    private static PdfReadDocument LoadTwoPageDocument() =>
        PdfReadDocument.Open(PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page two"))
            .ToBytes());

    private static byte[] BuildSinglePagePdf(string content) {
        int length = Encoding.ASCII.GetByteCount(content);
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>", "stream",
            content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
    }
}
