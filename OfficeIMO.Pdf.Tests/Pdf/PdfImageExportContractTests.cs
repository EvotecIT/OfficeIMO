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
    public void LoadedPdfPageSupportsDirectCancellationAndSharedFitWithinBounds() {
        PdfReadPage page = LoadTwoPageDocument().Pages[0];
        using var cancellation = new CancellationTokenSource();

        OfficeImageExportResult result = page.ExportImage(
            OfficeImageExportFormat.Svg,
            new PdfImageExportOptions {
                Scale = 4D,
                MaximumOutputWidth = 160,
                MaximumOutputHeight = 160
            },
            cancellation.Token);

        Assert.True(result.Width <= 160);
        Assert.True(result.Height <= 160);
    }

    [Fact]
    public void PdfBatchOptionsPreservesSelectionAndSharedRasterSafety() {
        PdfReadDocument document = LoadTwoPageDocument();

        IReadOnlyList<OfficeImageExportResult> results = document
            .CreateImageExportOperation(new PdfImageExportOptions {
                Scale = 10D, MaximumRasterPixels = 20_000L,
                PageSelector = PdfPageSelector.Parse("2,1"), Format = OfficeImageExportFormat.Webp
            }).Export();

        Assert.Equal(new[] { "Page 2", "Page 1" }, results.Select(result => result.Name));
        Assert.All(results, result => {
            Assert.Equal(OfficeImageExportFormat.Webp, result.Format);
            Assert.True((long)result.Width * result.Height <= 20_000L);
            Assert.Contains(result.Diagnostics, diagnostic =>
                diagnostic.Code == OfficeImageExportDiagnosticCodes.RasterScaleReduced);
        });
    }

    [Fact]
    public void PdfBatchOptionsUsesTheSharedOutputCountBudget() {
        PdfReadDocument document = LoadTwoPageDocument();

        OfficeImageExportBatchLimitException exception =
            Assert.Throws<OfficeImageExportBatchLimitException>(() =>
                document
                    .CreateImageExportOperation(new PdfImageExportOptions { MaximumOutputCount = 1 }).Export());

        Assert.Equal(nameof(OfficeImageExportOptions.MaximumOutputCount), exception.LimitName);
        Assert.Equal(2, exception.Actual);
        Assert.Equal(1, exception.Maximum);
    }

    [Fact]
    public void PdfSvgExportEnforcesTheEncodedByteBudgetDuringSerialization() {
        PdfReadPage page = LoadTwoPageDocument().Pages[0];

        OfficeImageExportBatchLimitException exception =
            Assert.Throws<OfficeImageExportBatchLimitException>(() =>
                page.ExportImage(
                    OfficeImageExportFormat.Svg,
                    new PdfImageExportOptions { MaximumTotalEncodedBytes = 32L }));

        Assert.Equal(nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes), exception.LimitName);
        Assert.Equal(32L, exception.Maximum);
    }

    [Fact]
    public void PdfSvgBatchReportsTheConfiguredAggregateByteBudget() {
        PdfReadDocument document = LoadTwoPageDocument();
        long onePageBytes = document.Pages[0]
            .ExportImage(OfficeImageExportFormat.Svg)
            .Bytes.LongLength;
        long maximumBytes = onePageBytes + 16L;

        OfficeImageExportBatchLimitException exception =
            Assert.Throws<OfficeImageExportBatchLimitException>(() =>
                document
                    .CreateImageExportOperation(new PdfImageExportOptions {
                        MaximumTotalEncodedBytes = maximumBytes, Format = OfficeImageExportFormat.Svg }).Export());

        Assert.Equal(nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes), exception.LimitName);
        Assert.True(exception.Actual > maximumBytes);
        Assert.Equal(maximumBytes, exception.Maximum);
    }

    [Fact]
    public void PdfPageOptionsSupportDpiThumbnailAndConfiguredEncoding() {
        PdfReadPage page = LoadTwoPageDocument().Pages[0];

        OfficeImageExportResult result = page
            .ExportImage(OfficeImageExportFormat.Jpeg, new PdfImageExportOptions {
                TargetDpi = 288D, ThumbnailMaxDimension = 96,
                RasterEncoding = new OfficeRasterEncodingOptions { Jpeg = new OfficeJpegEncodeOptions { Quality = 70 } }
            });

        Assert.Equal(OfficeImageExportFormat.Jpeg, result.Format);
        Assert.True(Math.Max(result.Width, result.Height) <= 96);
        Assert.InRange(result.DpiX, 1D, 287D);
        Assert.Equal(result.DpiX, result.DpiY);
        Assert.Equal("image/jpeg", OfficeImageReader.Identify(result.Bytes).MimeType);
    }

    [Fact]
    public void PdfRasterFitPreservesDensityRelativeToThePreFitScale() {
        PdfReadPage page = LoadTwoPageDocument().Pages[0];
        OfficeImageExportResult baseline = page.ExportImage(OfficeImageExportFormat.Png);

        OfficeImageExportResult fitted = page
            .ExportImage(OfficeImageExportFormat.Png, new PdfImageExportOptions {
                Scale = 2D, MaximumOutputWidth = baseline.Width, MaximumOutputHeight = baseline.Height
            });

        Assert.Equal(baseline.Width, fitted.Width);
        Assert.Equal(baseline.Height, fitted.Height);
        Assert.InRange(fitted.DpiX, baseline.DpiX / 2D - 0.05D, baseline.DpiX / 2D + 0.05D);
        Assert.InRange(fitted.DpiY, baseline.DpiY / 2D - 0.05D, baseline.DpiY / 2D + 0.05D);
    }

    [Fact]
    public void PdfSvgUsesTheSharedConfiguredBackground() {
        PdfReadPage page = LoadTwoPageDocument().Pages[0];

        OfficeImageExportResult result = page
            .ExportImage(OfficeImageExportFormat.Svg, new PdfImageExportOptions { BackgroundColor = OfficeIMO.OfficeColor.Parse("#123456") });
        string svg = System.Text.Encoding.UTF8.GetString(result.Bytes);

        Assert.Contains("fill=\"#123456\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void LoadedPdfPageChecksTheDeadlineAfterDirectSvgExport() {
        PdfReadPage page = LoadTwoPageDocument().Pages[0];
        var timeout = TimeSpan.FromTicks(1);

        OfficeImageExportTimeoutException exception = Assert.Throws<OfficeImageExportTimeoutException>(() =>
            page.ExportImage(
                OfficeImageExportFormat.Svg,
                new PdfImageExportOptions { RenderTimeout = timeout }));

        Assert.Equal(timeout, exception.Timeout);
    }

    [Fact]
    public void LoadedPdfBatchChecksTheDeadlineAfterOrderedExport() {
        PdfReadDocument document = LoadTwoPageDocument();
        var timeout = TimeSpan.FromTicks(1);

        OfficeImageExportTimeoutException exception = Assert.Throws<OfficeImageExportTimeoutException>(() =>
            document.ExportImages(
                OfficeImageExportFormat.Svg,
                new PdfImageExportOptions { RenderTimeout = timeout }));

        Assert.Equal(timeout, exception.Timeout);
    }

    [Fact]
    public void DocumentReaderUsesTheSharedImageExportContract() {
        PdfDocument document = PdfDocument.Load(PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Fluent image export"))
            .ToBytes());

        OfficeImageExportResult result = Assert.Single(document.Reader.ExportImages(
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
        byte[] png = Assert.Single(document.ExportImages()).Bytes;

        Assert.Equal(OfficeImageExportFormat.Webp, direct.Format);
        Assert.Equal("image/webp", OfficeImageReader.Identify(direct.Bytes).MimeType);
        Assert.Equal("image/png", OfficeImageReader.Identify(png).MimeType);
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

        var options = new PdfImageExportOptions();

        Assert.False(materialized);
        Assert.ThrowsAny<OperationCanceledException>(() =>
            document.ExportImages(OfficeImageExportFormat.Png, cancellationToken: cancellation.Token));
        Assert.ThrowsAny<OperationCanceledException>(() =>
            document.Reader.ExportImages(OfficeImageExportFormat.Png, cancellationToken: cancellation.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.ExportImagesAsync(options, cancellation.Token));
        Assert.False(materialized);
    }

    [Fact]
    public async Task AuthoredDocumentExportPropagatesCancellationThroughDeferredSnapshot() {
        using var cancellation = new CancellationTokenSource();
        bool materializedAfterCancellation = false;
        PdfDocument document = CreateCancelingDeferredDocument(cancellation, () => materializedAfterCancellation = true);

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            document.ExportImagesAsync(new PdfImageExportOptions { Format = OfficeImageExportFormat.Svg }, cancellation.Token));

        Assert.False(materializedAfterCancellation);
    }

    [Fact]
    public async Task ConversionExportPropagatesCancellationThroughDeferredSnapshot() {
        using var cancellation = new CancellationTokenSource();
        bool materializedAfterCancellation = false;
        var conversion = new PdfDocumentConversionResult(
            CreateCancelingDeferredDocument(cancellation, () => materializedAfterCancellation = true),
            new PdfConversionReport());

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            conversion.CreateImageExportOperation(new PdfImageExportOptions { Format = OfficeImageExportFormat.Svg }).ExportAsync(cancellation.Token));

        Assert.False(materializedAfterCancellation);
    }

    [Fact]
    public void AuthoredDocumentDirectExportIncludesSnapshotMaterializationInDeadline() {
        PdfDocument document = PdfDocument.Create()
            .Deferred(_ => item => item.Paragraph(paragraph => paragraph.Text("Deferred page")));
        var timeout = TimeSpan.FromTicks(1);

        OfficeImageExportTimeoutException exception = Assert.Throws<OfficeImageExportTimeoutException>(() =>
            document.ExportImages(
                OfficeImageExportFormat.Svg,
                new PdfImageExportOptions { RenderTimeout = timeout }));

        Assert.Equal(timeout, exception.Timeout);
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
            .CreateImageExportOperation().Export());

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
                diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public async Task PublicImageOptionsPreservePageOrderAndFreezeAsyncSettings() {
        using PdfDocument document = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak().Paragraph(paragraph => paragraph.Text("Page two"));
        var options = new PdfImageExportOptions {
            Format = OfficeImageExportFormat.Svg, PageSelector = PdfPageSelector.Parse("last,1"),
            MaximumOutputWidth = 160, MaximumOutputHeight = 160
        };
        Task<OfficeIMO.OfficeConversionResult<IReadOnlyList<OfficeImageExportResult>>> pending = document.ExportImagesAsync(options);
        options.Format = OfficeImageExportFormat.Jpeg;
        options.PageSelector = PdfPageSelector.Parse("1");
        var exported = await pending;
        Assert.Equal(2, exported.Value.Count);
        Assert.Equal(new[] { "Page 2", "Page 1" }, exported.Value.Select(result => result.Name));
        Assert.Equal(2, exported.Report.ItemsWritten);
        Assert.Equal("Svg", exported.Report.Target);
        Assert.Equal(document.ExportImages(new PdfImageExportOptions {
            Format = OfficeImageExportFormat.Svg, PageSelector = PdfPageSelector.Parse("2"),
            MaximumOutputWidth = 160, MaximumOutputHeight = 160
        }).Single().Bytes, exported.Value[0].Bytes);
        Assert.All(exported.Value, result => {
            var image = OfficeImageReader.Identify(result.Bytes);
            Assert.Equal(OfficeImageExportFormat.Svg, result.Format);
            Assert.Equal(image.Width, result.Width);
            Assert.Equal(image.Height, result.Height);
            Assert.Equal("image/svg+xml", image.MimeType);
            Assert.InRange(image.Width, 1, 160);
            Assert.InRange(image.Height, 1, 160);
        });
    }

    [Fact]
    public async Task PublicImageSavesReturnReportsAndRejectLossBeforeWriting() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-image-" + Guid.NewGuid().ToString("N"));
        try {
            using PdfDocument document = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Page"));
            var options = new PdfImageExportOptions { Format = OfficeImageExportFormat.Webp };
            OfficeIMO.OfficeConversionReport report = await document.SaveAsImagesAsync(directory, options);
            string path = Assert.Single(Directory.GetFiles(directory));
            Assert.Equal("image/webp", OfficeImageReader.Identify(File.ReadAllBytes(path)).MimeType);
            Assert.Equal(1, report.ItemsWritten);
            Assert.Equal(Path.GetFullPath(directory), report.OutputPath);
            Assert.Throws<IOException>(() => document.SaveAsImages(directory, options));
            using var input = new MemoryStream(BuildSinglePagePdf("1 2 FuturePaint"));
            using PdfDocument lossy = PdfDocument.Load(input);
            var strict = new PdfImageExportOptions { Format = OfficeImageExportFormat.Svg, FailOnLoss = true };
            string strictDirectory = Path.Combine(directory, "strict");
            var failure = Assert.Throws<OfficeIMO.OfficeConversionException>(() => lossy.SaveAsImages(strictDirectory, strict));
            Assert.Contains(Assert.IsType<OfficeIMO.OfficeConversionReport>(failure.Report).Diagnostics, finding => finding.Code == "render.operator.unsupported" &&
                finding.LossKind == OfficeConversionLossKind.Omission);
            Assert.Empty(Directory.Exists(strictDirectory) ? Directory.GetFiles(strictDirectory) : Array.Empty<string>());
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    private static PdfReadDocument LoadTwoPageDocument() =>
        PdfReadDocument.Open(PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page two"))
            .ToBytes());

    private static PdfDocument CreateCancelingDeferredDocument(
        CancellationTokenSource cancellation,
        Action onMaterializedAfterCancellation) =>
        PdfDocument.Create()
            .Deferred(_ => {
                cancellation.Cancel();
                return item => item.Paragraph(paragraph => paragraph.Text("Cancellation boundary"));
            })
            .Deferred(_ => {
                onMaterializedAfterCancellation();
                return item => item.Paragraph(paragraph => paragraph.Text("Must not materialize"));
            });

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
