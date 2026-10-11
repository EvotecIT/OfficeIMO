using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfVisualComparerTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Gallery_DistinguishesMatchingVisualPagesFromStructuralDifferences(bool extraExpectedPage) {
        byte[] single = BuildPdf("Common page");
        byte[] multiple = PdfDocument.Create(new PdfOptions { PageSize = new PageSize(240, 180) })
            .Paragraph(paragraph => paragraph.Text("Common page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Extra page"))
            .ToBytes();
        PdfVisualComparisonReport report = PdfVisualComparer.Compare(
            extraExpectedPage ? multiple : single, extraExpectedPage ? single : multiple);

        Assert.False(report.IsMatch);
        Assert.True(Assert.Single(report.Pages).IsMatch);
        Assert.Equal(2, report.StructuralDifferences.Count);
        Assert.Equal(0, report.DifferentPageCount);
        Assert.Equal(0, report.IncompletePageCount);
        string gallery = report.ToHtmlGallery();
        Assert.Contains("1 visual match", gallery, StringComparison.Ordinal);
        Assert.Contains("0 visual differences", gallery, StringComparison.Ordinal);
        Assert.Contains("2 structural differences", gallery, StringComparison.Ordinal);
    }

    [Fact]
    public void Compare_ReportsPixelDifferencesThresholdsIgnoredRegionsAndGallery() {
        byte[] expected = BuildPdf("Expected visual text");
        byte[] actual = BuildPdf("Changed visual text");

        PdfVisualComparisonReport exact = PdfDocument.Load(expected).CompareVisual(actual);
        PdfVisualPageComparison page = Assert.Single(exact.Pages);
        var ignoredOptions = new PdfVisualComparisonOptions();
        ignoredOptions.IgnoredRegions.Add(new PdfPixelRegion(0, 0, page.Width, page.Height));
        PdfVisualComparisonReport ignored = PdfVisualComparer.Compare(expected, actual, options: ignoredOptions);
        PdfVisualComparisonReport threshold = PdfVisualComparer.Compare(expected, actual, options: new PdfVisualComparisonOptions {
            AllowedDifferenceRatio = 1D
        });
        string gallery = exact.ToHtmlGallery("Review proof");

        Assert.False(exact.IsMatch);
        Assert.False(page.IsMatch);
        Assert.Equal(1, exact.DifferentPageCount);
        Assert.Equal(0, exact.IncompletePageCount);
        Assert.True(page.DifferentPixels > 0);
        Assert.True(page.DifferenceRatio > 0D);
        Assert.True(page.MaximumChannelDifference > 0);
        Assert.NotEmpty(page.DiffPng);
        Assert.False(page.HasSizeDifference);
        Assert.True(page.ChangedBounds.HasValue);
        Assert.InRange(page.ChangedBounds!.Value.Width, 1, page.Width);
        Assert.InRange(page.ChangedBounds.Value.Height, 1, page.Height);
        Assert.Null(Assert.Single(ignored.Pages).ChangedBounds);
        Assert.True(ignored.IsMatch);
        Assert.True(threshold.IsMatch);
        Assert.Equal(0, ignored.DifferentPageCount);
        Assert.Equal(0, threshold.DifferentPageCount);
        Assert.Contains("Review proof", gallery, StringComparison.Ordinal);
        Assert.Equal(3, Count(gallery, "data:image/png;base64,"));
    }

    [Fact]
    public void Compare_ReportsPageCountAndDimensionStructureWithCenterAlignment() {
        byte[] expected = PdfDocument.Create(new PdfOptions { PageSize = new PageSize(300, 400) })
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page two"))
            .ToBytes();
        byte[] actual = PdfDocument.Create(new PdfOptions { PageSize = new PageSize(320, 420) })
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .ToBytes();

        PdfVisualComparisonReport report = PdfVisualComparer.Compare(expected, actual, options: new PdfVisualComparisonOptions {
            Alignment = PdfVisualPageAlignment.Center,
            AllowedDifferenceRatio = 1D
        });

        Assert.False(report.IsMatch);
        Assert.Equal(2, report.ExpectedPageCount);
        Assert.Equal(1, report.ActualPageCount);
        Assert.Contains(report.StructuralDifferences, difference => difference.StartsWith("PageCount:", StringComparison.Ordinal));
        Assert.Contains(report.StructuralDifferences, difference => difference.StartsWith("Page 1 dimensions:", StringComparison.Ordinal));
        PdfVisualPageComparison page = Assert.Single(report.Pages);
        Assert.Equal(320, page.Width);
        Assert.Equal(420, page.Height);
        Assert.True(page.HasSizeDifference);
    }

    [Fact]
    public void SkippedPaintDoesNotCountAsAVisualMatch() {
        byte[] expected = UnsupportedOperatorPdf("UnknownPaintA");
        byte[] actual = UnsupportedOperatorPdf("UnknownPaintB");

        PdfVisualComparisonReport report = PdfVisualComparer.Compare(expected, actual);
        PdfVisualPageComparison page = Assert.Single(report.Pages);

        Assert.False(report.IsMatch);
        Assert.False(page.IsMatch);
        Assert.Equal(0, report.DifferentPageCount);
        Assert.Equal(1, report.IncompletePageCount);
        Assert.Equal(0, page.DifferentPixels);
        Assert.Contains(page.ExpectedCapabilityDiagnostics,
            diagnostic => diagnostic.Code == PdfRenderCapabilities.UnknownOperatorId);
        Assert.Contains("0 visual differences", report.ToHtmlGallery(), StringComparison.Ordinal);
        Assert.Contains("1 incomplete page", report.ToHtmlGallery(), StringComparison.Ordinal);
        Assert.Contains("warn\">Incomplete", report.ToHtmlGallery(), StringComparison.Ordinal);
        Assert.Contains(page.ActualCapabilityDiagnostics,
            diagnostic => diagnostic.Code == PdfRenderCapabilities.UnknownOperatorId);
    }

    [Fact]
    public void UnembeddedCustomFontDoesNotProveAVisualMatch() {
        const string content = "BT /F1 18 Tf 20 80 Td (Hello) Tj ET";
        byte[] pdf = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /CustomSans /Encoding /WinAnsiEncoding >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF", ""
        }));

        PdfVisualComparisonReport report = PdfVisualComparer.Compare(pdf, pdf);
        PdfVisualPageComparison page = Assert.Single(report.Pages);

        Assert.False(page.IsMatch);
        Assert.Equal(0, page.DifferentPixels);
        Assert.Equal(0, report.DifferentPageCount);
        Assert.Equal(1, report.IncompletePageCount);
        Assert.Contains(page.ExpectedCapabilityDiagnostics,
            diagnostic => diagnostic.Code == PdfRenderCapabilities.FontSubstitutionId);
        Assert.Contains("0 visual differences", report.ToHtmlGallery(), StringComparison.Ordinal);
        Assert.Contains("1 incomplete page", report.ToHtmlGallery(), StringComparison.Ordinal);
        Assert.Contains("warn\">Incomplete", report.ToHtmlGallery(), StringComparison.Ordinal);
        Assert.Equal(PdfPageChangeKind.ModifiedCandidate,
            Assert.Single(PdfPageChangeAnalyzer.Analyze(pdf, pdf).Changes).Kind);
    }

    [Fact]
    public void Compare_EnforcesPagePixelOutputAndCancellationBudgets() {
        byte[] pdf = BuildPdf("Bounded visual comparison");

        PdfReadLimitException pixels = Assert.Throws<PdfReadLimitException>(() =>
            PdfVisualComparer.Compare(pdf, pdf, options: new PdfVisualComparisonOptions {
                MaxPixelsPerImage = 1
            }));
        Assert.Equal(PdfReadLimitKind.RenderPixels, pixels.Kind);

        PdfReadLimitException output = Assert.Throws<PdfReadLimitException>(() =>
            PdfVisualComparer.Compare(pdf, pdf, options: new PdfVisualComparisonOptions {
                MaxTotalOutputBytes = 1
            }));
        Assert.Equal(PdfReadLimitKind.RenderBytes, output.Kind);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            PdfVisualComparer.Compare(pdf, pdf, cancellationToken: cancellation.Token));
        Assert.Throws<OperationCanceledException>(() =>
            PdfDocument.Load(pdf).Proof.CompareVisual(PdfDocument.Load(pdf), cancellation.Token));
    }

    [Fact]
    public void CompareRejectsIgnoredRegionWorkBeforeRasterizing() {
        byte[] pdf = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Text("Bounded ignored regions", 20D, 20D, 200D, 20D)).ToBytes();
        var options = new PdfVisualComparisonOptions();
        for (int index = 0; index < 3000; index++) {
            options.IgnoredRegions.Add(new PdfPixelRegion(0, 0, 240, 180));
        }

        PdfReadLimitException failure = Assert.Throws<PdfReadLimitException>(() =>
            PdfVisualComparer.Compare(pdf, pdf, options: options));

        Assert.Equal(PdfReadLimitKind.UnderstandingArtifacts, failure.Kind);
    }

    [Fact]
    public void CompareChargesIgnoredRegionWorkForTheCombinedCanvas() {
        byte[] wide = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 40D,
            MarginTop = 0D, MarginBottom = 0D, MarginLeft = 0D, MarginRight = 0D })
            .Canvas(canvas => canvas.Text("A", 0D, 0D, 10D, 10D)).ToBytes();
        byte[] tall = PdfDocument.Create(new PdfOptions { PageWidth = 40D, PageHeight = 400D,
            MarginTop = 0D, MarginBottom = 0D, MarginLeft = 0D, MarginRight = 0D })
            .Canvas(canvas => canvas.Text("A", 0D, 0D, 10D, 10D)).ToBytes();
        var options = new PdfVisualComparisonOptions();
        for (int index = 0; index < 4096; index++) options.IgnoredRegions.Add(new PdfPixelRegion(0, 0, 1, 1));

        PdfReadLimitException failure = Assert.Throws<PdfReadLimitException>(() =>
            PdfVisualComparer.Compare(wide, tall, options: options));

        Assert.Equal(PdfReadLimitKind.UnderstandingArtifacts, failure.Kind);
    }

    [Fact]
    public void CompareDoesNotChargeSourceRastersAsIgnoredRegionWork() {
        byte[] pdf = PdfDocument.Create(new PdfOptions { PageWidth = 500D, PageHeight = 500D })
            .Canvas(canvas => canvas.Text("A", 20D, 20D, 20D, 20D)).ToBytes();
        var options = new PdfVisualComparisonOptions();
        for (int index = 0; index < 300; index++) options.IgnoredRegions.Add(new PdfPixelRegion(0, 0, 500, 500));

        Assert.Single(PdfVisualComparer.Compare(pdf, pdf, options: options).Pages);
    }

    [Fact]
    public void ComparisonGallery_EnforcesUtf8OutputAndCancellationBudgetsDuringRendering() {
        byte[] pdf = BuildPdf("Bounded gallery");
        PdfVisualComparisonReport report = PdfVisualComparer.Compare(pdf, pdf);

        InvalidOperationException output = Assert.Throws<InvalidOperationException>(() =>
            report.ToHtmlGallery("Bounded gallery", maximumOutputBytes: 256L));
        Assert.Contains("being rendered", output.Message, StringComparison.Ordinal);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            report.ToHtmlGallery(
                "Cancelled gallery",
                maximumOutputBytes: 1024L * 1024L,
                cancellationToken: cancellation.Token));
    }

    private static byte[] BuildPdf(string text) => PdfDocument.Create(new PdfOptions { PageSize = new PageSize(240, 180) })
        .Paragraph(paragraph => paragraph.Text(text))
        .ToBytes();

    private static byte[] UnsupportedOperatorPdf(string operation) => System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
        "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
        "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Contents 4 0 R >>", "endobj",
        "4 0 obj", "<< /Length " + operation.Length + " >>", "stream", operation, "endstream", "endobj",
        "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF", ""
    }));

    private static int Count(string value, string token) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(token, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += token.Length;
        }

        return count;
    }
}
