using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReviewComparerTests {
    [Fact]
    public void ClassifiesChangedTextAndRetainsRenderedPageProof() {
        PdfDocument expected = Page("Original", 20D);
        PdfDocument actual = Page("Revised", 20D);

        PdfReviewComparisonReport report = expected.Proof.CompareReview(actual);

        Assert.False(report.IsMatch);
        Assert.Equal(PdfPageChangeKind.ModifiedCandidate, Assert.Single(report.PageAlignment.Changes).Kind);
        PdfReviewPageComparison page = Assert.Single(report.Pages);
        PdfReviewChange change = Assert.Single(page.Changes);
        Assert.Equal(PdfReviewChangeKind.TextChanged, change.Kind);
        Assert.Equal("Original", change.ExpectedText);
        Assert.Equal("Revised", change.ActualText);
        Assert.NotNull(change.ExpectedBounds);
        Assert.NotNull(change.ActualBounds);
        PdfVisualPageComparison visual = Assert.IsType<PdfVisualPageComparison>(page.Visual);
        Assert.NotNull(visual.ChangedBounds);
        Assert.NotEmpty(visual.DiffPng);
    }

    [Fact]
    public void ClassifiesMovedTextAndReportsPageInsertionSeparately() {
        PdfReviewComparisonReport moved = Page("Move me", 20D).Proof.CompareReview(Page("Move me", 80D));
        Assert.Equal(PdfReviewChangeKind.TextMoved, Assert.Single(Assert.Single(moved.Pages).Changes).Kind);

        PdfDocument expected = Pages("Alpha", "Bravo");
        PdfDocument actual = Pages("Intro", "Alpha", "Bravo");
        PdfReviewComparisonReport inserted = expected.Proof.CompareReview(actual);
        Assert.Empty(inserted.Pages);
        Assert.Single(inserted.PageAlignment.Changes, static change => change.Kind == PdfPageChangeKind.Inserted);
        Assert.False(inserted.IsMatch);
    }

    [Fact]
    public void EnforcesChangedPairBudgetBeforeDetailedComparison() {
        PdfDocument expected = Pages("First", "Second");
        PdfDocument actual = Pages("Changed first", "Changed second");
        Assert.Throws<PdfReadLimitException>(() => expected.Proof.CompareReview(actual,
            new PdfReviewComparisonOptions { MaxChangedPagePairs = 1 }));
    }

    [Fact]
    public void IgnoredPixelRegionAlsoExcludesItsSemanticTextChange() {
        PdfDocument expected = Page("Original", 20D);
        PdfDocument actual = Page("Revised", 20D);
        var options = new PdfReviewComparisonOptions();
        options.Visual.IgnoredRegions.Add(new PdfPixelRegion(0, 0, 240, 180));

        PdfReviewComparisonReport report = expected.Proof.CompareReview(actual, options);

        Assert.True(report.IsMatch);
        Assert.Empty(report.Pages);
    }

    [Fact]
    public void PartialPixelMaskRetainsTextBlockChangeForReview() {
        PdfDocument expected = Page("Invoice 123", 20D);
        PdfDocument actual = Page("Invoice 999", 20D);
        PdfPixelRegion changed = Assert.IsType<PdfPixelRegion>(
            expected.Proof.CompareVisualPages(1, actual, 1).ChangedBounds);
        var options = new PdfReviewComparisonOptions();
        options.Visual.IgnoredRegions.Add(changed);

        PdfReviewComparisonReport report = expected.Proof.CompareReview(actual, options);

        Assert.False(report.IsMatch);
        PdfReviewPageComparison page = Assert.Single(report.Pages);
        Assert.True(Assert.IsType<PdfVisualPageComparison>(page.Visual).IsMatch);
        Assert.Equal(PdfReviewChangeKind.TextChanged, Assert.Single(page.Changes).Kind);
    }

    [Fact]
    public void DetectsChangedSearchableTextWhenRenderedPixelsMatch() {
        PdfDocument expected = PdfDocument.Load(InvisibleTextPdf("Searchable original"));
        PdfDocument actual = PdfDocument.Load(InvisibleTextPdf("Searchable revised"));

        PdfReviewComparisonReport report = expected.Proof.CompareReview(actual);

        Assert.Equal(PdfPageChangeKind.Unchanged, Assert.Single(report.PageAlignment.Changes).Kind);
        PdfReviewPageComparison page = Assert.Single(report.Pages);
        Assert.Null(page.Visual);
        Assert.Equal(PdfReviewChangeKind.TextChanged, Assert.Single(page.Changes).Kind);
        Assert.False(report.IsMatch);
    }

    [Fact]
    public void EnforcesVisualPairAndOutputBudgetsAcrossTheReview() {
        PdfDocument expected = Pages("First", "Second");
        PdfDocument actual = Pages("Changed first", "Changed second");
        var pairLimit = new PdfReviewComparisonOptions();
        pairLimit.Visual.MaxPages = 1;
        Assert.Throws<PdfReadLimitException>(() => expected.Proof.CompareReview(actual, pairLimit));

        var outputLimit = new PdfReviewComparisonOptions();
        outputLimit.Visual.MaxTotalOutputBytes = 1;
        Assert.Throws<PdfReadLimitException>(() => expected.Proof.CompareReview(actual, outputLimit));
    }

    [Fact]
    public void RejectsDifferentAlignmentAndVisualRasterPolicies() {
        PdfDocument document = Page("Same", 20D);
        var options = new PdfReviewComparisonOptions();
        options.Visual.Scale = 2D;

        Assert.Throws<ArgumentException>(() => document.Proof.CompareReview(document, options));
    }

    [Fact]
    public void ClassifiesChangedImagePayloadAndKeepsScannedPagesUncertain() {
        byte[] blue = PdfPngTestImages.CreateRgbPng(20, 60, 180);
        byte[] red = PdfPngTestImages.CreateRgbPng(180, 30, 20);
        PdfDocument expected = ImagePage(blue, scan: false);
        PdfDocument actual = ImagePage(red, scan: false);

        PdfReviewComparisonReport image = expected.Proof.CompareReview(actual);
        Assert.Contains(Assert.Single(image.Pages).Changes,
            static change => change.Kind == PdfReviewChangeKind.ImageChangedCandidate);

        PdfReviewComparisonReport scanned = ImagePage(blue, scan: true).Proof.CompareReview(ImagePage(red, scan: true));
        Assert.Equal(PdfReviewChangeKind.ScannedPageUncertain, Assert.Single(Assert.Single(scanned.Pages).Changes).Kind);
    }

    private static PdfDocument Page(string text, double x) => PdfDocument.Load(PdfDocument.Create(
        new PdfOptions { PageWidth = 240D, PageHeight = 180D })
        .Canvas(canvas => canvas.Text(text, x, 30D, 130D, 25D)).ToBytes());

    private static PdfDocument ImagePage(byte[] png, bool scan) => PdfDocument.Load(PdfDocument.Create(
        new PdfOptions { PageWidth = 240D, PageHeight = 180D })
        .Canvas(canvas => {
            if (!scan) canvas.Text("Proof", 20D, 20D, 80D, 20D);
            canvas.Image(png, scan ? 0D : 20D, scan ? 0D : 50D,
                scan ? 240D : 40D, scan ? 180D : 40D);
        }).ToBytes());

    private static PdfDocument Pages(params string[] texts) {
        PdfDocument document = PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D });
        for (int index = 0; index < texts.Length; index++) {
            if (index > 0) document.PageBreak();
            string text = texts[index];
            document.Canvas(canvas => canvas.Text(text, 20D, 30D, 160D, 25D));
        }
        return PdfDocument.Load(document.ToBytes());
    }

    private static byte[] InvisibleTextPdf(string text) {
        byte[] content = System.Text.Encoding.ASCII.GetBytes("BT /F1 12 Tf 3 Tr 20 90 Td (" + text + ") Tj ET\n");
        using var stream = new System.IO.MemoryStream();
        void Write(string value) {
            byte[] bytes = System.Text.Encoding.ASCII.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
        }
        Write("%PDF-1.7\n");
        Write("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        Write("2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        Write("3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        Write("4 0 obj\n<< /Length " + content.Length + " >>\nstream\n");
        stream.Write(content, 0, content.Length);
        Write("endstream\nendobj\n");
        Write("5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n");
        Write("trailer\n<< /Root 1 0 R /Size 6 >>\n%%EOF\n");
        return stream.ToArray();
    }
}
