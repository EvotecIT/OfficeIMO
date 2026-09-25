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
    public void MixedTextAndVectorChangeRetainsUnclassifiedVisualEvidence() {
        PdfDocument expected = TextAndVector("Original", OfficeIMO.Drawing.OfficeColor.Black);
        PdfDocument actual = TextAndVector("Revised", OfficeIMO.Drawing.OfficeColor.Red);

        PdfReviewPageComparison page = Assert.Single(expected.Proof.CompareReview(actual).Pages);

        Assert.Contains(page.Changes, static change => change.Kind == PdfReviewChangeKind.TextChanged);
        Assert.Contains(page.Changes, static change => change.Kind == PdfReviewChangeKind.UnclassifiedVisual);
    }

    [Fact]
    public void ToleratedPageDoesNotSpendRetainedArtifactByteBudget() {
        var options = new PdfReviewComparisonOptions();
        options.Visual.AllowedDifferenceRatio = 1D;
        options.Visual.MaxTotalOutputBytes = 1L;

        PdfReviewComparisonReport report = TextAndVector("Same", OfficeIMO.Drawing.OfficeColor.Black)
            .Proof.CompareReview(TextAndVector("Same", OfficeIMO.Drawing.OfficeColor.Red), options);

        Assert.True(report.IsMatch);
        Assert.Empty(report.Pages);
    }

    private static PdfDocument TextAndVector(string text, OfficeIMO.Drawing.OfficeColor color) {
        var rectangle = OfficeIMO.Drawing.OfficeShape.Rectangle(30D, 30D);
        rectangle.FillColor = color;
        rectangle.StrokeColor = null;
        return PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D })
            .Canvas(canvas => canvas.Text(text, 20D, 20D, 100D, 20D).Shape(rectangle, 180D, 110D))
            .ToBytes());
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
        Assert.Null(page.Visual);
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
    public void DetectsChangedActualTextWhitespaceWhenRenderedPixelsMatch() {
        PdfDocument expected = PdfDocument.Load(InvisibleTextPdf("Visible", actualTextHex: "FEFF004100200042"));
        PdfDocument actual = PdfDocument.Load(InvisibleTextPdf("Visible", actualTextHex: "FEFF004100A00042"));

        Assert.Equal("A B", Assert.Single(PdfReadDocument.Open(expected.ToBytes()).Pages[0].GetTextSpans()).SourceActualText);
        Assert.Equal("A\u00a0B", Assert.Single(PdfReadDocument.Open(actual.ToBytes()).Pages[0].GetTextSpans()).SourceActualText);

        PdfReviewPageComparison page = Assert.Single(expected.Proof.CompareReview(actual).Pages);

        PdfReviewChange change = Assert.Single(page.Changes);
        Assert.Equal(PdfReviewChangeKind.TextChanged, change.Kind);
        Assert.Equal("A B", change.ExpectedText);
        Assert.Equal("A\u00a0B", change.ActualText);
    }

    [Fact]
    public void DetectsChangedActualTextAfterRotatedTextProjection() {
        PdfDocument expected = PdfDocument.Load(InvisibleTextPdf("Visible", rotated: true, actualTextHex: "FEFF004100200042"));
        PdfDocument actual = PdfDocument.Load(InvisibleTextPdf("Visible", rotated: true, actualTextHex: "FEFF004100A00042"));

        PdfReviewChange change = Assert.Single(Assert.Single(expected.Proof.CompareReview(actual).Pages).Changes);

        Assert.Equal(PdfReviewChangeKind.TextChanged, change.Kind);
        Assert.Equal("A B", change.ExpectedText);
        Assert.Equal("A\u00a0B", change.ActualText);
    }

    [Fact]
    public void SmallInvisibleTextMoveRemainsASemanticChange() {
        PdfDocument expected = PdfDocument.Load(InvisibleTextPdf("Searchable  text", y: 90D));
        PdfDocument moved = PdfDocument.Load(InvisibleTextPdf("Searchable  text", y: 92D));

        Assert.Contains(Assert.Single(expected.Proof.CompareReview(moved).Pages).Changes,
            static change => change.Kind == PdfReviewChangeKind.TextMoved);
    }

    [Fact]
    public void AlignsReorderedPagesUsingTheSelectedIgnoreRegion() {
        PdfDocument expected = PagesWithHeaders(("Old A", "Alpha"), ("Old B", "Bravo"));
        PdfDocument actual = PagesWithHeaders(("New B", "Bravo"), ("New A", "Alpha"));
        var options = new PdfReviewComparisonOptions();
        options.Visual.IgnoredRegions.Add(new PdfPixelRegion(0, 0, 240, 55));

        PdfReviewComparisonReport report = expected.Proof.CompareReview(actual, options);

        Assert.Empty(report.Pages);
        Assert.Contains(report.PageAlignment.Changes, static change =>
            change.ExpectedPageNumber == 1 && change.ActualPageNumber == 2 && change.UsesIgnoredRegions && !change.IsExactRenderedMatch);
        Assert.Contains(report.PageAlignment.Changes, static change =>
            change.ExpectedPageNumber == 2 && change.ActualPageNumber == 1 && change.UsesIgnoredRegions && !change.IsExactRenderedMatch);
    }

    [Fact]
    public void ExactPageMatchDoesNotClaimIgnoredPixelsWereNeeded() {
        PdfDocument document = Page("Same", 20D);
        var options = new PdfReviewComparisonOptions();
        options.Visual.IgnoredRegions.Add(new PdfPixelRegion(0, 0, 240, 55));

        PdfReviewComparisonReport report = document.Proof.CompareReview(PdfDocument.Load(document.ToBytes()), options);

        PdfPageChange alignment = Assert.Single(report.PageAlignment.Changes);
        Assert.True(alignment.IsExactRenderedMatch);
        Assert.False(alignment.UsesIgnoredRegions);
        Assert.True(report.IsMatch);
    }

    [Fact]
    public void CountsSemanticOnlyChangedPagesAgainstThePairBudget() {
        PdfDocument expected = PdfDocument.Merge(new PdfMergeOptions(),
            PdfDocument.Load(InvisibleTextPdf("Original one")), PdfDocument.Load(InvisibleTextPdf("Original two")));
        PdfDocument actual = PdfDocument.Merge(new PdfMergeOptions(),
            PdfDocument.Load(InvisibleTextPdf("Revised one")), PdfDocument.Load(InvisibleTextPdf("Revised two")));

        Assert.Throws<PdfReadLimitException>(() => expected.Proof.CompareReview(actual,
            new PdfReviewComparisonOptions { MaxChangedPagePairs = 1 }));
    }

    [Fact]
    public void KeepsRotatedInvisibleTextInSemanticComparison() {
        PdfDocument expected = PdfDocument.Load(InvisibleTextPdf("Original", rotated: true));
        PdfDocument actual = PdfDocument.Load(InvisibleTextPdf("Revised", rotated: true));

        PdfReviewComparisonReport report = expected.Proof.CompareReview(actual);

        Assert.False(report.IsMatch);
        Assert.Contains(Assert.Single(report.Pages).Changes, static change =>
            change.Kind == PdfReviewChangeKind.TextChanged);
    }

    [Fact]
    public void OrdersMixedTextChangesByPagePosition() {
        PdfDocument expected = PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D })
            .Canvas(canvas => {
                canvas.Text("Remove", 20D, 15D, 130D, 20D);
                canvas.Text("Original", 20D, 90D, 130D, 20D);
            }).ToBytes());
        PdfDocument actual = PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D })
            .Canvas(canvas => canvas.Text("Revised", 20D, 90D, 130D, 20D)).ToBytes());

        PdfReviewChangeKind[] kinds = Assert.Single(expected.Proof.CompareReview(actual).Pages).Changes
            .Select(static change => change.Kind).ToArray();

        Assert.Equal(new[] { PdfReviewChangeKind.TextRemoved, PdfReviewChangeKind.TextChanged }, kinds);
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

        PdfReviewComparisonReport stampedScan = ScanWithStamp(blue).Proof.CompareReview(ScanWithStamp(red));
        Assert.Contains(Assert.Single(stampedScan.Pages).Changes,
            static change => change.Kind == PdfReviewChangeKind.ScannedPageUncertain);
    }

    [Theory]
    [InlineData(23D, 20D, 20D, 20D)]
    [InlineData(20D, 20D, 24D, 20D)]
    public void ReportsImagePlacementChangeEvenWhenWhitePixelsStayIdentical(
        double actualX, double actualY, double actualWidth, double actualHeight) {
        byte[] white = PdfPngTestImages.CreateRgbPng(255, 255, 255);
        PdfDocument expected = PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D })
            .Canvas(canvas => canvas.Image(white, 20D, 20D, 20D, 20D)).ToBytes());
        PdfDocument actual = PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D })
            .Canvas(canvas => canvas.Image(white, actualX, actualY, actualWidth, actualHeight)).ToBytes());

        PdfReviewComparisonReport report = expected.Proof.CompareReview(actual);

        Assert.False(report.IsMatch);
        Assert.Contains(Assert.Single(report.Pages).Changes, static change => change.Kind == PdfReviewChangeKind.ImageMoved);
    }

    [Fact]
    public void RemovingTheFirstOfTwoIdenticalTextBlocksDoesNotMoveTheSecond() {
        PdfDocument expected = PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D })
            .Canvas(canvas => {
                canvas.Text("Duplicate", 20D, 20D, 100D, 20D);
                canvas.Text("Duplicate", 20D, 100D, 100D, 20D);
            }).ToBytes());
        PdfDocument actual = PageAt("Duplicate", 100D);

        PdfReviewChange change = Assert.Single(Assert.Single(expected.Proof.CompareReview(actual).Pages).Changes);
        Assert.Equal(PdfReviewChangeKind.TextRemoved, change.Kind);
    }

    [Fact]
    public void RemovingTheFirstOfTwoIdenticalImagesDoesNotMoveTheSecond() {
        byte[] image = PdfPngTestImages.CreateRgbPng(20, 60, 180);
        PdfDocument expected = PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D })
            .Canvas(canvas => {
                canvas.Image(image, 20D, 20D, 40D, 40D);
                canvas.Image(image, 20D, 100D, 40D, 40D);
            }).ToBytes());
        PdfDocument actual = PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D })
            .Canvas(canvas => canvas.Image(image, 20D, 100D, 40D, 40D)).ToBytes());

        PdfReviewChange change = Assert.Single(Assert.Single(expected.Proof.CompareReview(actual).Pages).Changes);
        Assert.Equal(PdfReviewChangeKind.ImageRemoved, change.Kind);
    }

    [Fact]
    public void ScanClassificationStillEnforcesImagePlacementBudget() {
        byte[] blue = PdfPngTestImages.CreateRgbPng(20, 60, 180);
        PdfDocument expected = ImagePage(blue, scan: true);
        PdfDocument actual = PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D })
            .Canvas(canvas => {
                canvas.Image(blue, 0D, 0D, 240D, 180D);
                canvas.Image(blue, 10D, 10D, 20D, 20D);
            }).ToBytes());

        Assert.Throws<PdfReadLimitException>(() => expected.Proof.CompareReview(actual,
            new PdfReviewComparisonOptions { MaxImagePlacementsPerPage = 1 }));
    }

    [Fact]
    public void AdjacentImageTilesAreClassifiedAsAScan() {
        byte[] blue = PdfPngTestImages.CreateRgbPng(20, 60, 180);
        byte[] red = PdfPngTestImages.CreateRgbPng(180, 30, 20);
        PdfDocument expected = TiledScan(blue);
        PdfDocument actual = TiledScan(red);

        PdfReviewChange change = Assert.Single(Assert.Single(expected.Proof.CompareReview(actual).Pages).Changes);
        Assert.Equal(PdfReviewChangeKind.ScannedPageUncertain, change.Kind);
    }

    private static PdfDocument TiledScan(byte[] image) => PdfDocument.Load(PdfDocument.Create(
        new PdfOptions { PageWidth = 240D, PageHeight = 180D })
        .Canvas(canvas => {
            canvas.Image(image, 0D, 0D, 120D, 180D);
            canvas.Image(image, 120D, 0D, 120D, 180D);
        }).ToBytes());

    [Fact]
    public void MostlyOffPageImageIsComparedAsAnImageRatherThanAScan() {
        PdfDocument expected = PdfDocument.Load(OffPageImagePdf("ABC"));
        PdfDocument actual = PdfDocument.Load(OffPageImagePdf("DEF"));

        Assert.Contains(Assert.Single(expected.Proof.CompareReview(actual).Pages).Changes,
            static change => change.Kind == PdfReviewChangeKind.ImageChangedCandidate);
    }

    [Fact]
    public void ZeroOpacityImagePayloadDoesNotCreateSemanticChangeOrScan() {
        PdfDocument expected = PdfDocument.Load(ImageWithPaintStatePdf("ABC",
            "q /GS0 gs 240 0 0 180 0 0 cm /Im0 Do Q\n", withZeroOpacity: true));
        PdfDocument actual = PdfDocument.Load(ImageWithPaintStatePdf("DEF",
            "q /GS0 gs 240 0 0 180 0 0 cm /Im0 Do Q\n", withZeroOpacity: true));

        PdfReviewComparisonReport report = expected.Proof.CompareReview(actual);

        Assert.True(report.IsMatch);
        Assert.Empty(report.Pages);
    }

    [Fact]
    public void IgnoredRegionCoversEffectiveClippedImageBounds() {
        const string content = "q 20 20 40 40 re W n 240 0 0 180 0 0 cm /Im0 Do Q\n";
        PdfDocument expected = PdfDocument.Load(ImageWithPaintStatePdf("ABC", content));
        PdfDocument actual = PdfDocument.Load(ImageWithPaintStatePdf("DEF", content));
        var options = new PdfReviewComparisonOptions();
        options.Visual.IgnoredRegions.Add(new PdfPixelRegion(20, 120, 40, 40));

        PdfReviewComparisonReport report = expected.Proof.CompareReview(actual, options);

        Assert.True(report.IsMatch);
        Assert.Empty(report.Pages);
    }

    [Fact]
    public void PartiallyMaskedScanWithChangedImagePayloadRemainsUncertain() {
        const string content = "q 240 0 0 180 0 0 cm /Im0 Do Q\n";
        PdfDocument expected = PdfDocument.Load(ImageWithPaintStatePdf("ABCDEF", content, width: 2));
        PdfDocument actual = PdfDocument.Load(ImageWithPaintStatePdf("XYZDEF", content, width: 2));
        var options = new PdfReviewComparisonOptions();
        options.Visual.IgnoredRegions.Add(new PdfPixelRegion(0, 0, 130, 180));

        PdfReviewComparisonReport report = expected.Proof.CompareReview(actual, options);

        Assert.False(report.IsMatch);
        Assert.Contains(Assert.Single(report.Pages).Changes,
            static change => change.Kind == PdfReviewChangeKind.ScannedPageUncertain);
    }

    private static byte[] ImageWithPaintStatePdf(string pixels, string content, bool withZeroOpacity = false, int width = 1) =>
        System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /XObject << /Im0 5 0 R >>" +
                (withZeroOpacity ? " /ExtGState << /GS0 6 0 R >>" : string.Empty) + " >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + System.Text.Encoding.ASCII.GetByteCount(content) + " >>", "stream", content.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", $"<< /Type /XObject /Subtype /Image /Width {width} /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length {pixels.Length} >>", "stream", pixels, "endstream", "endobj",
            withZeroOpacity ? "6 0 obj\n<< /Type /ExtGState /ca 0 >>\nendobj" : string.Empty,
            "trailer", withZeroOpacity ? "<< /Root 1 0 R /Size 7 >>" : "<< /Root 1 0 R /Size 6 >>", "%%EOF", string.Empty
        }));

    private static PdfDocument PageAt(string text, double y) => PdfDocument.Load(PdfDocument.Create(
        new PdfOptions { PageWidth = 240D, PageHeight = 180D })
        .Canvas(canvas => canvas.Text(text, 20D, y, 100D, 20D)).ToBytes());

    private static byte[] OffPageImagePdf(string pixels) {
        const string content = "q 240 0 0 180 230 0 cm /Im0 Do Q\n";
        return System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.7",
        "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
        "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
        "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /XObject << /Im0 5 0 R >> >> /Contents 4 0 R >>", "endobj",
        "4 0 obj", "<< /Length " + System.Text.Encoding.ASCII.GetByteCount(content) + " >>", "stream", content.TrimEnd('\n'), "endstream", "endobj",
        "5 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", pixels, "endstream", "endobj",
        "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF", string.Empty
        }));
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

    private static PdfDocument ScanWithStamp(byte[] png) => PdfDocument.Load(PdfDocument.Create(
        new PdfOptions { PageWidth = 240D, PageHeight = 180D })
        .Canvas(canvas => {
            canvas.Image(png, 0D, 0D, 240D, 180D);
            canvas.Text("Reviewed", 20D, 20D, 100D, 20D);
        }).ToBytes());

    private static PdfDocument PagesWithHeaders(params (string Header, string Body)[] pages) {
        PdfDocument document = PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D });
        for (int index = 0; index < pages.Length; index++) {
            if (index > 0) document.PageBreak();
            (string header, string body) = pages[index];
            document.Canvas(canvas => {
                canvas.Text(header, 20D, 10D, 160D, 20D);
                canvas.Text(body, 20D, 90D, 160D, 25D);
            });
        }
        return PdfDocument.Load(document.ToBytes());
    }

    private static PdfDocument Pages(params string[] texts) {
        PdfDocument document = PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D });
        for (int index = 0; index < texts.Length; index++) {
            if (index > 0) document.PageBreak();
            string text = texts[index];
            document.Canvas(canvas => canvas.Text(text, 20D, 30D, 160D, 25D));
        }
        return PdfDocument.Load(document.ToBytes());
    }

    private static byte[] InvisibleTextPdf(string text, bool rotated = false, double y = 90D,
        string? actualTextHex = null) {
        string position = rotated ? "0 1 -1 0 100 20 Tm " : "20 " + y.ToString(System.Globalization.CultureInfo.InvariantCulture) + " Td ";
        string prefix = actualTextHex is null ? string.Empty : "/Span << /ActualText <" + actualTextHex + "> >> BDC ";
        string suffix = actualTextHex is null ? string.Empty : " EMC";
        byte[] content = System.Text.Encoding.ASCII.GetBytes(prefix + "BT /F1 12 Tf 3 Tr " + position + "(" + text + ") Tj ET" + suffix + "\n");
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
