using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfWord = UglyToad.PdfPig.Content.Word;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordPdfRunningPictureWrapTests {
    private const double PictureLeft = 144D;
    private const double PictureWidth = 96D;
    private const double PictureHeight = 48D;
    private const double DistanceLeft = 12D;
    private const double DistanceRight = 18D;
    private const double DistanceTop = 6D;
    private const double DistanceBottom = 18D;
    private const double GeometryTolerance = .01D;
    private static readonly string[] FlowTokens = new[] { "FLOW_START" }
        .Concat(Enumerable.Range(1, 48).Select(index => "W" + index.ToString("00")))
        .Concat(new[] { "FLOW_END" }).ToArray();

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_SquareRunningPictureUsesBothFragmentsOnOneBaseline(bool footer) {
        using WordDocument document = CreateDocument(footer, WordImageTextWrapping.Square, out _);
        string sourceXml = SourceXml(document);
        Assert.Empty(document.ValidateDocument());

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());
        AssertNoPictureWrapFallback(result);
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var picture = Assert.Single(page.GetImages());
        Assert.Equal(PictureLeft, picture.BoundingBox.Left, 3);
        Assert.Equal(PictureWidth, picture.BoundingBox.Width, 3);
        Assert.Equal(PictureHeight, picture.BoundingBox.Height, 3);
        PdfWord[] words = AssertFlowTokens(page.GetWords());
        double excludedLeft = picture.BoundingBox.Left - DistanceLeft;
        double excludedRight = picture.BoundingBox.Right + DistanceRight;
        double excludedBottom = picture.BoundingBox.Bottom - DistanceBottom;
        double excludedTop = picture.BoundingBox.Top + DistanceTop;
        var intersectingRows = words.Where(word => word.BoundingBox.Top > excludedBottom &&
                word.BoundingBox.Bottom < excludedTop)
            .GroupBy(word => Math.Round(word.Letters[0].StartBaseLine.Y, 3)).ToArray();
        Assert.NotEmpty(intersectingRows);
        Assert.Contains(intersectingRows, row =>
            row.Any(word => word.BoundingBox.Right <= excludedLeft + GeometryTolerance) &&
            row.Any(word => word.BoundingBox.Left >= excludedRight - GeometryTolerance));
        Assert.All(intersectingRows.SelectMany(row => row), word => Assert.True(
            word.BoundingBox.Right <= excludedLeft + GeometryTolerance ||
            word.BoundingBox.Left >= excludedRight - GeometryTolerance,
            "A wrapped word must stay outside the authored picture distances."));
        Assert.Contains(intersectingRows.SelectMany(row => row), word =>
            Math.Abs(word.Letters[0].StartBaseLine.X - excludedRight) < GeometryTolerance);
        if (!footer) {
            var body = Assert.Single(page.GetWords(), word => word.Text == "BODY");
            Assert.True(body.BoundingBox.Top <= words.Min(word => word.BoundingBox.Bottom) + GeometryTolerance,
                "The body must begin below the expanded header text, not overlap its final row.");
        }
        AssertNoPictureWrapFallback(result);
        Assert.Equal(sourceXml, SourceXml(document));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_TopAndBottomRunningPictureClearsTheWholeTextFrame(bool footer) {
        using WordDocument document = CreateDocument(footer, WordImageTextWrapping.TopAndBottom, out _);
        string sourceXml = SourceXml(document);
        Assert.Empty(document.ValidateDocument());

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());
        AssertNoPictureWrapFallback(result);
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var picture = Assert.Single(page.GetImages());
        Assert.Equal(PictureLeft, picture.BoundingBox.Left, 3);
        Assert.Equal(PictureWidth, picture.BoundingBox.Width, 3);
        Assert.Equal(PictureHeight, picture.BoundingBox.Height, 3);
        PdfWord[] words = AssertFlowTokens(page.GetWords());
        Assert.All(words, word => Assert.True(
            word.BoundingBox.Top <= picture.BoundingBox.Bottom - DistanceBottom + GeometryTolerance,
            "TopAndBottom must clear the image and authored bottom distance across the entire text frame."));
        Assert.Equal(72D, words[0].Letters[0].StartBaseLine.X, 3);
        if (!footer) {
            var body = Assert.Single(page.GetWords(), word => word.Text == "BODY");
            Assert.True(body.BoundingBox.Top <= words.Min(word => word.BoundingBox.Bottom) + GeometryTolerance);
        }
        AssertNoPictureWrapFallback(result);
        Assert.Equal(sourceXml, SourceXml(document));
    }

    [Theory]
    [InlineData("missing-distance")]
    [InlineData("child-distance")]
    [InlineData("one-side")]
    [InlineData("rotation")]
    [InlineData("table-cell")]
    public void SaveAsPdf_UnqualifiedPictureWrapReportsItsFallbackBeforeSerialization(string profile) {
        using WordDocument document = CreateDocument(false, WordImageTextWrapping.Square, out WordImage image,
            table: profile == "table-cell");
        DW.Anchor anchor = image._Image.Anchor!;
        switch (profile) {
            case "missing-distance": anchor.DistanceFromLeft = null; break;
            case "child-distance":
                anchor.GetFirstChild<DW.WrapSquare>()!.SetAttribute(new OpenXmlAttribute("distL", string.Empty, "0"));
                break;
            case "one-side": anchor.GetFirstChild<DW.WrapSquare>()!.WrapText = DW.WrapTextValues.Left; break;
            case "rotation": image.Rotation = 17; break;
            case "table-cell": break;
            default: throw new ArgumentOutOfRangeException(nameof(profile));
        }
        Assert.Empty(document.ValidateDocument());
        string sourceXml = SourceXml(document);

        WordPdfExportOptions options = Options();
        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(options);
        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code == "native-anchored-image-wrap-approximated");
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Single(page.GetImages());
        AssertFlowTokens(page.GetWords());
        Assert.Single(page.GetWords(), word => word.Text == "BODY");
        // The returned operation report already owns this warning. A deferred
        // callback must not be its first observation or add another copy.
        Assert.Single(result.Warnings, item => item.Code == "native-anchored-image-wrap-approximated");
        Assert.Equal(sourceXml, SourceXml(document));
        if (profile == "child-distance") {
            // Reusing the caller's options creates a fresh operation report;
            // source identity must not suppress this operation's warning.
            PdfDocumentConversionResult repeated = document.ConvertToPdfWithEvidence(options);
            PdfConversionWarning repeatedWarning = Assert.Single(repeated.Warnings,
                item => item.Code == "native-anchored-image-wrap-approximated");
            Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, repeatedWarning.LossKind);
            Assert.Throws<InvalidOperationException>(() => repeated.RequireNoLoss());
            using var repeatedPdf = PdfPigDocument.Open(repeated.ToBytes());
            var repeatedPage = Assert.Single(repeatedPdf.GetPages());
            Assert.Single(repeatedPage.GetImages());
            AssertFlowTokens(repeatedPage.GetWords());
            Assert.Single(repeated.Warnings, item => item.Code == "native-anchored-image-wrap-approximated");
            Assert.Single(result.Warnings, item => item.Code == "native-anchored-image-wrap-approximated");
            Assert.Equal(sourceXml, SourceXml(document));
        }
    }

    [Fact]
    public void SaveAsPdf_ReusedHeaderPictureRetainsOnePaintPerPageAcrossOperations() {
        using WordDocument document = CreateDocument(false, WordImageTextWrapping.Square, out _);
        WordParagraph nextPage = document.AddParagraph();
        ConfigureParagraph(nextPage, "SECOND_PAGE");
        nextPage.PageBreakBefore = true;
        Assert.Empty(document.ValidateDocument());
        string sourceXml = SourceXml(document);
        WordPdfExportOptions options = Options();

        PdfDocumentConversionResult first = document.ConvertToPdfWithEvidence(options);
        AssertNoPictureWrapFallback(first);
        using var firstPdf = PdfPigDocument.Open(first.ToBytes());
        Assert.Equal(2, firstPdf.NumberOfPages);
        PdfDocumentConversionResult second = document.ConvertToPdfWithEvidence(options);
        AssertNoPictureWrapFallback(second);
        using var secondPdf = PdfPigDocument.Open(second.ToBytes());
        Assert.Equal(2, secondPdf.NumberOfPages);
        for (int pageNumber = 1; pageNumber <= 2; pageNumber++) {
            var firstPage = firstPdf.GetPage(pageNumber);
            var secondPage = secondPdf.GetPage(pageNumber);
            var firstImage = Assert.Single(firstPage.GetImages());
            var secondImage = Assert.Single(secondPage.GetImages());
            Assert.Equal(PictureLeft, firstImage.BoundingBox.Left, 3);
            Assert.Equal(PictureWidth, firstImage.BoundingBox.Width, 3);
            Assert.Equal(PictureHeight, firstImage.BoundingBox.Height, 3);
            Assert.Equal(firstImage.BoundingBox.Left, secondImage.BoundingBox.Left, 3);
            Assert.Equal(firstImage.BoundingBox.Bottom, secondImage.BoundingBox.Bottom, 3);
            Assert.Equal(firstImage.BoundingBox.Right, secondImage.BoundingBox.Right, 3);
            Assert.Equal(firstImage.BoundingBox.Top, secondImage.BoundingBox.Top, 3);
            PdfWord[] firstFlow = AssertFlowTokens(firstPage.GetWords());
            PdfWord[] secondFlow = AssertFlowTokens(secondPage.GetWords());
            for (int index = 0; index < FlowTokens.Length; index++) {
                Assert.Equal(firstFlow[index].Letters[0].StartBaseLine.X,
                    secondFlow[index].Letters[0].StartBaseLine.X, 3);
                Assert.Equal(firstFlow[index].Letters[0].StartBaseLine.Y,
                    secondFlow[index].Letters[0].StartBaseLine.Y, 3);
            }
        }
        Assert.Single(firstPdf.GetPage(1).GetWords(), word => word.Text == "BODY");
        Assert.Single(firstPdf.GetPage(2).GetWords(), word => word.Text == "SECOND_PAGE");
        AssertNoPictureWrapFallback(first);
        AssertNoPictureWrapFallback(second);
        Assert.Equal(sourceXml, SourceXml(document));
    }

    [Fact]
    public void SaveAsPdf_WrappedRunningPictureReportsPhysicalPageClippingBeforeSerialization() {
        using WordDocument document = CreateDocument(false, WordImageTextWrapping.Square, out WordImage image);
        // The supported page-relative rectangle crosses the physical edge;
        // this remains an omission, independently of its supported wrap mode.
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        image.HorizontalPositionOffset = 580L * 12700L;
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        image.VerticalPositionOffset = 24L * 12700L;
        Assert.Empty(document.ValidateDocument());
        string sourceXml = SourceXml(document);

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());
        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code == "native-anchored-image-clipped");
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        AssertNoPictureWrapFallback(result);
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var picture = Assert.Single(page.GetImages());
        Assert.Equal(580D, picture.BoundingBox.Left, 3);
        Assert.Equal(PictureWidth, picture.BoundingBox.Width, 3);
        Assert.True(picture.BoundingBox.Left < page.Width);
        Assert.True(picture.BoundingBox.Right > page.Width,
            "The saved PDF must retain the source rectangle that is clipped by its physical page edge.");
        AssertFlowTokens(page.GetWords());
        Assert.Single(result.Warnings, item => item.Code == "native-anchored-image-clipped");
        AssertNoPictureWrapFallback(result);
        Assert.Equal(sourceXml, SourceXml(document));
    }

    [Fact]
    public void SaveAsPdf_WrappedRunningPictureRetainsCropMirrorsAndOneImagePaint() {
        using WordDocument document = CreateDocument(false, WordImageTextWrapping.Square, out WordImage image);
        image.CropLeft = 25000;
        image.CropRight = 12500;
        image.CropTop = 12500;
        image.CropBottom = 25000;
        image.HorizontalFlip = true;
        image.VerticalFlip = true;
        Assert.Empty(document.ValidateDocument());
        string sourceXml = SourceXml(document);

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());
        AssertNoPictureWrapFallback(result);
        byte[] bytes = result.ToBytes();
        PdfImagePlacement placement = Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
        // Crop keeps62.5% of each source axis. Mirroring changes the signed
        // source-pixel transform while the96x48pt anchor frame stays authored.
        Assert.Equal(-PictureWidth / .625D, placement.A, 3);
        Assert.Equal(0D, placement.B, 3);
        Assert.Equal(0D, placement.C, 3);
        Assert.Equal(-PictureHeight / .625D, placement.D, 3);
        Assert.NotNull(placement.Clip);
        using var pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        Assert.Single(page.GetImages());
        AssertFlowTokens(page.GetWords());
        AssertNoPictureWrapFallback(result);
        Assert.Equal(sourceXml, SourceXml(document));
    }

    private static WordDocument CreateDocument(bool footer, WordImageTextWrapping wrapping, out WordImage image, bool table = false) {
        WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        document.AddHeadersAndFooters();
        WordSection section = document.Sections[0];
        section.PageSettings.Width = 12240;
        section.PageSettings.Height = 15840;
        section.Margins.Left = section.Margins.Right = 1440;
        section.Margins.Top = section.Margins.Bottom = 1920;
        section.Margins.HeaderDistance = section.Margins.FooterDistance = 480;
        section.Header.Default!._header!.RemoveAllChildren();
        section.Footer.Default!._footer!.RemoveAllChildren();
        WordHeaderFooter story = footer ? section.Footer.Default! : section.Header.Default!;
        WordParagraph paragraph = table
            ? document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0]
            : story.AddParagraph();
        ConfigureParagraph(paragraph, string.Join(" ", FlowTokens));
        using var png = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(4, 2));
        image = paragraph.AddText(string.Empty).InsertImage(png, "wrapped.png", PictureWidth / .75D, PictureHeight / .75D, wrapping);
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Column;
        image.HorizontalPositionOffset = (long)((PictureLeft - 72D) * 12700D);
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Paragraph;
        image.VerticalPositionOffset = 0L;
        DW.Anchor anchor = image._Image.Anchor!;
        anchor.DistanceFromLeft = (uint)(DistanceLeft * 12700D);
        anchor.DistanceFromRight = (uint)(DistanceRight * 12700D);
        anchor.DistanceFromTop = (uint)(DistanceTop * 12700D);
        anchor.DistanceFromBottom = (uint)(DistanceBottom * 12700D);
        anchor.BehindDoc = footer;
        foreach (DW.EffectExtent extent in anchor.Descendants<DW.EffectExtent>()) {
            extent.LeftEdge = extent.TopEdge = extent.RightEdge = extent.BottomEdge = 0L;
        }
        ConfigureParagraph(document.AddParagraph(), "BODY");
        return document;
    }

    private static void ConfigureParagraph(WordParagraph paragraph, string text) {
        paragraph.Text = text;
        paragraph.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        paragraph.FontSize = 12;
        paragraph.LineSpacingPoints = 15;
        paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
        paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
    }

    private static WordPdfExportOptions Options() {
        var pdf = new PdfCreateOptions();
        pdf.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, Enumerable.Range(32, 95).ToArray())));
        return new WordPdfExportOptions {
            IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = pdf
        };
    }

    private static PdfWord[] AssertFlowTokens(IEnumerable<PdfWord> words) {
        PdfWord[] all = words.ToArray();
        PdfWord[] flow = FlowTokens.Select(token => Assert.Single(all, word => word.Text == token)).ToArray();
        return flow;
    }

    private static void AssertNoPictureWrapFallback(PdfDocumentConversionResult result) =>
        Assert.DoesNotContain(result.Warnings, warning => warning.Code is
            "native-anchored-image-flowed" or "native-anchored-image-wrap-approximated" or "header-footer-page-bounds-clipped");

    private static string SourceXml(WordDocument document) {
        var main = document.OpenXmlDocument.MainDocumentPart!;
        return main.Document.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml +
            main.DocumentSettingsPart!.Settings!.OuterXml +
            string.Concat(main.HeaderParts.Select(part => part.Header.OuterXml)) +
            string.Concat(main.FooterParts.Select(part => part.Footer.OuterXml));
    }
}
