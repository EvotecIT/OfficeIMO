using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfUnderstandingPipelineTests {
    [Fact]
    public void FastPipeline_ExposesAllStagesAndCallerOrderedPages() {
        byte[] pdf = PdfDocument.Create()
            .H1("Pipeline heading")
            .Paragraph(p => p.Text("First body line"))
            .PageBreak()
            .Paragraph(p => p.Text("Second page body"))
            .ToBytes();

        PdfUnderstandingResult result = new PdfUnderstandingPipeline().Run(
            PdfReadDocument.Open(pdf),
            PdfPageSelection.From(2, 1));

        Assert.Equal(new[] { 2, 1 }, result.Pages.Select(static page => page.PageNumber));
        Assert.All(result.Pages, page => {
            Assert.NotEmpty(page.DecodedRuns);
            Assert.NotEmpty(page.Words);
            Assert.NotEmpty(page.Lines);
            Assert.NotEmpty(page.Regions);
            Assert.NotEmpty(page.ReadingOrder);
            Assert.NotEmpty(page.Elements);
            Assert.All(page.Words, word => { Assert.InRange(word.Confidence, 0D, 1D); Assert.NotEmpty(word.Evidence); });
            Assert.All(page.Lines, line => { Assert.InRange(line.Confidence, 0D, 1D); Assert.NotEmpty(line.Evidence); });
            Assert.All(page.Regions, region => { Assert.InRange(region.Confidence, 0D, 1D); Assert.NotEmpty(region.Evidence); });
            Assert.All(page.ReadingOrderEvidence, order => { Assert.InRange(order.Confidence, 0D, 1D); Assert.NotEmpty(order.Evidence); });
            Assert.All(page.Elements, element => { Assert.InRange(element.Confidence, 0D, 1D); Assert.NotEmpty(element.Evidence); });
            Assert.Equal(new[] { "glyph-decoding", "word-grouping", "line-grouping", "page-segmentation", "reading-order", "semantic-classification" }, page.Trace.Select(static trace => trace.Stage));
        });
        Assert.Contains(result.Pages[1].Elements, static element => element.Kind == PdfUnderstandingSemanticKind.Heading);
    }

    [Fact]
    public void Pipeline_UsesCallerSuppliedStageAndRecordsItsProvider() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(p => p.Text("Top region"))
            .Paragraph(p => p.Text("Bottom region"), style: new PdfParagraphStyle { SpacingBefore = 40 })
            .ToBytes();
        var custom = new ReverseReadingOrderStage();
        var options = new PdfUnderstandingPipelineOptions { ReadingOrder = custom };

        PdfUnderstandingPageResult page = Assert.Single(PdfDocument.Open(pdf).Read.Understand(options).Pages);

        Assert.Equal(typeof(ReverseReadingOrderStage), Assert.Single(page.Trace, static trace => trace.Stage == "reading-order").ProviderType);
        Assert.Equal(page.Regions.Reverse().Select(static region => region.Text), page.ReadingOrder.Select(static region => region.Text));
    }

    [Fact]
    public void Pipeline_RejectsOversizedCustomStageArtifacts() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var options = new PdfUnderstandingPipelineOptions {
            GlyphDecoding = new FixedGlyphStage(new[] {
                new PdfTextSpan("one", "F1", 12, 10, 10, 20),
                new PdfTextSpan("two", "F1", 12, 40, 10, 20)
            }),
            MaxRunsPerPage = 1
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            new PdfUnderstandingPipeline(options).Run(PdfReadDocument.Open(pdf)));

        Assert.Equal(PdfReadLimitKind.UnderstandingArtifacts, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }

    [Fact]
    public void AdvancedPipeline_GroupsRotatedBaselinesAndOrdersMultipleColumns() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Left top", "F1", 12, 50, 700, 48),
            new PdfTextSpan("Left bottom", "F1", 12, 50, 650, 60),
            new PdfTextSpan("Right top", "F1", 12, 300, 700, 54),
            new PdfTextSpan("Right bottom", "F1", 12, 300, 650, 66),
            new PdfTextSpan("Vertical one", "F1", 12, 520, 300, 66, rotationDegrees: 90),
            new PdfTextSpan("Vertical two", "F1", 12, 520, 370, 66, rotationDegrees: 90)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Advanced();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(new PdfUnderstandingPipeline(options).Run(PdfReadDocument.Open(pdf)).Pages);

        Assert.Contains(page.Lines, line => Math.Abs(line.RotationDegrees - 90D) < 0.1D && line.Text.Contains("Vertical one Vertical two", StringComparison.Ordinal));
        string[] horizontalOrder = page.ReadingOrder.Select(region => region.Text).Where(text => text.StartsWith("Left", StringComparison.Ordinal) || text.StartsWith("Right", StringComparison.Ordinal)).ToArray();
        Assert.Equal(new[] { "Left top", "Left bottom", "Right top", "Right bottom" }, horizontalOrder);
        Assert.Equal(typeof(PdfAdvancedUnderstandingStages).Assembly, Assert.Single(page.Trace, trace => trace.Stage == "reading-order").ProviderType.Assembly);
    }

    [Fact]
    public void AdvancedPipeline_UsesRecursiveWhitespaceCutsForIndentedColumnContent() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Left one", "F1", 12, 50, 700, 110),
            new PdfTextSpan("Left two", "F1", 12, 80, 650, 110),
            new PdfTextSpan("Left three", "F1", 12, 50, 600, 110),
            new PdfTextSpan("Right one", "F1", 12, 320, 700, 110),
            new PdfTextSpan("Right two", "F1", 12, 320, 650, 110),
            new PdfTextSpan("Right three", "F1", 12, 320, 600, 110)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Advanced();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(new PdfUnderstandingPipeline(options).Run(PdfReadDocument.Open(pdf)).Pages);

        Assert.Equal(new[] {
            "Left one", "Left two", "Left three",
            "Right one", "Right two", "Right three"
        }, page.ReadingOrder.Select(static region => region.Text));
    }

    [Fact]
    public void AdvancedPipeline_ForceSingleColumnOrdersRowsBeforeColumns() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Left top", "F1", 12, 50, 700, 90),
            new PdfTextSpan("Right top", "F1", 12, 320, 700, 90),
            new PdfTextSpan("Left bottom", "F1", 12, 50, 640, 90),
            new PdfTextSpan("Right bottom", "F1", 12, 320, 640, 90)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Advanced(new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(new PdfUnderstandingPipeline(options).Run(PdfReadDocument.Open(pdf)).Pages);

        Assert.Equal(new[] { "Left top", "Right top", "Left bottom", "Right bottom" },
            page.ReadingOrder.Select(static region => region.Text));
    }

    [Fact]
    public void AdvancedPipeline_OrdersStaggeredRegionsTopToBottomInsteadOfAsColumns() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Upper right heading", "F1", 16, 320, 700, 150),
            new PdfTextSpan("Lower left body", "F1", 11, 50, 500, 120)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Advanced();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(new PdfUnderstandingPipeline(options).Run(PdfReadDocument.Open(pdf)).Pages);

        Assert.Equal(new[] { "Upper right heading", "Lower left body" },
            page.ReadingOrder.Select(static region => region.Text));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AdvancedPipeline_OrdersInterleavedStaggeredRegionsTopToBottom(bool mirrorColumns) {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        double outerColumn = mirrorColumns ? 320 : 50;
        double middleColumn = mirrorColumns ? 50 : 320;
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Upper outer", "F1", 12, outerColumn, 700, 110),
            new PdfTextSpan("Middle inner", "F1", 12, middleColumn, 600, 110),
            new PdfTextSpan("Lower outer", "F1", 12, outerColumn, 500, 110)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Advanced();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(new PdfUnderstandingPipeline(options).Run(PdfReadDocument.Open(pdf)).Pages);

        Assert.Equal(new[] { "Upper outer", "Middle inner", "Lower outer" },
            page.ReadingOrder.Select(static region => region.Text));
    }

    [Fact]
    public void AdvancedPipeline_ClassifiesTablesCaptionsHeadersAndFootnotes() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Quarterly report", "F1", 12, 50, 800, 90),
            new PdfTextSpan("Item", "F1", 11, 50, 500, 24), new PdfTextSpan("Amount", "F1", 11, 90, 500, 42),
            new PdfTextSpan("Licenses", "F1", 11, 50, 482, 24), new PdfTextSpan("42", "F1", 11, 90, 482, 12),
            new PdfTextSpan("Figure 1. Revenue by region", "F1", 10, 50, 400, 150),
            new PdfTextSpan("1 Audited values exclude pending adjustments.", "F1", 8, 50, 20, 190)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Advanced();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(new PdfUnderstandingPipeline(options).Run(PdfReadDocument.Open(pdf)).Pages);

        Assert.Contains(page.Elements, element => element.Kind == PdfUnderstandingSemanticKind.Header);
        Assert.Contains(page.Elements, element => element.Kind == PdfUnderstandingSemanticKind.Table);
        Assert.Contains(page.Elements, element => element.Kind == PdfUnderstandingSemanticKind.Caption);
        Assert.Contains(page.Elements, element => element.Kind == PdfUnderstandingSemanticKind.Footnote);
    }

    [Fact]
    public void AdvancedPipeline_UsesCanonicalLayoutTableRegions() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Table metric", "Value" },
                new[] { "Quality", "Premium" }
            })
            .ToBytes();

        PdfUnderstandingPageResult page = Assert.Single(
            new PdfUnderstandingPipeline(PdfUnderstandingPipelineOptions.Advanced())
                .Run(PdfReadDocument.Open(pdf))
                .Pages);

        PdfUnderstandingSemanticElement table = Assert.Single(
            page.Elements,
            static element => element.Kind == PdfUnderstandingSemanticKind.Table);
        Assert.Contains("Table metric", table.Region.Text, StringComparison.Ordinal);
        Assert.Contains("Premium", table.Region.Text, StringComparison.Ordinal);
        Assert.Contains(table.Region.Evidence, static evidence => evidence.Code == "region.canonical-table");
        Assert.DoesNotContain(
            page.Elements,
            static element => element.Kind == PdfUnderstandingSemanticKind.Paragraph &&
                              element.Region.Text.Contains("Premium", StringComparison.Ordinal));
    }

    [Fact]
    public void AdvancedPipeline_UsesCanonicalRegularFontProseTableRegions() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Assigned owner", "Current workflow" },
                new[] { "North region coordinator", "Review pending requests" },
                new[] { "South region coordinator", "Approve completed requests" }
            }, style: new PdfTableStyle {
                HeaderBold = false,
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 220D, 220D }
            })
            .ToBytes();

        PdfUnderstandingPageResult page = Assert.Single(
            new PdfUnderstandingPipeline(PdfUnderstandingPipelineOptions.Advanced())
                .Run(PdfReadDocument.Open(pdf))
                .Pages);

        PdfUnderstandingSemanticElement table = Assert.Single(
            page.Elements,
            static element => element.Kind == PdfUnderstandingSemanticKind.Table);
        Assert.Contains("North region coordinator", table.Region.Text, StringComparison.Ordinal);
        Assert.Contains("Approve completed requests", table.Region.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void AdvancedPipeline_PrioritizesCanonicalTablesAtPageEdges() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Advanced();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("Item", "Helvetica-Bold", 11D, 50D, 820D, 40D),
            new PdfTextSpan("Amount", "Helvetica-Bold", 11D, 220D, 820D, 55D),
            new PdfTextSpan("Licenses", "Helvetica", 11D, 50D, 802D, 55D),
            new PdfTextSpan("42", "Helvetica", 11D, 220D, 802D, 16D)
        });

        PdfUnderstandingPageResult page = Assert.Single(
            new PdfUnderstandingPipeline(options).Run(PdfReadDocument.Open(pdf)).Pages);

        Assert.Equal(
            PdfUnderstandingSemanticKind.Table,
            Assert.Single(page.Elements, static element => element.Region.Text.Contains("Licenses", StringComparison.Ordinal)).Kind);
    }

    [Fact]
    public void AdvancedPipeline_TableOwnershipRequiresMeaningfulHorizontalOverlap() {
        PdfUnderstandingLine clipped = CreateUnderstandingLine("Adjacent prose", 249.5D, 339.5D, 475D);
        PdfUnderstandingLine owned = CreateUnderstandingLine("Table prose", 100D, 190D, 475D);

        Assert.False(PdfAdvancedUnderstandingStages.HasMeaningfulTableOverlap(clipped, 500D, 450D, 50D, 250D));
        Assert.True(PdfAdvancedUnderstandingStages.HasMeaningfulTableOverlap(owned, 500D, 450D, 50D, 250D));
    }

    [Fact]
    public void AdvancedReadingOrder_UsesRotatedSourceRunExtents() {
        var sourceRun = new PdfTextSpan("Vertical section label", "Helvetica", 11D, 280D, 200D, 400D, rotationDegrees: 90D);
        var word = new PdfUnderstandingWord(
            sourceRun.Text,
            280D,
            280D,
            200D,
            11D,
            90D,
            new[] { sourceRun });
        var region = new PdfUnderstandingRegion(new[] { new PdfUnderstandingLine(new[] { word }) });

        (double left, double right, double bottom, double top, _) = PdfRecursiveXyCutReadingOrderStage.GetSourceBounds(region);

        Assert.True(top - bottom >= 399D, $"Expected the rotated extent to span about 400 points, but it spanned {top - bottom:0.###}.");
        Assert.True(right - left >= 10D, $"Expected the rotated glyph thickness to span about one font size, but it spanned {right - left:0.###}.");
    }

    [Fact]
    public void AdvancedReadingOrder_LimitsSourceRunExtentToEachWordSegment() {
        var sourceRun = new PdfTextSpan("Left Right", "Helvetica", 11D, 50D, 500D, 500D);
        var leftRegion = new PdfUnderstandingRegion(new[] {
            new PdfUnderstandingLine(new[] {
                new PdfUnderstandingWord("Left", 50D, 250D, 500D, 11D, 0D, new[] { sourceRun })
            })
        });
        var rightRegion = new PdfUnderstandingRegion(new[] {
            new PdfUnderstandingLine(new[] {
                new PdfUnderstandingWord("Right", 350D, 550D, 500D, 11D, 0D, new[] { sourceRun })
            })
        });

        var leftBounds = PdfRecursiveXyCutReadingOrderStage.GetSourceBounds(leftRegion);
        var rightBounds = PdfRecursiveXyCutReadingOrderStage.GetSourceBounds(rightRegion);

        Assert.True(leftBounds.Right < rightBounds.Left, $"Expected the split word segments to retain their gutter, but bounds overlap at {leftBounds.Right:0.###} and {rightBounds.Left:0.###}.");
    }

    [Fact]
    public void AdvancedReadingOrder_IsolatesCloseSpanningHeadingBeforeColumns() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        PdfReadPage sourcePage = PdfReadDocument.Open(pdf).Pages[0];
        var context = new PdfUnderstandingPageContext(sourcePage, 1, new PdfTextLayoutOptions(), 10000, 10000);
        PdfUnderstandingRegion heading = new(new[] { CreateUnderstandingLine("Spanning heading", 50D, 500D, 710D) });
        PdfUnderstandingRegion leftTop = new(new[] { CreateUnderstandingLine("Left top", 50D, 160D, 700D) });
        PdfUnderstandingRegion leftBottom = new(new[] { CreateUnderstandingLine("Left bottom", 50D, 160D, 650D) });
        PdfUnderstandingRegion rightTop = new(new[] { CreateUnderstandingLine("Right top", 320D, 430D, 700D) });
        PdfUnderstandingRegion rightBottom = new(new[] { CreateUnderstandingLine("Right bottom", 320D, 430D, 650D) });

        IReadOnlyList<PdfUnderstandingRegion> ordered = new PdfRecursiveXyCutReadingOrderStage().Order(
            context,
            new[] { heading, leftTop, rightTop, leftBottom, rightBottom });

        Assert.Equal(
            new[] { "Spanning heading", "Left top", "Left bottom", "Right top", "Right bottom" },
            ordered.Select(static region => region.Text));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Pipeline_DoesNotClassifyDecimalAmountsAsListItems(bool advanced) {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("1037.25", "F1", 11, 50, 500, 45),
            new PdfTextSpan("1. Actual numbered item", "F1", 11, 50, 430, 120),
            new PdfTextSpan("-42", "F1", 11, 50, 360, 24)
        });
        PdfUnderstandingPipelineOptions options = advanced
            ? PdfUnderstandingPipelineOptions.Advanced()
            : new PdfUnderstandingPipelineOptions();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(new PdfUnderstandingPipeline(options).Run(PdfReadDocument.Open(pdf)).Pages);

        Assert.Equal(PdfUnderstandingSemanticKind.Paragraph,
            Assert.Single(page.Elements, element => element.Region.Text == "1037.25").Kind);
        Assert.Equal(PdfUnderstandingSemanticKind.ListItem,
            Assert.Single(page.Elements, element => element.Region.Text == "1. Actual numbered item").Kind);
        Assert.Equal(PdfUnderstandingSemanticKind.Paragraph,
            Assert.Single(page.Elements, element => element.Region.Text == "-42").Kind);
    }

    [Theory]
    [InlineData(false, "1.2. Nested numbered item")]
    [InlineData(true, "1.2. Nested numbered item")]
    [InlineData(false, "2.3.1)Deep numbered item")]
    [InlineData(true, "2.3.1)Deep numbered item")]
    [InlineData(false, "3.Compact numbered item")]
    [InlineData(true, "3.Compact numbered item")]
    [InlineData(false, "(a)Compact parenthesized item")]
    [InlineData(true, "(a)Compact parenthesized item")]
    [InlineData(false, "(1)Compact numeric parenthesized item")]
    [InlineData(true, "(1)Compact numeric parenthesized item")]
    [InlineData(false, "-Compact ASCII bullet")]
    [InlineData(true, "-Compact ASCII bullet")]
    [InlineData(false, "*Compact ASCII bullet")]
    [InlineData(true, "*Compact ASCII bullet")]
    public void Pipeline_ClassifiesHierarchicalAndCompactListItems(bool advanced, string text) {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        PdfUnderstandingPipelineOptions options = advanced
            ? PdfUnderstandingPipelineOptions.Advanced()
            : new PdfUnderstandingPipelineOptions();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan(text, "F1", 11, 50, 500, 180)
        });

        PdfUnderstandingPageResult page = Assert.Single(new PdfUnderstandingPipeline(options).Run(PdfReadDocument.Open(pdf)).Pages);

        Assert.Equal(PdfUnderstandingSemanticKind.ListItem,
            Assert.Single(page.Elements, element => element.Region.Text == text).Kind);
    }

    [Theory]
    [InlineData("-$42 total")]
    [InlineData("-€42 total")]
    [InlineData("-£ 42 total")]
    public void SharedListParser_DoesNotClassifySignedCurrencyAsCompactBullets(string text) {
        Assert.False(ContentStructureExtractor.IsListItemText(text));
    }

    [Theory]
    [InlineData("-.5 variance")]
    [InlineData("-,25 margin")]
    [InlineData("-$.5 variance")]
    public void SharedListParser_DoesNotClassifyLeadingDecimalValuesAsCompactBullets(string text) {
        Assert.False(ContentStructureExtractor.IsListItemText(text));
    }

    [Theory]
    [InlineData("--output path")]
    [InlineData("**bold**")]
    [InlineData("-*literal")]
    public void SharedListParser_DoesNotClassifyRepeatedPunctuationAsCompactBullets(string text) {
        Assert.False(ContentStructureExtractor.IsListItemText(text));
    }

    private sealed class ReverseReadingOrderStage : IPdfReadingOrderStage {
        public IReadOnlyList<PdfUnderstandingRegion> Order(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> regions) => regions.Reverse().ToArray();
    }

    private static PdfUnderstandingLine CreateUnderstandingLine(
        string text,
        double xStart,
        double xEnd,
        double baselineY) {
        var run = new PdfTextSpan(text, "Helvetica", 11D, xStart, baselineY, xEnd - xStart);
        var word = new PdfUnderstandingWord(text, xStart, xEnd, baselineY, 11D, 0D, new[] { run });
        return new PdfUnderstandingLine(new[] { word });
    }

    private sealed class FixedGlyphStage : IPdfGlyphDecodingStage {
        private readonly IReadOnlyList<PdfTextSpan> _spans;
        internal FixedGlyphStage(IReadOnlyList<PdfTextSpan> spans) { _spans = spans; }
        public IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context) => _spans;
    }
}
