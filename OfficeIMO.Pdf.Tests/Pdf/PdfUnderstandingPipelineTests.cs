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

    private sealed class ReverseReadingOrderStage : IPdfReadingOrderStage {
        public IReadOnlyList<PdfUnderstandingRegion> Order(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> regions) => regions.Reverse().ToArray();
    }

    private sealed class FixedGlyphStage : IPdfGlyphDecodingStage {
        private readonly IReadOnlyList<PdfTextSpan> _spans;
        internal FixedGlyphStage(IReadOnlyList<PdfTextSpan> spans) { _spans = spans; }
        public IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context) => _spans;
    }
}
