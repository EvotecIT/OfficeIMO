using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReadingFrameRegressionTests {
    [Theory]
    [InlineData(90D)]
    [InlineData(180D)]
    [InlineData(270D)]
    public void RestoredSearchableWordAndLineRetainExactSelectionBounds(double angle) {
        var run = new PdfTextSpan("Selection", "F1", 12D, 200D, 300D, 60D, null, false,
            angle, null, null, textRenderingMode: 3, hasActualText: true);
        PdfReadDocument source = PdfReadDocument.Open(PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes());
        var options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphs(new[] { run });
        PdfUnderstandingPageResult analysis = new PdfUnderstandingPipeline(new PdfTextLayoutOptions(), options)
            .RunPages(source, new[] { 1 })[0];
        double radians = angle * Math.PI / 180D;
        double endX = 200D + Math.Cos(radians) * 60D, endY = 300D + Math.Sin(radians) * 60D;
        double upX = -Math.Sin(radians) * 12D, upY = Math.Cos(radians) * 12D;
        PdfVisualBounds expected = source.Pages[0].TransformBoundsToVisual(
            Math.Min(Math.Min(200D, endX), Math.Min(200D + upX, endX + upX)),
            Math.Min(Math.Min(300D, endY), Math.Min(300D + upY, endY + upY)),
            Math.Max(Math.Max(200D, endX), Math.Max(200D + upX, endX + upX)),
            Math.Max(Math.Max(300D, endY), Math.Max(300D + upY, endY + upY)));
        PdfUnderstandingWord word = Assert.Single(analysis.Words);
        AssertBounds(expected, word.VisualBounds);
        AssertBounds(expected, Assert.Single(analysis.Lines).VisualBounds);
        Assert.True(word.IsSelectionBox);
    }

    [Fact]
    public void CorrectedFrameDoesNotMakeClippedTextEligibleForCompleteCellRecovery() {
        PdfReadPage page = PdfReadDocument.Open(PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes()).Pages[0];
        var context = new PdfUnderstandingPageContext(page, 1, new PdfTextLayoutOptions(), 1000, 100);
        Assert.True(PdfPageClipPathBuilder.TryCreateTransformedRectangle(new Matrix2D(1, 0, 0, 1, 0, 0),
            395D, 40D, 10D, 15D, context.Height, OfficeFillRule.NonZero, out PdfPageClipPath clip));
        var source = new PdfTextSpan("Partially clipped", "F1", 10D, 395D, 40D, 60D, null, true, 90D, null, clip);
        Assert.False(source.CanProjectCompleteText(context.Height));
        PdfUnderstandingReadingFrame frame = Assert.IsType<PdfUnderstandingReadingFrame>(PdfUnderstandingReadingFrame.TryCreate(context, new[] { source }));
        PdfTextSpan projected = Assert.Single(frame.ProjectRuns(new[] { source }));
        Assert.False(projected.CanProjectCompleteText(frame.Height));
    }

    [Theory]
    [InlineData(PdfReadingDirection.LeftToRight, "שמאל ימין")]
    [InlineData(PdfReadingDirection.RightToLeft, "ימין שמאל")]
    [InlineData(PdfReadingDirection.Auto, "ימין שמאל")]
    public void NativeTableCellCompositionHonorsRequestedDirection(PdfReadingDirection direction, string expected) {
        var spans = new List<PdfTextSpan> {
            new PdfTextSpan("שמאל", "F1", 10D, 50D, 700D, 25D),
            new PdfTextSpan("ימין", "F1", 10D, 80D, 700D, 25D),
            new PdfTextSpan("24", "F1", 10D, 200D, 700D, 15D)
        };
        TextLayoutEngine.TextLine line = TextLayoutEngine.BuildLine(spans, new TextLayoutEngine.Options { ReadingDirection = direction });
        Assert.Equal(new[] { expected, "24" }, TableDetector.SplitBySplits(line, new List<double> { 150D }));
    }

    private static void AssertBounds(PdfVisualBounds expected, PdfLogicalVisualBounds? actual) {
        Assert.NotNull(actual);
        Assert.Equal(expected.Left, actual!.Left, 6);
        Assert.Equal(expected.Top, actual.Top, 6);
        Assert.Equal(expected.Right, actual.Right, 6);
        Assert.Equal(expected.Bottom, actual.Bottom, 6);
    }

    private sealed class FixedGlyphs : IPdfGlyphDecodingStage {
        private readonly PdfTextSpan[] _runs;
        internal FixedGlyphs(PdfTextSpan[] runs) => _runs = runs;
        public IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context) => _runs;
    }
}
