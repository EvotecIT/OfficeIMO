using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfInlineScriptExtractionTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PublicPointActualTextKeepsItsScriptBesideTheCompletePaint(bool split) {
        PdfReadPage page = CreateTransformedCanvas(1D, 1D, 0D, canvas => {
            canvas.ActualText("mass", 100D, 100D, paint => {
                if (split) paint.Text("m", 100D, 100D, 60D, 40D, fontSize: 18D, font: PdfStandardFont.Courier)
                    .Text("ass", 110.8D, 100D, 60D, 40D, fontSize: 18D, font: PdfStandardFont.Courier);
                else paint.Text("mass", 100D, 100D, 80D, 40D, fontSize: 18D, font: PdfStandardFont.Courier);
            });
            AddMassScript(canvas);
        });

        Assert.Equal("mass-2", page.ExtractText());
        Assert.Equal("mass-2", page.ExtractTextWithColumns(new PdfTextLayoutOptions { ForceSingleColumn = true }));
        // Selection/OCR still sees the first native paint, rather than a widened synthetic glyph.
        PdfTextSpan mass = Assert.Single(page.GetTextSpans(), span => span.Text == "mass");
        Assert.Equal(split ? 10.8D : 43.2D, mass.Advance, 6);
        Assert.Equal(18D, mass.FontSize);
        Assert.True(mass.IsVisible);
        Assert.False(mass.TryGetGlyphs(out _));
    }

    [Fact]
    public void PublicTransformedPointActualTextKeepsItsScriptAndPaintGeometry() {
        PdfReadPage page = CreateTransformedCanvas(2D, 3D, 30D, canvas => {
            canvas.ActualText("mass", 100D, 100D, paint => paint
                .Text("m", 100D, 100D, 43.2D, 40D, fontSize: 18D, font: PdfStandardFont.Courier)
                .Text("ass", 100D, 100D, 43.2D, 40D, fontSize: 18D, align: PdfAlign.Right, font: PdfStandardFont.Courier));
            AddMassScript(canvas);
        });
        PdfTextSpan mass = Assert.Single(page.GetTextSpans(), span => span.Text == "mass");
        PdfReadPage paintControl = CreateTransformedCanvas(2D, 3D, 30D, canvas => canvas
            .Text("m", 100D, 100D, 43.2D, 40D, fontSize: 18D, font: PdfStandardFont.Courier));
        PdfTextSpan firstPaint = Assert.Single(paintControl.GetTextSpans());

        Assert.Equal("mass-2", page.ExtractText());
        Assert.Equal("mass-2", page.ExtractTextWithColumns(new PdfTextLayoutOptions { ForceSingleColumn = true }));
        Assert.Equal(firstPaint.X, mass.X, 6);
        Assert.Equal(firstPaint.Y, mass.Y, 6);
        Assert.Equal(firstPaint.Advance, mass.Advance, 6);
        Assert.Equal(18D, mass.FontSize);
        Assert.Equal(firstPaint.RotationDegrees, mass.RotationDegrees, 6);
    }

    [Theory]
    [InlineData("multiline")]
    [InlineData("backward")]
    [InlineData("extra-paint")]
    [InlineData("cropped")]
    public void IncompatibleActualTextPaintDoesNotInventAnAdjacentScript(string incompatibility) {
        PdfReadPage page = CreateTransformedCanvas(1D, 1D, 0D, canvas => {
            canvas.ActualText("mass", 100D, 100D, paint => {
                paint.Text("m", 100D, 100D, 60D, 40D, fontSize: 18D, font: PdfStandardFont.Courier);
                if (incompatibility == "cropped") paint.Clip(110.8D, 100D, 10.8D, 40D, clipped =>
                    clipped.Text("ass", 110.8D, 100D, 60D, 40D, fontSize: 18D, font: PdfStandardFont.Courier));
                else paint.Text("ass", incompatibility == "backward" ? 90D : 110.8D,
                    incompatibility == "multiline" ? 130D : 100D,
                    60D, 40D, fontSize: 18D, font: PdfStandardFont.Courier);
                // Even a complete early paint match is discarded when later paint contradicts it.
                if (incompatibility == "extra-paint") paint.Text("X", 143.2D, 100D,
                    30D, 40D, fontSize: 18D, font: PdfStandardFont.Courier);
            });
            AddMassScript(canvas);
        });

        Assert.Equal("-2\nmass", page.ExtractText());
        Assert.Equal("-2\nmass", page.ExtractTextWithColumns(new PdfTextLayoutOptions { ForceSingleColumn = true }));
        Assert.Equal(10.8D, Assert.Single(page.GetTextSpans(), span => span.Text == "mass").Advance, 6);
    }

    private static void AddMassScript(PdfPageCanvas canvas) => canvas.Text(new[] {
        new PdfTextRun("-2", fontSize: 18D, font: PdfStandardFont.Courier, baseline: PdfTextBaseline.Superscript)
    }, 143.2D, 100D, 60D, 40D, fontSize: 18D);

    [Fact]
    public void PublicExtractionKeepsRaisedAndLoweredTextInTheirAuthoredLines() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
            DefaultFontSize = 18D,
            CompressContentStreams = false
        }).Compose(document => document.Content(content => content
            .Paragraph(paragraph => paragraph.Text("m").Superscript("-1").Text(" tail"))
            .Paragraph(paragraph => paragraph.Text("H").Subscript("2").Text("O"))))
            .ToBytes();
        PdfReadPage page = PdfReadDocument.Open(bytes).Pages[0];

        Assert.Equal("m-1 tail\nH2O", page.ExtractText());
        Assert.Equal("m-1 tail\nH2O", page.ExtractTextWithColumns(
            new PdfTextLayoutOptions { ForceSingleColumn = true }));
    }

    [Theory]
    [InlineData(4.77D)]
    [InlineData(-4.77D)]
    public void ContiguousMixedFontScriptRunKeepsBaseAndFollowingText(double rise) {
        var spans = new[] {
            new PdfTextSpan("m", "Captured", 12.72D, 50D, 700D, 10.545D,
                baseFont: "SourceSansProWeb-Regular"),
            new PdfTextSpan("\u2212", "Declared", 9.54D, 60.545D, 700D + rise, 4.742D,
                baseFont: "NISTDeclaredSourceSans-Regular"),
            new PdfTextSpan("1", "Captured", 9.54D, 65.287D, 700D + rise, 4.742D,
                baseFont: "SourceSansProWeb-Regular"),
            new PdfTextSpan(" tail", "Captured", 12.72D, 73D, 700D, 24D,
                baseFont: "SourceSansProWeb-Regular")
        };

        TextLayoutEngine.TextLine line = Assert.Single(TextLayoutEngine.BuildLines(spans));

        Assert.Equal("m\u22121 tail", line.Text);
        Assert.Equal(700D, line.Y);
        Assert.Equal(spans.Length, line.Spans.Count);
    }

    [Fact]
    public void RaisedPrefixKeepsItsBaseLineBelowAnUnrelatedRow() {
        var spans = new[] {
            new PdfTextSpan("2", "F1", 12D, 50D, 706D, 6D),
            new PdfTextSpan("H", "F1", 18D, 56D, 700D, 10D),
            new PdfTextSpan("above", "F1", 18D, 200D, 703D, 45D)
        };

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans);
        string text = TextLayoutEngine.EmitText(lines,
            new TextLayoutEngine.ColumnLayout((0D, 600D), (0D, 0D), false));

        Assert.Equal(2, lines.Count);
        Assert.Equal("above\n2H", text);
        Assert.Equal(700D, lines.Single(line => line.Text == "2H").Y);
    }

    [Theory]
    [InlineData(5.729577951308233D)]
    [InlineData(-5.729577951308233D)]
    [InlineData(80D)]
    [InlineData(-80D)]
    [InlineData(90D)]
    public void ScriptGroupingPreservesOriginalRotatedSpanGeometry(double angle) {
        PdfTextSpan[] spans = CreateScriptLine(angle);
        var geometry = spans.Select(span =>
            (span.X, span.Y, span.Advance, span.FontSize, span.RotationDegrees, span.FontResource)).ToArray();

        TextLayoutEngine.TextLine line = Assert.Single(TextLayoutEngine.BuildLines(spans));

        Assert.Equal("m-1H2O", line.Text);
        Assert.Equal(spans.Length, line.Spans.Count);
        Assert.All(spans, span => Assert.Contains(span, line.Spans));
        Assert.Equal(geometry, spans.Select(span =>
            (span.X, span.Y, span.Advance, span.FontSize, span.RotationDegrees, span.FontResource)).ToArray());
    }

    [Theory]
    [InlineData(60D, 678D, 12D)] // A real following row.
    [InlineData(300D, 706D, 12D)] // A raised label in a distant column.
    [InlineData(60D, 706D, 18D)] // Same-size text on a distinct baseline.
    [InlineData(50D, 706D, 12D)] // A separate label above, rather than beside, the base.
    public void SeparateRowsColumnsAndShiftedLabelsRemainSeparate(double x, double y, double fontSize) {
        var spans = new[] {
            new PdfTextSpan("m", "F1", 18D, 50D, 700D, 10D),
            new PdfTextSpan("1", "F1", fontSize, x, y, 6D)
        };

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans);

        Assert.Equal(2, lines.Count);
        Assert.Contains(lines, line => line.Text == "m");
        Assert.Contains(lines, line => line.Text == "1");
    }

    [Fact]
    public void ARunBetweenTwoEligibleBaseLinesRemainsUnattached() {
        var spans = new[] {
            new PdfTextSpan("U", "F1", 18D, 50D, 704D, 10D),
            new PdfTextSpan("-1", "F2", 12D, 60D, 700D, 6D),
            new PdfTextSpan("L", "F1", 18D, 66D, 696D, 10D)
        };

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans);

        Assert.Equal(new[] { "U", "-1", "L" }, lines.Select(line => line.Text));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnExplicitBreakAtTheAdjacentBoundaryPreventsAttachment(bool scriptPrecedesBase) {
        PdfTextSpan script = SpanWithBreak("1", 12D, scriptPrecedesBase ? 50D : 60D,
            706D, 6D, scriptPrecedesBase ? 0 : 2);
        PdfTextSpan normal = SpanWithBreak("m", 18D, scriptPrecedesBase ? 56D : 50D,
            700D, 10D, scriptPrecedesBase ? 2 : 0);

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(new[] { normal, script });

        Assert.Equal(2, lines.Count);
        Assert.Equal(2, lines.Single(line => line.Spans.Contains(
            scriptPrecedesBase ? normal : script)).LogicalLineBreaksBefore);
    }

    [Fact]
    public void AnEarlierBreakOnTheBaseLineDoesNotDisableItsLaterInlineScript() {
        var spans = new[] {
            SpanWithBreak("row ", 18D, 10D, 700D, 40D, 1),
            new PdfTextSpan("m", "F1", 18D, 50D, 700D, 10D),
            new PdfTextSpan("-1", "F2", 12D, 60D, 706D, 8D)
        };

        TextLayoutEngine.TextLine line = Assert.Single(TextLayoutEngine.BuildLines(spans));

        Assert.Equal("row m-1", line.Text);
        Assert.Equal(1, line.LogicalLineBreaksBefore);
    }

    [Fact]
    public void PublicDownscaledCanvasKeepsRealNeighboringRowsSeparate() {
        PdfReadPage page = CreateTransformedCanvas(0.25D, 0.25D, 0D, canvas => canvas
            .Text("m", 100D, 100D, 60D, 30D, fontSize: 18D, font: PdfStandardFont.Courier)
            .Text("row", 110.8D, 124D, 60D, 30D, fontSize: 12D, font: PdfStandardFont.Courier));
        IReadOnlyList<PdfTextSpan> spans = page.GetTextSpans();
        Assert.Contains(spans, span => span.FontSize == 18D && span.Text == "m");
        Assert.All(spans, span => Assert.True(span.TextToPageTransform.HasValue));

        Assert.Equal("m\nrow", page.ExtractText());
        Assert.Equal("m\nrow", page.ExtractTextWithColumns(
            new PdfTextLayoutOptions { ForceSingleColumn = true }));
    }

    [Theory]
    [InlineData(2D, 2D, 0D)]
    [InlineData(3D, 2D, 0D)]
    [InlineData(2D, 2D, 30D)]
    public void PublicTransformedCanvasKeepsAConnectedRaisedRunAdjacent(double scaleX, double scaleY, double angle) {
        PdfReadPage page = CreateTransformedCanvas(scaleX, scaleY, angle, canvas => canvas
            .Text("m", 100D, 100D, 60D, 40D, fontSize: 18D, font: PdfStandardFont.Courier)
            .Text(new[] { new PdfTextRun("-", fontSize: 18D, font: PdfStandardFont.Courier,
                baseline: PdfTextBaseline.Superscript) }, 110.8D, 100D, 60D, 40D, fontSize: 18D)
            .Text(new[] { new PdfTextRun("1", fontSize: 18D, font: PdfStandardFont.Courier,
                baseline: PdfTextBaseline.Superscript) }, 119.82D, 100D, 60D, 40D, fontSize: 18D));
        IReadOnlyList<PdfTextSpan> spans = page.GetTextSpans();
        var geometry = spans.Select(span =>
            (span.X, span.Y, span.Advance, span.FontSize, span.RotationDegrees, span.TextToPageTransform)).ToArray();
        Assert.All(spans, span => Assert.True(span.TextToPageTransform.HasValue));

        Assert.Equal("m-1", page.ExtractText());
        Assert.Equal("m-1", page.ExtractTextWithColumns(
            new PdfTextLayoutOptions { ForceSingleColumn = true }));
        Assert.Equal(geometry, spans.Select(span =>
            (span.X, span.Y, span.Advance, span.FontSize, span.RotationDegrees, span.TextToPageTransform)).ToArray());
    }

    [Theory]
    [InlineData(PdfReadingDirection.RightToLeft)]
    [InlineData(PdfReadingDirection.Auto)]
    public void RightToLeftSuffixDoesNotAttachAcrossItsOwnExplicitBreak(PdfReadingDirection direction) {
        PdfTextSpan normal = OrderedSpanWithBreak("\u05D0", 18D, 66D, 700D, 10D, 0, 0);
        PdfTextSpan script = OrderedSpanWithBreak("1", 12D, 60D, 706D, 6D, 2, 1);
        var spans = new[] { normal, script };
        var geometry = spans.Select(span => (span.X, span.Y, span.Advance, span.FontSize)).ToArray();

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans,
            new TextLayoutEngine.Options { ReadingDirection = direction });

        Assert.Equal(2, lines.Count);
        Assert.Equal(2, lines.Single(line => line.Spans.Contains(script)).LogicalLineBreaksBefore);
        Assert.Equal(new[] { normal, script }, spans);
        Assert.Equal(geometry, spans.Select(span => (span.X, span.Y, span.Advance, span.FontSize)).ToArray());
    }

    [Theory]
    [InlineData(PdfReadingDirection.RightToLeft)]
    [InlineData(PdfReadingDirection.Auto)]
    public void ExplicitBreakStartingARightToLeftBaseDoesNotBlockItsLaterSuffix(PdfReadingDirection direction) {
        PdfTextSpan normal = OrderedSpanWithBreak("\u05D0", 18D, 66D, 700D, 10D, 1, 0);
        PdfTextSpan script = OrderedSpanWithBreak("1", 12D, 60D, 706D, 6D, 0, 1);
        var spans = new[] { normal, script };
        var geometry = spans.Select(span => (span.X, span.Y, span.Advance, span.FontSize)).ToArray();

        TextLayoutEngine.TextLine line = Assert.Single(TextLayoutEngine.BuildLines(spans,
            new TextLayoutEngine.Options { ReadingDirection = direction }));

        Assert.Equal("\u05D01", line.Text);
        Assert.Equal(700D, line.Y);
        Assert.Equal(1, line.LogicalLineBreaksBefore);
        Assert.Equal(new[] { script, normal }, line.Spans);
        Assert.Equal(new[] { normal, script }, spans);
        Assert.Equal(geometry, spans.Select(span => (span.X, span.Y, span.Advance, span.FontSize)).ToArray());
    }

    [Theory]
    [InlineData(PdfReadingDirection.RightToLeft)]
    [InlineData(PdfReadingDirection.Auto)]
    public void RightToLeftScriptFragmentsDoNotBridgeAnInternalExplicitBreak(PdfReadingDirection direction) {
        PdfTextSpan normal = OrderedSpanWithBreak("\u05D0", 18D, 72D, 700D, 10D, 0, 0);
        PdfTextSpan firstSuffix = OrderedSpanWithBreak("1", 12D, 66D, 706D, 6D, 0, 1);
        PdfTextSpan nextLine = OrderedSpanWithBreak("2", 12D, 60D, 706D, 6D, 1, 2);
        var spans = new[] { normal, firstSuffix, nextLine };

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans,
            new TextLayoutEngine.Options { ReadingDirection = direction });

        Assert.Equal(2, lines.Count);
        Assert.Contains(lines, line => line.Text == "\u05D01");
        Assert.Equal(1, lines.Single(line => line.Text == "2").LogicalLineBreaksBefore);
        Assert.Equal(new[] { normal, firstSuffix, nextLine }, spans);
        Assert.All(spans, span => Assert.Contains(lines, line => line.Spans.Contains(span)));
    }

    [Theory]
    [InlineData(0.1D, 2.5D)]
    [InlineData(0.6D, 0.75D)]
    public void ExplicitlyStricterMergeLimitsKeepRaisedRunsSeparate(double em, double points) {
        var spans = new[] {
            new PdfTextSpan("m", "F1", 18D, 50D, 700D, 10D),
            new PdfTextSpan("1", "F1", 12D, 60D, 706D, 6D)
        };

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans,
            new TextLayoutEngine.Options { LineMergeToleranceEm = em, LineMergeMaxPoints = points });

        Assert.Equal(2, lines.Count);
    }

    [Fact]
    public void ScriptComparisonsReportWorkAndPropagateTheCallersBudgetStop() {
        PdfTextSpan[] spans = CreateScriptLine(0D);
        long charged = 0L;
        var stop = new InvalidOperationException("Layout work limit reached.");

        InvalidOperationException actual = Assert.Throws<InvalidOperationException>(() =>
            TextLayoutEngine.BuildLines(spans, options: null, consumeWork: units => {
                charged += units;
                if (charged > spans.Length) throw stop;
            }, cancellationCheck: null));

        Assert.Same(stop, actual);
        Assert.True(charged > spans.Length);
    }

    [Fact]
    public void ScriptComparisonsObserveCancellationAfterWorkHasStarted() {
        PdfTextSpan[] spans = CreateScriptLine(0D);
        using var cancellation = new CancellationTokenSource();
        long charged = 0L;

        Assert.Throws<OperationCanceledException>(() => TextLayoutEngine.BuildLines(spans,
            options: null, consumeWork: units => {
                charged += units;
                if (charged > spans.Length) cancellation.Cancel();
            }, cancellationCheck: cancellation.Token.ThrowIfCancellationRequested));

        Assert.True(charged > spans.Length);
    }

    private static PdfTextSpan[] CreateScriptLine(double angle, double baselineY = 700D) => new[] {
        RotatedSpan("m", 18D, 0D, 0D, 10D, angle, baselineY),
        RotatedSpan("-1", 11.7D, 10D, 6.3D, 8D, angle, baselineY),
        RotatedSpan("H", 18D, 18D, 0D, 10D, angle, baselineY),
        RotatedSpan("2", 11.7D, 28D, -3.24D, 6D, angle, baselineY),
        RotatedSpan("O", 18D, 34D, 0D, 10D, angle, baselineY)
    };

    private static PdfTextSpan RotatedSpan(string text, double fontSize, double along, double rise,
        double advance, double angle, double baselineY) {
        double radians = angle * Math.PI / 180D;
        return new PdfTextSpan(text, "F1", fontSize,
            50D + along * Math.Cos(radians) - rise * Math.Sin(radians),
            baselineY + along * Math.Sin(radians) + rise * Math.Cos(radians),
            advance, rotationDegrees: angle);
    }

    private static PdfTextSpan SpanWithBreak(string text, double fontSize, double x, double y,
        double advance, int lineBreaks) => new PdfTextSpan(text, "F1", fontSize, x, y, advance,
            color: null, isVisible: true, rotationDegrees: 0D, baseFont: null, clipPath: null,
            logicalLineBreaksBefore: lineBreaks);

    private static PdfReadPage CreateTransformedCanvas(double scaleX, double scaleY, double angle,
        Action<PdfPageCanvas> build) {
        OfficeTransform transform = OfficeTransform.Scale(scaleX, scaleY)
            .Then(OfficeTransform.RotateDegrees(angle, 100D * scaleX, 100D * scaleY));
        byte[] bytes = PdfDocument.Create(new PdfOptions {
            PageWidth = 600D, PageHeight = 800D, DefaultFont = PdfStandardFont.Courier,
            CompressContentStreams = false
        }).Compose(document => document.Content(content => content.Canvas(canvas =>
            canvas.Effect(transform, 1D, build)))).ToBytes();
        return PdfReadDocument.Open(bytes).Pages[0];
    }

    private static PdfTextSpan OrderedSpanWithBreak(string text, double fontSize, double x, double y,
        double advance, int lineBreaks, int sourceOrder) => new PdfTextSpan(text, "F1", fontSize, x, y, advance,
            color: null, isVisible: true, rotationDegrees: 0D, baseFont: null, clipPath: null,
            paintOrder: sourceOrder, logicalLineBreaksBefore: lineBreaks,
            contentOrderKey: PdfContentOrderKey.Root.Append(sourceOrder));
}
