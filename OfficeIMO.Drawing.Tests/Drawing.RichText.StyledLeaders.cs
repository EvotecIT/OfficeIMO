using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextStyledLeaderTests {
    private static double Measure(string? text, double size, string? font, OfficeFontStyle style) => (text?.Length ?? 0) * size / 2;
    private static double Start(OfficeRichTextLine line, string text) => line.OffsetX + line.Segments.TakeWhile(s => s.Text != text).Sum(s => s.Width);
    private static OfficeRichTextParagraph Paragraph(OfficeTextTabLeaderStyle style) => new OfficeRichTextParagraph(new[] {
        new OfficeRichTextRun("A\t", 10, OfficeColor.Red, bold: true), new OfficeRichTextRun("B\tC", 20, OfficeColor.Green)
    }).WithTabStops(new OfficeTextTabStops(new[] {
        new OfficeTextTabStop(100).WithLeader(".").WithLeaderStyle(style), new OfficeTextTabStop(200).WithLeader(".").WithLeaderStyle(style)
    }));
    private static OfficeRichTextSegment[] Glyphs(OfficeRichTextBlockLayout layout) => layout.Lines.SelectMany(l => l.Segments)
        .Where(s => s.Text.Length > 0 && s.Text.All(c => c == '.')).ToArray();

    [Fact]
    public void PartialStylesInheritEachActiveRunWithoutChangingFieldAnchorsOrBodyFonts() {
        var paragraph = Paragraph(new OfficeTextTabLeaderStyle(color: OfficeColor.Blue));
        var layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 400, 100, Measure);
        var glyphs = Glyphs(layout); Assert.Equal(2, glyphs.Length);
        Assert.Equal(new[] { 10D, 20D }, glyphs.Select(s => s.FontSize));
        Assert.True(glyphs[0].Bold); Assert.False(glyphs[1].Bold); Assert.All(glyphs, s => Assert.Equal(OfficeColor.Blue, s.Color));
        var line = Assert.Single(layout.Lines); Assert.Equal(100, Start(line, "B")); Assert.Equal(200, Start(line, "C"));
        Assert.Equal(new[] { 10D, 20D }, line.Segments.Where(s => s.Text is "A" or "B").Select(s => s.FontSize));
        var drawing = new OfficeDrawing(400, 100).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 400, 100);
        Assert.Equal("A\tB\tC", Assert.IsType<OfficeDrawingRichText>(drawing.Elements[0]).PlainText);
    }

    [Theory]
    [InlineData(false, 1D)]
    [InlineData(false, 2D)]
    [InlineData(true, 1D)]
    [InlineData(true, 2D)]
    public void AbsoluteAndRelativeSizesFollowDrawingScaleAndLeaveStopsFixed(bool absolute, double scale) {
        var style = absolute ? new OfficeTextTabLeaderStyle(fontSizePoints: 15) : new OfficeTextTabLeaderStyle(fontSizeFactor: 1.5);
        var layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { Paragraph(style) }, 600, 200, Measure, scale: scale);
        Assert.Equal(new[] { 15 * scale, (absolute ? 15 : 30) * scale }, Glyphs(layout).Select(s => s.FontSize));
        Assert.Equal(100 * scale, Start(layout.Lines[0], "B")); Assert.Equal(200 * scale, Start(layout.Lines[0], "C"));
        Assert.Equal(absolute ? 15D : (double?)null, style.FontSizePoints);
    }

    [Fact]
    public void AbsoluteLeaderSizesFitTogetherWithBodyFontsWithoutMovingTheStop() {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB", 10, OfficeColor.Red) })
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(100).WithLeader(".")
                .WithLeaderStyle(new OfficeTextTabLeaderStyle(fontSizePoints: 20)) }));
        var layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 400, 12, Measure, shrinkToFit: true);
        var glyph = Assert.Single(Glyphs(layout)); var body = Assert.Single(layout.Lines[0].Segments, s => s.Text == "B");
        Assert.False(layout.Clipped); Assert.InRange(body.FontSize, 1, 9.999); Assert.Equal(body.FontSize * 2, glyph.FontSize, 6);
        Assert.Equal(100, Start(layout.Lines[0], "B"));
    }

    [Fact]
    public void DecorativeTextBudgetAndUnrepresentableRelativeSizesDoNotShrinkOrAbortBodyLayout() {
        foreach (double factor in new[] { .000001D, double.MaxValue }) {
            var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB", 10, OfficeColor.Red) })
                .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(100).WithLeader(".")
                    .WithLeaderStyle(new OfficeTextTabLeaderStyle(fontSizeFactor: factor)) }));
            var layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 400, 100, Measure, shrinkToFit: true);
            Assert.True(layout.Clipped); Assert.Equal(100, Start(layout.Lines[0], "B"));
            Assert.Equal(10, Assert.Single(layout.Lines[0].Segments, s => s.Text == "B").FontSize);
            Assert.InRange(Glyphs(layout).Sum(s => s.Text.Length), 0, OfficeTextLayoutEngine.MaximumLayoutTextCharacters);
        }
    }

    [Fact]
    public void CloningTintingAndIndependentEditsRetainStyleAndAlphaWithoutMutatingTheSource() {
        var style = new OfficeTextTabLeaderStyle(color: OfficeColor.FromRgba(0, 0, 255, 128), bold: false,
            italic: true, underlineStyle: OfficeTextDecorationStyle.Double, strikethroughStyle: OfficeTextDecorationStyle.Dashed,
            backgroundColor: OfficeColor.FromRgba(255, 255, 0, 64));
        var paragraph = Paragraph(style); var stop = paragraph.TabStops!.Stops[0];
        Assert.Same(style, stop.WithLeader("_").WithLineLeader(new OfficeTextTabLineLeader()).LeaderStyle);
        Assert.Equal(".", stop.WithLeaderStyle(null).LeaderText); Assert.Null(stop.WithLeaderStyle(null).LeaderStyle);
        var drawing = new OfficeDrawing(400, 100).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 400, 100);
        var clone = drawing.Clone(); clone.ApplyColorTint(OfficeColor.Green);
        var copy = Assert.IsType<OfficeDrawingRichText>(clone.Elements[0]).Paragraphs[0].TabStops!.Stops[0].LeaderStyle!;
        Assert.Equal(OfficeColor.FromRgba(0, 128, 0, 128), copy.Color); Assert.Equal(OfficeColor.FromRgba(0, 128, 0, 64), copy.BackgroundColor);
        var paint = Glyphs(OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 400, 100, Measure));
        Assert.All(paint, s => { Assert.False(s.Bold); Assert.True(s.Italic); Assert.Equal(OfficeTextDecorationStyle.Double, s.UnderlineStyle); });
        Assert.Equal(OfficeColor.FromRgba(0, 0, 255, 128), style.Color);
    }

    [Fact]
    public void InvalidStyleValuesRejectBeforeLayoutAndTransparentBackgroundClearsInheritedPaint() {
        foreach (double value in new[] { 0D, -1D, double.NaN, double.PositiveInfinity }) {
            Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeTextTabLeaderStyle(fontSizePoints: value));
            Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeTextTabLeaderStyle(fontSizeFactor: value));
        }
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeTextTabLeaderStyle(strikethroughStyle: OfficeTextDecorationStyle.Words));
        var run = new OfficeRichTextRun("A\tB", 10, OfficeColor.Red, backgroundColor: OfficeColor.Yellow);
        var paragraph = new OfficeRichTextParagraph(new[] { run }).WithTabStops(new OfficeTextTabStops(new[] {
            new OfficeTextTabStop(100).WithLeader(".").WithLeaderStyle(new OfficeTextTabLeaderStyle(backgroundColor: OfficeColor.Transparent))
        }));
        var glyph = Assert.Single(Glyphs(OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 400, 100, Measure)));
        Assert.Equal(OfficeColor.Transparent, glyph.BackgroundColor); Assert.Equal(OfficeColor.Yellow, run.BackgroundColor);
    }

    [Theory]
    [InlineData(false, null, 64)]
    [InlineData(true, null, 128)]
    [InlineData(true, .25D, 64)]
    [InlineData(false, 0D, 0)]
    public void OpacityOverridesAndInheritanceHaveExplicitPrecedence(bool inherit, double? opacity, byte alpha) {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB", 10, OfficeColor.FromRgba(255, 0, 0, 128)) })
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(100).WithLeader(".").WithLeaderStyle(
                new OfficeTextTabLeaderStyle(color: OfficeColor.FromRgba(0, 0, 255, 64), inheritOpacity: inherit, opacity: opacity)) }));
        var glyph = Assert.Single(Glyphs(OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 400, 100, Measure)));
        Assert.Equal(OfficeColor.FromRgba(0, 0, 255, alpha), glyph.Color);
    }
}
