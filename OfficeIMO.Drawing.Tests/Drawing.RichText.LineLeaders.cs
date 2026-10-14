using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextLineLeaderTests {
    private static double Measure(string? value, double size, string? family, OfficeFontStyle style) => (value?.Length ?? 0) * size / 2;
    private static OfficeRichTextParagraph Paragraph(OfficeTextTabLineLeader leader, string text = "A\tB", double position = 100) =>
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, 10, OfficeColor.Red) })
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(position).WithLineLeader(leader) }));
    private static OfficeRichTextBlockLayout Layout(OfficeRichTextParagraph paragraph, double scale = 1) =>
        OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 400, 100, Measure, scale: scale);
    private static double Start(OfficeRichTextLine line, string text) => line.OffsetX + line.Segments.TakeWhile(s => s.Text != text).Sum(s => s.Width);
    private static OfficeTextTabLineLeaderPaint Paint(OfficeRichTextBlockLayout layout) =>
        Assert.Single(layout.Lines.SelectMany(l => l.Segments), s => s.TabLinePaint != null).TabLinePaint!;

    [Theory]
    [InlineData(OfficeTextTabLineLeaderStyle.Solid)]
    [InlineData(OfficeTextTabLineLeaderStyle.Dotted)]
    [InlineData(OfficeTextTabLineLeaderStyle.Dash)]
    [InlineData(OfficeTextTabLineLeaderStyle.LongDash)]
    [InlineData(OfficeTextTabLineLeaderStyle.DotDash)]
    [InlineData(OfficeTextTabLineLeaderStyle.DotDotDash)]
    [InlineData(OfficeTextTabLineLeaderStyle.Wave)]
    public void EveryPatternPaintsWithinTheGapWithoutAddingTextOrMovingTheField(OfficeTextTabLineLeaderStyle style) {
        var paragraph = Paragraph(new OfficeTextTabLineLeader(style, color: OfficeColor.Blue));
        var layout = Layout(paragraph); var paint = Paint(layout);
        Assert.Equal(100, Start(Assert.Single(layout.Lines), "B"));
        Assert.Equal("AB", string.Concat(layout.Lines.SelectMany(l => l.Segments).Select(s => s.Text)));
        Assert.Equal(OfficeColor.Blue, paint.Color);
        Assert.InRange(paint.Left, 0, 95); Assert.InRange(paint.Right, paint.Left, 95);
        var drawing = new OfficeDrawing(400, 100).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 400, 100);
        var svg = System.Xml.Linq.XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));
        var path = Assert.Single(svg.Descendants(), e => e.Name.LocalName == "path");
        Assert.Equal(OfficeColor.Blue, OfficeColor.Parse((string)path.Attribute("fill")!));
        Assert.Equal("AB", string.Concat(svg.Descendants().Where(e => e.Name.LocalName == "text").Select(e => e.Value)));
    }

    [Fact]
    public void DoublePatternsHaveTwoRowsAndFontColorAbsoluteWidthAndScaleStayIndependent() {
        var leader = new OfficeTextTabLineLeader(doubleLine: true, widthPoints: 2);
        var paint = Paint(Layout(Paragraph(leader)));
        Assert.Equal(2, paint.Contours.Count); Assert.Equal(OfficeColor.Red, paint.Color);
        Assert.Equal(6, paint.Bottom - paint.Top, 8);
        var scaled = Layout(Paragraph(leader), scale: 2);
        Assert.Equal(200, Start(scaled.Lines[0], "B"));
        Assert.Equal(12, Paint(scaled).Bottom - Paint(scaled).Top, 8);
        Assert.Equal(2, leader.WidthPoints);
        var relative = Paint(Layout(Paragraph(new OfficeTextTabLineLeader(widthFontFraction: .1)), scale: 2));
        Assert.Equal(2, relative.Bottom - relative.Top, 8);
    }

    [Fact]
    public void TextAndSpaceTakePrecedenceAndIndependentEditsRetainBothDeclarations() {
        var leader = new OfficeTextTabLineLeader(OfficeTextTabLineLeaderStyle.Wave);
        var stop = new OfficeTextTabStop(100).WithLineLeader(leader).WithLeader(" ");
        Assert.Same(leader, stop.LineLeader); Assert.Null(stop.WithLineLeader(null).LineLeader);
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB", 10, OfficeColor.Red) })
            .WithTabStops(new OfficeTextTabStops(new[] { stop }));
        Assert.DoesNotContain(Layout(paragraph).Lines[0].Segments, s => s.TabLinePaint != null);
        Assert.Same(leader, stop.WithLeader(null).LineLeader);
        var none = Layout(Paragraph(new OfficeTextTabLineLeader(OfficeTextTabLineLeaderStyle.None)));
        Assert.DoesNotContain(none.Lines[0].Segments, s => s.TabLinePaint != null);
        Assert.Equal(100, Start(none.Lines[0], "B"));
    }

    [Fact]
    public void OverlappingFieldAndDefaultStopsHaveNoLinePaint() {
        var paragraph = Paragraph(new OfficeTextTabLineLeader(), "AAA\t1234\tB", 20);
        paragraph = paragraph.WithTabStops(new OfficeTextTabStops(new[] {
            new OfficeTextTabStop(20, OfficeTextTabAlignment.Right).WithLineLeader(new OfficeTextTabLineLeader())
        }));
        var line = Assert.Single(Layout(paragraph).Lines);
        Assert.Equal(15, Start(line, "1234")); Assert.Equal(36, Start(line, "B"));
        Assert.DoesNotContain(line.Segments, s => s.TabLinePaint != null);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TinyWidthPaintIsBoundedAcrossTheFrameWithoutShrinkingBodyOrMovingFields(bool fit) {
        var stops = Enumerable.Range(1, 20).Select(i => new OfficeTextTabStop(i * 100)
            .WithLineLeader(new OfficeTextTabLineLeader(OfficeTextTabLineLeaderStyle.Dotted, widthPoints: .000001))).ToArray();
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(string.Concat(Enumerable.Repeat("\tB", 20)), 10, OfficeColor.Red) })
            .WithTabStops(new OfficeTextTabStops(stops));
        var layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 2200, 100, Measure, shrinkToFit: fit, minimumFontSize: 1);
        Assert.True(layout.Clipped);
        Assert.All(layout.Lines.SelectMany(l => l.Segments), s => Assert.Equal(10, s.FontSize));
        var paints = layout.Lines.SelectMany(l => l.Segments).Where(s => s.TabLinePaint != null).Select(s => s.TabLinePaint!).ToArray();
        Assert.InRange(paints.Sum(p => p.Contours.Sum(c => c.Count)), 1, 100000);
        Assert.All(paints, p => Assert.InRange(p.Contours.Sum(c => c.Count), 1, 8192));
        double cursor = 0; int field = 0;
        foreach (var segment in Assert.Single(layout.Lines).Segments) {
            if (segment.Text == "B") Assert.Equal(++field * 100, cursor);
            cursor += segment.Width;
        }
        Assert.Equal(20, field);
    }

    [Fact]
    public void InvalidWidthsFailBeforePlanningAndCancellationStopsGeometry() {
        foreach (double value in new[] { 0D, -1D, double.NaN, double.PositiveInfinity }) {
            Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeTextTabLineLeader(widthPoints: value));
            Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeTextTabLineLeader(widthFontFraction: value));
        }
        var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeTextTabLineLeaderLayout.Create(new OfficeTextTabLineLeader(),
            100, 10, OfficeColor.Black, new OfficeTextTabLeaderLayout.Budget(), cancellation.Token, out _));
    }

    [Fact]
    public void SceneCloneAndPatternTintRetainLineSettingsWithoutRecoloringTheSource() {
        var paragraph = Paragraph(new OfficeTextTabLineLeader(OfficeTextTabLineLeaderStyle.Wave, true, 2,
            color: OfficeColor.FromRgba(0, 0, 255, 128)));
        var drawing = new OfficeDrawing(400, 100).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 400, 100);
        var cloned = drawing.Clone(); cloned.ApplyColorTint(OfficeColor.Green);
        var copied = ((OfficeDrawingRichText)cloned.Elements[0]).Paragraphs[0].TabStops!.Stops[0];
        Assert.Equal(OfficeTextTabLineLeaderStyle.Wave, copied.LineLeader!.Style); Assert.True(copied.LineLeader.DoubleLine);
        Assert.Equal(2, copied.LineLeader.WidthPoints); Assert.Equal(100, copied.Position);
        Assert.Equal(OfficeColor.FromRgba(0, 128, 0, 128), copied.LineLeader.Color);
        Assert.Equal(OfficeColor.FromRgba(0, 0, 255, 128), paragraph.TabStops!.Stops[0].LineLeader!.Color);
    }

    [Theory]
    [InlineData(double.Epsilon, 1D)]
    [InlineData(double.Epsilon, .5D)]
    [InlineData(double.MaxValue, 2D)]
    public void UnpaintableWidthsFailClosedThroughScalingWithoutChangingTheBodyOrField(double width, double scale) {
        var leader = new OfficeTextTabLineLeader(OfficeTextTabLineLeaderStyle.Dotted, widthPoints: width);
        var layout = Layout(Paragraph(leader), scale);
        Assert.True(layout.Clipped); Assert.Equal(100 * scale, Start(layout.Lines[0], "B"));
        Assert.DoesNotContain(layout.Lines[0].Segments, s => s.TabLinePaint != null);
        Assert.All(layout.Lines[0].Segments, s => Assert.Equal(10 * scale, s.FontSize));
        Assert.Equal(width, leader.WidthPoints);
    }
}
