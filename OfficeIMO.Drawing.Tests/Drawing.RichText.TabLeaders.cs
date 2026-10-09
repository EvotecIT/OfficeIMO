using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextTabLeaderTests {
    private static double Measure(string? text, double size, string? family, OfficeFontStyle style) => (text?.Length ?? 0) * size / 2;
    private static OfficeRichTextParagraph Paragraph(string text, params OfficeTextTabStop[] stops) =>
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, 10, OfficeColor.Black) })
            .WithTabStops(new OfficeTextTabStops(stops));
    private static OfficeRichTextBlockLayout Layout(OfficeRichTextParagraph paragraph,
        Func<string?, double, string?, OfficeFontStyle, double>? measure = null) =>
        OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 200, 200, measure ?? Measure);
    private static double Start(OfficeRichTextLine line, string text) => line.OffsetX + line.Segments.TakeWhile(s => s.Text != text).Sum(s => s.Width);

    [Theory]
    [InlineData(OfficeTextTabAlignment.Left, 50)]
    [InlineData(OfficeTextTabAlignment.Center, 35)]
    [InlineData(OfficeTextTabAlignment.Right, 20)]
    [InlineData(OfficeTextTabAlignment.Character, 35)]
    public void TextualPaintFillsTheGapWithoutChangingFieldAlignmentOrLogicalText(OfficeTextTabAlignment alignment, double start) {
        var paragraph = Paragraph("A\t123.45", new OfficeTextTabStop(50, alignment).WithLeader("."));
        var line = Assert.Single(Layout(paragraph).Lines);
        Assert.Equal(start, Start(line, "123.45"));
        Assert.Equal((start - 5) / 5, line.Segments.Single(s => s.Text.StartsWith(".", StringComparison.Ordinal)).Text.Length);
        var drawing = new OfficeDrawing(200, 200).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 200, 200);
        Assert.Equal("A\t123.45", ((OfficeDrawingRichText)drawing.Elements[0]).PlainText);
    }

    [Fact]
    public void PaintUsesTheTabsStyleAndAnOverlapOrDefaultStopAddsNoLeader() {
        var paragraph = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("AA", 10, OfficeColor.Black), new OfficeRichTextRun("\t", 10, OfficeColor.Red, bold: true),
            new OfficeRichTextRun("1234\tB\tC", 10, OfficeColor.Blue)
        }).WithTabStops(new OfficeTextTabStops(new[] {
            new OfficeTextTabStop(20, OfficeTextTabAlignment.Right).WithLeader("-"), new OfficeTextTabStop(60).WithLeader("_")
        }));
        var line = Assert.Single(Layout(paragraph).Lines);
        Assert.DoesNotContain(line.Segments, s => s.Text.Contains('-'));
        Assert.Equal(10, Start(line, "1234")); Assert.Equal(60, Start(line, "B")); Assert.Equal(72, Start(line, "C"));
        Assert.Equal(OfficeColor.Blue, line.Segments.Single(s => s.Text.Contains('_')).Color);
        var red = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB", 10, OfficeColor.Red, bold: true) })
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(50).WithLeader("-") }));
        var paint = Assert.Single(Layout(red).Lines).Segments.Single(s => s.Text.Contains('-'));
        Assert.Equal(OfficeColor.Red, paint.Color); Assert.True(paint.Bold);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GeneratedPaintBudgetSpansTabsAndParagraphsWhileSpacingAndBodyRemain(bool fit) {
        var paragraph = Paragraph("\tB\tC", new OfficeTextTabStop(300000).WithLeader("."), new OfficeTextTabStop(600000).WithLeader("."));
        var layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph, paragraph }, 700000, 200,
            (text, size, family, style) => text?.Length ?? 0, shrinkToFit: fit, minimumFontSize: 1);
        Assert.True(layout.Clipped); Assert.Equal(2, layout.Lines.Count);
        Assert.Equal(100000, layout.Lines.SelectMany(l => l.Segments).Where(s => s.Text.StartsWith(".", StringComparison.Ordinal)).Sum(s => s.Text.Length));
        Assert.All(layout.Lines, line => {
            Assert.Equal(300000, Start(line, "B")); Assert.Equal(600000, Start(line, "C"));
            Assert.All(line.Segments, segment => Assert.Equal(10, segment.FontSize));
        });
    }

    [Fact]
    public void WholeRunShapingCannotPaintBeyondTheFollowingField() {
        var paragraph = Paragraph("A\tB", new OfficeTextTabStop(50).WithLeader("."));
        var line = Assert.Single(Layout(paragraph, (text, size, family, style) =>
            text?.StartsWith(".", StringComparison.Ordinal) == true && text.Length > 1 ? text.Length * 8 : Measure(text, size, family, style)).Lines);
        var paint = line.Segments.Single(s => s.Text.StartsWith(".", StringComparison.Ordinal));
        Assert.InRange(paint.Width, 0, 45); Assert.Equal(50, Start(line, "B"));
    }

    [Fact]
    public void ScalingAndFrameFittingRetainLeadersAtTheirMeasuredAnchors() {
        var paragraph = Paragraph("A\tB", new OfficeTextTabStop(50).WithLeader("."));
        var scaled = Assert.Single(OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 400, 200, Measure, scale: 2).Lines);
        Assert.Equal(100, Start(scaled, "B")); Assert.Contains(scaled.Segments, s => s.Text == ".........");
        var fitted = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph, paragraph }, 200, 12, Measure,
            shrinkToFit: true, minimumFontSize: 1);
        Assert.False(fitted.Clipped); Assert.Equal(2, fitted.Lines.Count);
        Assert.All(fitted.Lines, line => {
            Assert.Equal(50, Start(line, "B"));
            var paint = line.Segments.Single(s => s.Text.StartsWith(".", StringComparison.Ordinal));
            Assert.InRange(paint.FontSize, 1, 5); Assert.True(paint.Text.Length > 9);
        });
        Assert.Equal(10, paragraph.Runs[0].FontSize);
    }

    [Fact]
    public void CancellationRaisedByFontMeasurementStopsLeaderGeneration() {
        using var cancellation = new CancellationTokenSource();
        var paragraph = Paragraph("A\tB", new OfficeTextTabStop(50).WithLeader("."));
        Assert.Throws<OperationCanceledException>(() => OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 200, 200,
            (text, size, family, style) => { if (text == ".") cancellation.Cancel(); return Measure(text, size, family, style); },
            cancellationToken: cancellation.Token));
    }

    [Fact]
    public void ImmutableLeaderEditingSupportsAScalarAndRejectsInvalidOrControlText() {
        var stop = new OfficeTextTabStop(50, OfficeTextTabAlignment.Character, ",");
        var updated = stop.WithLeader("\U0001F7E2");
        Assert.Null(stop.LeaderText); Assert.Equal("\U0001F7E2", updated.LeaderText);
        Assert.Equal(",", updated.Character); Assert.Null(updated.WithLeader(null).LeaderText);
        foreach (string invalid in new[] { "", "..", "\t", "\uD800" }) Assert.Throws<ArgumentException>(() => stop.WithLeader(invalid));
        var spaced = Paragraph("A\tB", stop.WithLeader(" "));
        Assert.Equal("AB", string.Concat(Assert.Single(Layout(spaced).Lines).Segments.Select(s => s.Text)));
    }
}
