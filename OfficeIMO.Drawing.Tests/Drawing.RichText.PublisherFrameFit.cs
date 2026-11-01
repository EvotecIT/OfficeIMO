using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingPublisherFrameFitTests {
    private static OfficeRasterCanvas Canvas() => new OfficeRasterCanvas(new OfficeRasterImage(1, 1),
        fonts: new OfficeFontFaceCollection().Add("Frame fixture",
            ManagedTextShapingTestAssets.CreatePublisherMetricsFont("dlig", "latn", false)));

    private static OfficeRichTextParagraph Paragraph(string text, double after = 0) => new OfficeRichTextParagraph(
        new[] { new OfficeRichTextRun(text, 10, OfficeColor.Black, fontFamily: "Frame fixture") },
        lineHeight: 18, margins: new OfficeTextPadding(0, 0, 0, after)) { UsesPublisherFontMetrics = true };

    [Theory]
    [InlineData(6D, 4.8D)]
    [InlineData(10D, 7.92D)]
    [InlineData(18D, 15.96D)]
    public void ExactSpacingUsesSelectedDeviceStrutAndFullAllocatedHeight(double spacing, double expectedBaseline) {
        var paragraph = new OfficeRichTextParagraph(Paragraph("A\nB\nA").Runs, lineHeight: spacing) { UsesPublisherFontMetrics = true };
        var canvas = Canvas();
        var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 100, 100, canvas.CreateDrawingTextMetrics());
        Assert.Equal(3, layout.Lines.Count); Assert.Equal(spacing * 3, layout.Height, 8);
        Assert.All(layout.Lines, line => {
            Assert.Equal(spacing, line.LineHeight, 8);
            Assert.Equal(expectedBaseline, OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, 0, spacing, true), 8);
        });
        canvas.FontMetricScale = 2;
        var scaled = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 200, 200,
            canvas.CreateDrawingTextMetrics(), scale: 2);
        Assert.Equal(layout.Height * 2, scaled.Height, 8);
        Assert.Equal(expectedBaseline * 2, OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(scaled.Lines[0], 0, spacing * 2, true), 8);
    }

    [Fact]
    public void NativeFirstLineIsForcedOncePerRegionAndAfterSpacingStillPrecedesFollowingText() {
        var source = new[] { Paragraph("A", 99), Paragraph("B"), Paragraph("A") };
        var metrics = Canvas().CreateDrawingTextMetrics();
        var flow = new OfficeRichTextFlow(source, _ => { });
        var first = Assert.Single(flow.Take(100, 1, metrics, default, allowFirstLineOverflow: true));
        Assert.Equal("A", first.Runs[0].Text); Assert.Equal(0, first.Margins.Bottom);
        Assert.Equal(2, flow.CharacterPosition); Assert.True(flow.HasRemaining);
        var second = Assert.Single(flow.Take(100, 35, metrics, default, allowFirstLineOverflow: true));
        Assert.Equal("B", second.Runs[0].Text); Assert.Equal(4, flow.CharacterPosition); Assert.True(flow.HasRemaining);
        Assert.Equal("A", Assert.Single(flow.Take(100, 18, metrics, default, allowFirstLineOverflow: true)).Runs[0].Text);
        Assert.False(flow.HasRemaining); Assert.Equal(5, flow.CharacterPosition);
        Assert.Equal(99, source[0].Margins.Bottom);
    }

    [Fact]
    public void StrictFlowDoesNotForceAnOverheightNativeStyleLine() {
        var flow = new OfficeRichTextFlow(new[] { Paragraph("A") }, _ => { });
        Assert.Empty(flow.Take(100, 1, Canvas().CreateDrawingTextMetrics(), default));
        Assert.True(flow.HasRemaining); Assert.Equal(0, flow.CharacterPosition);
    }

    [Theory]
    [InlineData(35D, 1)]
    [InlineData(36D, 2)]
    public void NativeWrappedLinesRequireTheirFullExactHeightAfterTheFirst(double height, int expectedLines) {
        var metrics = Canvas().CreateDrawingTextMetrics();
        var flow = new OfficeRichTextFlow(new[] { Paragraph("A\nB\nA") }, _ => { });
        var visible = flow.Take(100, height, metrics, default, allowFirstLineOverflow: true);
        var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(visible, 100, height, metrics);
        Assert.Equal(expectedLines, layout.Lines.Count);
        Assert.Equal(expectedLines * 18D, layout.Height, 8);
        Assert.Equal(expectedLines * 2, flow.CharacterPosition); Assert.True(flow.HasRemaining);
    }

    [Fact]
    public void DrawingPlacementSnapshotsAndTintRetainNativeFirstLinePlacement() {
        var drawing = new OfficeDrawing(100, 50).AddRichTextParagraphs(new[] { Paragraph("A") }, 0, 0, 100, 1);
        ((OfficeDrawingRichText)drawing.Elements[0]).NativePublisherFrame = true;
        var placed = new OfficeDrawing(120, 70).AddDrawing(drawing, 10, 10);
        var copied = placed.Clone();
        copied.ApplyColorTint(OfficeColor.Red);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(copied.Elements));
        var layout = OfficeDrawingTextLayout.CreateWithRasterMetrics(text, 100, 1, Canvas());
        Assert.Single(layout.Lines); Assert.Equal(18D, layout.Height, 8);
        Assert.Equal(15.96D, OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(layout.Lines[0], 0, 18, true), 8);
        Assert.Equal("A", text.PlainText);
    }
}
