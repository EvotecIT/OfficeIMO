using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextTabParagraphAlignmentTests {
    private static double Measure(string? text, double size, string? family, OfficeFontStyle style) => (text?.Length ?? 0) * size / 2;
    private static OfficeRichTextBlockLayout Layout(OfficeRichTextParagraph paragraph) {
        var drawing = new OfficeDrawing(200, 100).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 200, 100);
        return OfficeDrawingTextLayout.Create((OfficeDrawingRichText)drawing.Elements[0], 200, 100, Measure);
    }
    private static double Start(OfficeRichTextLine line, string text) => line.OffsetX + line.Segments.TakeWhile(s => s.Text != text).Sum(s => s.Width);

    [Theory]
    [InlineData(OfficeTextAlignment.Center, OfficeTextTabAlignment.Left, "123.45", 80, 45)]
    [InlineData(OfficeTextAlignment.Center, OfficeTextTabAlignment.Center, "123.45", 65, 52.5)]
    [InlineData(OfficeTextAlignment.Center, OfficeTextTabAlignment.Right, "123.45", 50, 60)]
    [InlineData(OfficeTextAlignment.Center, OfficeTextTabAlignment.Character, "123.45", 65, 52.5)]
    [InlineData(OfficeTextAlignment.Center, OfficeTextTabAlignment.Character, "123456", 50, 60)]
    [InlineData(OfficeTextAlignment.Right, OfficeTextTabAlignment.Left, "123.45", 80, 90)]
    [InlineData(OfficeTextAlignment.Right, OfficeTextTabAlignment.Center, "123.45", 65, 105)]
    [InlineData(OfficeTextAlignment.Right, OfficeTextTabAlignment.Right, "123.45", 50, 120)]
    [InlineData(OfficeTextAlignment.Right, OfficeTextTabAlignment.Character, "123.45", 65, 105)]
    [InlineData(OfficeTextAlignment.Right, OfficeTextTabAlignment.Character, "123456", 50, 120)]
    public void OptInMovesTheWholeLineWithoutChangingTabFieldAlignment(OfficeTextAlignment alignment,
        OfficeTextTabAlignment tabAlignment, string value, double fieldStart, double shift) {
        var settings = new OfficeTextTabStops(new[] { new OfficeTextTabStop(80, tabAlignment) });
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\t" + value, 10, OfficeColor.Black) }, alignment);
        var fixedLine = Assert.Single(Layout(paragraph.WithTabStops(settings)).Lines);
        var alignedLine = Assert.Single(Layout(paragraph.WithTabStops(settings.WithParagraphAlignment())).Lines);
        Assert.Equal(0, fixedLine.OffsetX);
        Assert.Equal(fieldStart, Start(fixedLine, value));
        Assert.Equal(shift, Start(alignedLine, "A"));
        Assert.Equal(fieldStart + shift, Start(alignedLine, value));
        Assert.Equal(fixedLine.Width, alignedLine.Width);
        Assert.False(settings.AlignWithParagraph);
    }

    [Theory]
    [InlineData(OfficeTextAlignment.Center, 57.5, 87.5)]
    [InlineData(OfficeTextAlignment.Right, 115, 175)]
    public void HardBreakPlainLinesRetainAlignmentWithEitherTabGridPolicy(OfficeTextAlignment alignment,
        double tabbedOffset, double plainOffset) {
        var settings = new OfficeTextTabStops(new[] { new OfficeTextTabStop(80) });
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB\nplain\nC\tD", 10, OfficeColor.Black) }, alignment);
        var fixedLines = Layout(paragraph.WithTabStops(settings)).Lines;
        var alignedLines = Layout(paragraph.WithTabStops(settings.WithParagraphAlignment())).Lines;
        Assert.Equal(new[] { 0D, plainOffset, 0D }, fixedLines.Select(line => line.OffsetX));
        Assert.Equal(new[] { tabbedOffset, plainOffset, tabbedOffset }, alignedLines.Select(line => line.OffsetX));
        Assert.Equal(80 + tabbedOffset, Start(alignedLines[2], "D"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TabbedLinesNeverJustifyTheirSpaces(bool alignWithParagraph) {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tone two\nlast", 10, OfficeColor.Black) }, OfficeTextAlignment.Justify)
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(80) }).WithParagraphAlignment(alignWithParagraph));
        var layout = Layout(paragraph);
        Assert.Equal(2, layout.Lines.Count);
        Assert.Equal(0, layout.Lines[0].OffsetX);
        Assert.Equal(115, layout.Lines[0].Width);
        Assert.Equal(80, Start(layout.Lines[0], "one two"));
    }

    [Fact]
    public void ImmutablePolicyChangesAndSceneCloneTintKeepStopsAndRenderedAlignment() {
        var settings = new OfficeTextTabStops(new[] { new OfficeTextTabStop(80).WithLeader(".") }, defaultInterval: 50, origin: 5);
        var aligned = settings.WithParagraphAlignment();
        var restored = aligned.WithParagraphAlignment(false);
        Assert.NotSame(settings, aligned);
        Assert.NotSame(aligned, restored);
        Assert.False(settings.AlignWithParagraph);
        Assert.True(aligned.AlignWithParagraph);
        Assert.False(restored.AlignWithParagraph);
        Assert.Equal(50, aligned.DefaultInterval);
        Assert.Equal(5, aligned.Origin);
        Assert.Equal(".", Assert.Single(aligned.Stops).LeaderText);
        var scaled = aligned.Scale(2);
        Assert.True(scaled.AlignWithParagraph);
        Assert.Equal(100, scaled.DefaultInterval);
        Assert.Equal(10, scaled.Origin);
        Assert.Equal(160, Assert.Single(scaled.Stops).Position);
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB\tC", 12, OfficeColor.Red) }, OfficeTextAlignment.Center);
        OfficeDrawing Drawing(OfficeTextTabStops tabs) => new OfficeDrawing(200, 100)
            .AddRichTextParagraphs(new[] { paragraph.WithTabStops(tabs) }, 0, 0, 200, 100);
        var original = Drawing(aligned);
        var clone = original.Clone();
        clone.ApplyColorTint(OfficeColor.Blue);
        var clonedParagraph = Assert.Single(((OfficeDrawingRichText)Assert.Single(clone.Elements)).Paragraphs);
        Assert.True(clonedParagraph.TabStops!.AlignWithParagraph);
        Assert.Equal(50, clonedParagraph.TabStops.DefaultInterval);
        Assert.Equal(5, clonedParagraph.TabStops.Origin);
        Assert.Equal(".", Assert.Single(clonedParagraph.TabStops.Stops).LeaderText);
        Assert.Equal(OfficeColor.Red, paragraph.Runs[0].Color);
        var fixedSvg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(Drawing(restored)));
        var originalSvg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(original));
        var cloneSvg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(clone));
        double shift = SvgStart(originalSvg, "A") - SvgStart(fixedSvg, "A");
        Assert.True(shift > 0);
        Assert.Equal(85, SvgStart(fixedSvg, "B"));
        Assert.Equal(105, SvgStart(fixedSvg, "C"));
        foreach (string field in new[] { "A", "B", "C" }) {
            Assert.Equal(shift, SvgStart(originalSvg, field) - SvgStart(fixedSvg, field), 3);
            Assert.Equal(SvgStart(originalSvg, field), SvgStart(cloneSvg, field));
        }
    }

    [Theory]
    [InlineData(1D)]
    [InlineData(2D)]
    public void RasterScaleRetainsTheSameWholeLineShiftAsSvg(double scale) {
        var paragraph = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("A\t", 12, OfficeColor.Black, fontFamily: "Tab Proof"),
            new OfficeRichTextRun("B", 12, OfficeColor.Red, fontFamily: "Tab Proof")
        }, OfficeTextAlignment.Center);
        var settings = new OfficeTextTabStops(new[] { new OfficeTextTabStop(80) }, origin: 5);
        OfficeDrawing Drawing(OfficeTextTabStops tabs) {
            var drawing = new OfficeDrawing(200, 40).AddRichTextParagraphs(new[] { paragraph.WithTabStops(tabs) }, 0, 0, 200, 40);
            drawing.Fonts.Add("Tab Proof", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf")));
            return drawing;
        }
        var fixedDrawing = Drawing(settings);
        var alignedDrawing = Drawing(settings.WithParagraphAlignment());
        var fixedSvg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(fixedDrawing));
        var alignedSvg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(alignedDrawing));
        double expectedShift = (SvgStart(alignedSvg, "A") - SvgStart(fixedSvg, "A")) * scale;
        Assert.True(expectedShift > 0);
        var fixedImage = OfficeDrawingRasterRenderer.Render(fixedDrawing, scale, OfficeColor.White);
        var alignedImage = OfficeDrawingRasterRenderer.Render(alignedDrawing, scale, OfficeColor.White);
        foreach (bool red in new[] { false, true }) {
            double actualShift = FirstInkColumn(alignedImage, red) - FirstInkColumn(fixedImage, red);
            Assert.InRange(actualShift, expectedShift - 1, expectedShift + 1);
        }
    }

    private static double SvgStart(XDocument svg, string text) => double.Parse(
        Assert.Single(svg.Descendants(), element => element.Name.LocalName == "text" && element.Value == text).Attribute("x")!.Value,
        CultureInfo.InvariantCulture);

    private static int FirstInkColumn(OfficeRasterImage image, bool red) {
        int first = image.Width;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (red ? pixel.R > 200 && pixel.G < 100 && pixel.B < 100 : pixel.R < 100 && pixel.G < 100 && pixel.B < 100)
                    first = Math.Min(first, x);
            }
        }
        Assert.InRange(first, 0, image.Width - 1);
        return first;
    }
}
