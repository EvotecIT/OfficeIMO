using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingRichTextParagraphIndentTests {
    [Fact]
    public void OriginalPublicConstructorRemainsAvailableToCompiledConsumers() {
        Assert.NotNull(typeof(OfficeRichTextRun).GetConstructor(new[] {
            typeof(string), typeof(double), typeof(OfficeColor), typeof(bool), typeof(bool), typeof(bool),
            typeof(string), typeof(bool), typeof(OfficeColor?), typeof(OfficeTextDecorationStyle),
            typeof(OfficeTextDecorationStyle), typeof(OfficeTextBaseline)
        }));
    }

    [Fact]
    public void IndentationOnTheRunAfterAHardBreakAppliesToThatParagraph() {
        var runs = new[] {
            new OfficeRichTextRun("Plain\n", 12D, OfficeColor.Black),
            new OfficeRichTextRun("Indented content", 12D, OfficeColor.Black)
                .WithParagraphIndent(new OfficeTextParagraphIndent(24D, 36D))
        };
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(runs,
            90D, 120D, 1.2D, (value, size, _) => (value?.Length ?? 0) * size * 0.5D, wrap: true);

        Assert.Equal(0D, layout.Lines[0].OffsetX);
        Assert.Equal(24D, layout.Lines[1].OffsetX);
        Assert.Equal(36D, layout.Lines[2].OffsetX);
    }

    [Fact]
    public void UnwrappedShrinkToFitIncludesAndScalesParagraphIndentation() {
        var runs = new[] {
            new OfficeRichTextRun("123456789", 10D, OfficeColor.Black)
                .WithParagraphIndent(OfficeTextParagraphIndent.FirstLine(20D))
        };
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(runs,
            100D, 30D, 1.2D, (value, size, _) => (value?.Length ?? 0) * size,
            wrap: false, shrinkToFit: true, minimumFontSize: 1D, overflowBehavior: OfficeTextOverflowBehavior.Clip);

        OfficeRichTextLine line = Assert.Single(layout.Lines);
        Assert.InRange(line.OffsetX, 18D, 18.3D);
        Assert.InRange(Assert.Single(line.Segments).FontSize, 9D, 9.2D);
        Assert.InRange(layout.Width, 99.9D, 100.01D);
        Assert.False(layout.Clipped);
    }

    [Fact]
    public void MixedParagraphIndentsSurviveDrawingCloneAndLayout() {
        var runs = new[] {
            new OfficeRichTextRun("Plain\n", 12D, OfficeColor.Black),
            new OfficeRichTextRun("\n", 12D, OfficeColor.Black).WithParagraphIndent(new OfficeTextParagraphIndent(24D, 36D)),
            new OfficeRichTextRun("Marker wrapped content\n", 12D, OfficeColor.Black),
            new OfficeRichTextRun("\n", 12D, OfficeColor.Black).WithParagraphIndent(new OfficeTextParagraphIndent(48D, 48D)),
            new OfficeRichTextRun("Markerless", 12D, OfficeColor.Black)
        };
        OfficeDrawingRichText text = new OfficeDrawingRichText(runs, 0D, 0D, 180D, 140D).Clone();
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(text.Runs,
            110D, 140D, 1.2D, (value, size, _) => (value?.Length ?? 0) * size * 0.5D, wrap: true);

        Assert.Contains(layout.Lines, line => line.OffsetX == 0D &&
            line.Segments.Any(segment => segment.Text.Contains("Plain", StringComparison.Ordinal)));
        Assert.Contains(layout.Lines, line => line.OffsetX == 24D &&
            line.Segments.Any(segment => segment.Text.Contains("Marker", StringComparison.Ordinal)));
        Assert.Contains(layout.Lines, line => line.OffsetX == 48D &&
            line.Segments.Any(segment => segment.Text.Contains("Markerless", StringComparison.Ordinal)));
    }
}
