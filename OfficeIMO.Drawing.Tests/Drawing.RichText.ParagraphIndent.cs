using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingRichTextParagraphIndentTests {
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
