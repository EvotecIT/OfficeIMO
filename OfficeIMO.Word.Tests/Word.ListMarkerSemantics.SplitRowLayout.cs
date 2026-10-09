using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class WordListMarkerSemanticsTests {
    private static double GetSplitRowTokenX(OfficeDrawingRichText text, string token) {
        OfficeRichTextLine line = Assert.Single(OfficeTextLayoutEngine.LayoutStyledRichTextBlock(
            text.Runs, text.Width - text.Padding.Horizontal, text.Height - text.Padding.Vertical,
            1.25D, MeasureRichTextWidth, wrap: false, paragraphIndent: text.ParagraphIndent).Lines);
        return text.X + text.Padding.Left + GetRichTextTokenBounds(line, token).X;
    }
}
