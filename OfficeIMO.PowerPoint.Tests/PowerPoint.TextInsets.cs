using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public class PowerPointTextInsetTests {
    [Theory]
    [InlineData(false, 7.2D, 3.6D)]
    [InlineData(true, 0D, 0D)]
    public void ImageTextFramesUseDrawingMlDefaultsAndPreserveExplicitZero(bool explicitZero, double horizontal, double vertical) {
        using var presentation = PowerPointDocument.Create();
        var slide = presentation.AddSlide();
        var topLeft = slide.AddTextBoxPoints("Top left", 18, 18, 220, 60);
        var bottomRight = slide.AddTextBoxPoints("Bottom right", 18, 100, 220, 60);
        bottomRight.Paragraphs[0].Alignment = PowerPointTextAlignment.Right;
        bottomRight.TextVerticalAlignment = PowerPointTextVerticalAlignment.Bottom;
        var boxes = new[] { topLeft, bottomRight };
        foreach (var box in boxes) {
            box.FontSize = 12;
            box.FontName = "Helvetica";
            if (explicitZero) {
                box.TextMarginLeftPoints = box.TextMarginRightPoints = 0D;
                box.TextMarginTopPoints = box.TextMarginBottomPoints = 0D;
            }
        }
        var inset = slide.CreateVisualSnapshot();
        foreach (var box in boxes) {
            box.TextMarginLeftPoints = box.TextMarginRightPoints = 0D;
            box.TextMarginTopPoints = box.TextMarginBottomPoints = 0D;
        }
        var zero = slide.CreateVisualSnapshot();
        foreach (var box in boxes) {
            var actual = Assert.Single(inset.TextSpans, span => span.Text == box.Text);
            var reference = Assert.Single(zero.TextSpans, span => span.Text == box.Text);
            bool fromRight = ReferenceEquals(box, bottomRight);
            Assert.InRange(Math.Abs((fromRight ? actual.X + actual.Advance : actual.X)
                - (fromRight ? 238D - horizontal : 18D + horizontal)), 0D, 0.002D);
            Assert.InRange(Math.Abs(actual.X - reference.X - (fromRight ? -horizontal : horizontal)), 0D, 0.002D);
            // PDF user space starts at the bottom: top and bottom insets move baselines in opposite directions.
            Assert.InRange(Math.Abs(actual.Y - reference.Y - (fromRight ? vertical : -vertical)), 0D, 0.002D);
        }
    }
}
