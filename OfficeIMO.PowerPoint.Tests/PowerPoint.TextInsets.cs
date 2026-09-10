using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public class PowerPointTextInsetTests {
    [Theory]
    [InlineData(false, 7.2D, 3.6D)]
    [InlineData(true, 0D, 0D)]
    public void ImageTextFramesUseDrawingMlDefaultsAndPreserveExplicitZero(bool explicitZero, double horizontal, double vertical) {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var box = slide.AddTextBoxPoints("Insets", 18, 18, 220, 60);
        if (explicitZero) {
            box.TextMarginLeftPoints = box.TextMarginRightPoints = 0D;
            box.TextMarginTopPoints = box.TextMarginBottomPoints = 0D;
        }
        var text = Assert.Single(slide.CreateVisualSnapshot().Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal(horizontal, text.Padding.Left, 6);
        Assert.Equal(horizontal, text.Padding.Right, 6);
        Assert.Equal(vertical, text.Padding.Top, 6);
        Assert.Equal(vertical, text.Padding.Bottom, 6);
    }
}
