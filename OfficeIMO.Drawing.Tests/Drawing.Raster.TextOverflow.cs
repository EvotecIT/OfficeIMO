using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingRasterTests {
    [Fact]
    public void OfficeRasterCanvas_ClipOverflowKeepsExactTextWithoutImplicitInset() {
        const string text = "Sales";
        const double fontSize = 20D;
        var measuredImage = new OfficeRasterImage(180, 40, OfficeColor.Transparent);
        var measuredCanvas = new OfficeRasterCanvas(measuredImage);
        double advance = measuredCanvas.MeasureText(text, fontSize, "Arial");

        var positioned = new OfficeDrawing(180D, 40D).AddPositionedText(
            text,
            3D,
            0D,
            advance,
            32D,
            new OfficeFontInfo("Arial", fontSize),
            OfficeColor.Black);
        var bounded = new OfficeDrawing(180D, 40D).AddText(
            text,
            0D,
            0D,
            advance + 6D,
            32D,
            new OfficeFontInfo("Arial", fontSize),
            OfficeColor.Black);

        AssertRasterImagesEqual(
            OfficeDrawingRasterRenderer.Render(bounded),
            OfficeDrawingRasterRenderer.Render(positioned));
    }

    [Fact]
    public void OfficeDrawing_ClippedPositionedTextRetainsSourceAdvance() {
        const string text = "Quarterly Operations Dashboard";
        var drawing = new OfficeDrawing(180D, 40D).AddClippedPositionedText(
            text,
            0D,
            0D,
            160D,
            32D,
            0D,
            0D,
            OfficeClipPath.Rectangle(120D, 32D),
            new OfficeFontInfo("Arial", 20D),
            OfficeColor.Black,
            textAdvanceWidth: 150D);

        OfficeDrawingGroup group = Assert.IsType<OfficeDrawingGroup>(Assert.Single(drawing.Elements));
        OfficeDrawingText positioned = Assert.IsType<OfficeDrawingText>(Assert.Single(group.Drawing.Elements));
        Assert.Equal(text, positioned.Text);
        Assert.Equal(OfficeTextOverflowBehavior.Clip, positioned.OverflowBehavior);
        Assert.Equal(150D, positioned.TextAdvanceWidth);

        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.True(rendered.GetPixels().Any(static value => value != 0));
    }
}
