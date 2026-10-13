using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherGroupLineCanvasTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Grouped_shapes_retain_square_caps_and_blurred_miter_shadows(bool shadow) {
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(PublisherGroupCanvasTests.WidePaint(false, shadow).Pages[0].Drawing);
        OfficeRasterImage grouped = OfficeDrawingRasterRenderer.Render(PublisherGroupCanvasTests.WidePaint(true, shadow).Pages[0].Drawing);
        int x = shadow ? 121 : 40, y = shadow ? 61 : 101;
        OfficeColor expected = reference.GetPixel(x, y);
        Assert.True(expected.R < 255);
        Assert.Equal(expected, grouped.GetPixel(x, y));
    }
}
