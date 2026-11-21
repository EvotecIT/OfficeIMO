using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherGroupLineCanvasTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Grouped_shapes_retain_square_caps_and_blurred_miter_shadows(bool shadow) {
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(PublisherGroupCanvasTests.WidePaint(false, shadow).Pages[0].CreateDrawing());
        OfficeRasterImage grouped = OfficeDrawingRasterRenderer.Render(PublisherGroupCanvasTests.WidePaint(true, shadow).Pages[0].CreateDrawing());
        // With a shadow, the probe lies in the blurred tail above the mitered apex.
        int x = shadow ? 121 : 40, y = shadow ? 92 : 101;
        OfficeColor expected = reference.GetPixel(x, y);
        Assert.True(expected.R < 255);
        OfficeColor actual = grouped.GetPixel(x, y);
        // The grouped layer resamples the blurred shadow, so allow one level of rounding.
        Assert.InRange(actual.R - expected.R, -1, 1);
        Assert.InRange(actual.G - expected.G, -1, 1);
        Assert.InRange(actual.B - expected.B, -1, 1);
        Assert.Equal(expected.A, actual.A);
    }
}
