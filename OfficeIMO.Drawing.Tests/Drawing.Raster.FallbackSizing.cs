using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingRasterTests {
    [Theory]
    [InlineData("A\u05D0", 'A', 0x05D0)]
    [InlineData("\u05D0\u05D1", 0x05D0, 0x05D1)]
    public void AuthoredRtlPlacesFallbackFacesInVisualOrder(string text, int primaryScalar, int fallbackScalar) {
        var fonts = new OfficeFontFaceCollection()
            .Add("Primary", ManagedTextShapingTestAssets.CreateFont(primaryScalar))
            .Add("Fallback", ManagedTextShapingTestAssets.CreateColorFont(fallbackScalar))
            .AddFallbackFamily("Fallback");
        var image = new OfficeRasterImage(100, 60, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(image, fonts: fonts);

        canvas.DrawPositionedText(text, 2D, 2D, 90D, 54D, OfficeColor.Black, 36D,
            OfficeTextAlignment.Left, OfficeFontStyle.Regular, "Primary", 40D,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None,
            fontPalette: "light", textDirection: OfficeTextDirection.RightToLeft);

        int firstColorX = int.MaxValue;
        int firstBlackX = int.MaxValue;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (pixel.R > 180 && pixel.G < 80 && pixel.B < 80) firstColorX = Math.Min(firstColorX, x);
                if (pixel.R < 80 && pixel.G < 80 && pixel.B < 80) firstBlackX = Math.Min(firstBlackX, x);
            }
        }
        Assert.NotEqual(int.MaxValue, firstColorX);
        Assert.NotEqual(int.MaxValue, firstBlackX);
        Assert.True(firstColorX < firstBlackX, $"Expected visual RTL fallback order, got color at {firstColorX} and primary ink at {firstBlackX}.");
    }

    [Fact]
    public void FallbackRunsRetainTheSameHeightFittingAsASingleFont() {
        var fallback = new OfficeFontFaceCollection()
            .Add("Primary", ManagedTextShapingTestAssets.CreateFont('A'))
            .Add("Fallback", ManagedTextShapingTestAssets.CreateFont('B'))
            .AddFallbackFamily("Fallback");
        var complete = new OfficeFontFaceCollection()
            .Add("Primary", ManagedTextShapingTestAssets.CreateFont('A', 'B'));
        static OfficeRasterImage Render(OfficeFontFaceCollection fonts) {
            var image = new OfficeRasterImage(100, 40, OfficeColor.White);
            new OfficeRasterCanvas(image, fonts: fonts).DrawText("AB", 4D, 4D, 80D, 14D,
                OfficeColor.Black, 24D, fontFamily: "Primary");
            return image;
        }
        Assert.Equal(2, fallback.PlanFallbackRuns("AB", "Primary").Count);
        AssertRasterImagesEqual(Render(complete), Render(fallback));
    }
}
