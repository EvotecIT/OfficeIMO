using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingRasterTests {
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
