using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeShadowLayerPlanner_UsesExpandedPrimitiveLayersWithoutOverdarkeningCorners() {
        OfficeShape roundedRectangle = OfficeShape.RoundedRectangle(80D, 30D, 10D);
        IReadOnlyList<OfficeShadowLayer> layers = OfficeShadowLayerPlanner.Create(
            opacity: 0.2D,
            blurRadius: 28D,
            baseStrokeWidth: 0D,
            hasFill: false,
            hasStroke: false,
            canExpand: OfficeShadowLayerPlanner.CanExpand(roundedRectangle));

        Assert.Equal(14, layers.Count);
        Assert.All(layers, layer => {
            Assert.True(layer.Expansion > 0D);
            Assert.True(layer.HasFill);
            Assert.False(layer.HasStroke);
        });
        double compositeOpacity = 1D;
        foreach (OfficeShadowLayer layer in layers) compositeOpacity *= 1D - layer.Opacity;
        Assert.Equal(0.2D, 1D - compositeOpacity, 6);

        OfficeDrawing drawing = new OfficeDrawing(140D, 90D);
        roundedRectangle.FillColor = OfficeColor.White;
        roundedRectangle.StrokeWidth = 0D;
        roundedRectangle.Shadow = new OfficeShadow(OfficeColor.Black, 0.2D, 0D, 8D, 28D);
        drawing.AddShape(roundedRectangle, 30D, 15D);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);

        OfficeColor center = raster.GetPixel(70, 50);
        OfficeColor roundedCorner = raster.GetPixel(30, 50);
        Assert.InRange(center.R, (byte)195, (byte)220);
        Assert.True(roundedCorner.R >= center.R - 20, $"Rounded blur corner {roundedCorner} was materially darker than center {center}.");
    }
}
