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
            canExpand: OfficeShadowLayerPlanner.CanExpand(roundedRectangle),
            minimumShapeDimension: Math.Min(roundedRectangle.Width, roundedRectangle.Height));

        Assert.True(layers.Count > 1);
        Assert.Contains(layers, layer => layer.Expansion > 0D);
        Assert.Contains(layers, layer => layer.Expansion == 0D);
        Assert.Contains(layers, layer => layer.Expansion < 0D);
        Assert.All(layers, layer => {
            Assert.True(layer.HasFill);
            Assert.False(layer.HasStroke);
        });
        Assert.True(layers[layers.Count - 1].Expansion < 0D);
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
        Assert.InRange(center.R, (byte)195, (byte)225);
        Assert.True(roundedCorner.R >= center.R - 20, $"Rounded blur corner {roundedCorner} was materially darker than center {center}.");
    }

    [Fact]
    public void OfficeShadowLayerPlanner_TapersOpaqueBlurInsteadOfPaintingSolidExpandedLayers() {
        OfficeShape rectangle = OfficeShape.Rectangle(80D, 30D);
        IReadOnlyList<OfficeShadowLayer> layers = OfficeShadowLayerPlanner.Create(
            opacity: 1D,
            blurRadius: 12D,
            baseStrokeWidth: 0D,
            hasFill: true,
            hasStroke: false,
            canExpand: OfficeShadowLayerPlanner.CanExpand(rectangle),
            minimumShapeDimension: Math.Min(rectangle.Width, rectangle.Height));

        Assert.True(layers.Count > 1);
        Assert.All(layers.Take(layers.Count - 1), layer => Assert.InRange(layer.Opacity, 0.000001D, 0.999999D));
        Assert.Equal(1D, layers[layers.Count - 1].Opacity);
        Assert.True(layers[0].Opacity < layers[layers.Count - 2].Opacity);

        IReadOnlyList<OfficeShadowLayer> nearOpaqueLayers = OfficeShadowLayerPlanner.Create(
            opacity: 0.999D,
            blurRadius: 12D,
            baseStrokeWidth: 0D,
            hasFill: true,
            hasStroke: false,
            canExpand: true,
            minimumShapeDimension: 30D);
        Assert.Equal(layers.Count, nearOpaqueLayers.Count);
        int firstCoreIndex = Array.FindIndex(layers.ToArray(), layer => layer.Expansion == 0D);
        Assert.True(firstCoreIndex >= 0);
        Assert.InRange(Math.Abs(layers[firstCoreIndex].Opacity - nearOpaqueLayers[firstCoreIndex].Opacity), 0D, 0.01D);

        rectangle.FillColor = OfficeColor.White;
        rectangle.StrokeWidth = 0D;
        rectangle.Shadow = new OfficeShadow(OfficeColor.Black, 1D, 0D, 0D, 12D);
        var drawing = new OfficeDrawing(140D, 80D);
        drawing.AddShape(rectangle, 30D, 25D);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
        OfficeColor outerBlur = raster.GetPixel(20, 40);
        Assert.InRange(outerBlur.R, (byte)20, (byte)250);
        Assert.Equal(outerBlur.R, outerBlur.G);
        Assert.Equal(outerBlur.G, outerBlur.B);
    }

    [Fact]
    public void OfficeShadowLayerPlanner_ExpandedBlurIncludesPaintedStrokeSilhouette() {
        OfficeShape rectangle = OfficeShape.Rectangle(80D, 30D);

        IReadOnlyList<OfficeShadowLayer> layers = OfficeShadowLayerPlanner.Create(
            opacity: 0.4D,
            blurRadius: 12D,
            baseStrokeWidth: 10D,
            hasFill: true,
            hasStroke: true,
            canExpand: OfficeShadowLayerPlanner.CanExpand(rectangle),
            minimumShapeDimension: Math.Min(rectangle.Width, rectangle.Height));

        Assert.True(layers.Count > 1);
        Assert.Equal(17D, layers[0].Expansion, 6);
        Assert.Contains(layers, layer => Math.Abs(layer.Expansion - 5D) < 0.000001D);
        Assert.True(layers[layers.Count - 1].Expansion < 5D);
        Assert.All(layers, layer => {
            Assert.True(layer.HasFill);
            Assert.False(layer.HasStroke);
        });
    }

    [Fact]
    public void OfficeShadowLayerPlanner_NonExpandableFillPreservesRequestedInteriorOpacity() {
        IReadOnlyList<OfficeShadowLayer> layers = OfficeShadowLayerPlanner.Create(
            opacity: 0.5D,
            blurRadius: 12D,
            baseStrokeWidth: 0D,
            hasFill: true,
            hasStroke: false,
            canExpand: false,
            minimumShapeDimension: 30D);

        OfficeShadowLayer core = Assert.Single(layers, layer => layer.HasFill);
        Assert.False(core.HasStroke);
        Assert.Equal(0.5D, core.Opacity, 6);
        Assert.All(layers.Where(layer => !layer.HasFill), layer => Assert.True(layer.HasStroke));
    }
}
