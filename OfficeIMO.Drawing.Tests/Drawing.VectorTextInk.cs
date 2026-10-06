using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingVectorTextInkTests {
    [Theory]
    [InlineData(0, false)]
    [InlineData(90, true)]
    public void DrawingViewportCropsTextBeforePlacement(double textX, bool cropped) {
        var drawing = Create(textX);
        var results = Inspect(drawing, OfficeTransform.Scale(2, 2).Then(OfficeTransform.Translate(10, 20)));
        var ink = Assert.Single(results);
        Assert.True(ink.Measured && ink.HasInk); Assert.Equal(cropped, ink.Clipped);
        Assert.InRange(ink.Left, 10, 210); Assert.InRange(ink.Right, ink.Left + .01, 210);
        var raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Contains(Enumerable.Range(0, 100).SelectMany(y => Enumerable.Range(0, 100).Select(x => raster.GetPixel(x, y).A)), a => a > 0);
    }

    [Fact]
    public void NestedEffectUsesItsOwnFontAndViewport() {
        var child = Create(90);
        var parent = new OfficeDrawing(200, 200);
        parent.AddEffectDrawing(child, OfficeTransform.Translate(20, 30));
        var ink = Assert.Single(Inspect(parent, OfficeTransform.Identity));
        Assert.True(ink.Measured && ink.Clipped && ink.HasInk);
        Assert.InRange(ink.Left, 110, 120); Assert.Equal(120, ink.Right, 5);
    }

    [Fact]
    public void UnsupportedLayoutIsExplicitAndHiddenGroupsAreSkipped() {
        var child = Create(0);
        var hidden = new OfficeDrawing(100, 100); hidden.AddEffectDrawing(child, OfficeTransform.Identity, opacity: 0);
        Assert.Empty(Inspect(hidden, OfficeTransform.Identity));
        var wrapped = new OfficeDrawing(100, 100);
        wrapped.AddText("A", 0, 0, 100, 100, wrapText: true);
        var result = Assert.Single(Inspect(wrapped, OfficeTransform.Identity));
        Assert.False(result.Measured); Assert.NotNull(result.Reason);
    }

    [Fact]
    public void OuterClipAndGroupContentOffsetAreRespected() {
        var child = Create(0);
        var parent = new OfficeDrawing(200, 200);
        parent.AddClippedDrawing(child, 20, 30, OfficeClipPath.Rectangle(100, 100), 10, 0);
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1));
        int count = 0;
        canvas.InspectDrawingTextInk(parent, OfficeTransform.Identity,
            new[] { new OfficeTextInkClip(0, 0, 1, 1, true, true, OfficeTransform.Identity) }, (ink, reason) => {
                count++; Assert.True(ink.IsMeasured && ink.IsClipped); Assert.False(ink.HasInk); Assert.Null(reason);
            });
        Assert.Equal(1, count);
    }

    private static OfficeDrawing Create(double x) {
        var drawing = new OfficeDrawing(100, 100).AddFont("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddPositionedTextWithNaturalAdvance("A", x, 20, 10, 40, new OfficeFontInfo("Ink", 30), OfficeColor.Black, 40);
        return drawing;
    }

    private static List<(double Left, double Right, bool HasInk, bool Measured, bool Clipped, string? Reason)> Inspect(OfficeDrawing drawing, OfficeTransform transform) {
        var result = new List<(double, double, bool, bool, bool, string?)>();
        new OfficeRasterCanvas(new OfficeRasterImage(1, 1)).InspectDrawingTextInk(drawing, transform,
            Array.Empty<OfficeTextInkClip>(), (ink, reason) => result.Add((ink.Left, ink.Right, ink.HasInk, ink.IsMeasured, ink.IsClipped, reason)));
        return result;
    }
}
