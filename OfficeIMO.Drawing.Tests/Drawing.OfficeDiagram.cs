using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(OfficeDiagramKind.Process)]
    [InlineData(OfficeDiagramKind.Hierarchy)]
    [InlineData(OfficeDiagramKind.Cycle)]
    public void OfficeDiagramDrawingRenderer_RendersBoundedSemanticNodes(
        OfficeDiagramKind kind) {
        var snapshot = new OfficeDiagramSnapshot("Delivery", kind,
            new[] { "Discover", "Build", "Validate", "Ship" },
            320D, 180D);

        OfficeDrawing drawing = OfficeDiagramDrawingRenderer.Render(snapshot);
        byte[] png = OfficeDrawingRasterRenderer.ToPng(drawing,
            background: OfficeColor.White);

        Assert.Equal(320D, drawing.Width);
        Assert.Equal(180D, drawing.Height);
        Assert.True(drawing.Shapes.Count >= 8);
        Assert.True(OfficePngReader.TryDecode(png,
            out OfficeRasterImage? raster));
        Assert.NotNull(raster);
        Assert.Equal(320, raster!.Width);
        Assert.Equal(180, raster.Height);
        Assert.Contains(raster.GetPixels(), channel => channel != 255);
    }
}
