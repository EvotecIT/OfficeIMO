using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeDiagramDrawingRenderer_ClipsProcessConnectorToNodeEdges() {
        var snapshot = new OfficeDiagramSnapshot("Delivery",
            OfficeDiagramKind.Process, new[] { "Discover", "Deliver" },
            320D, 180D);

        OfficeDrawing drawing = OfficeDiagramDrawingRenderer.Render(snapshot);

        OfficeDrawingShape connector = drawing.Shapes[1];
        Assert.Equal(OfficeShapeKind.Line, connector.Shape.Kind);
        Assert.Equal(136D, connector.X, 6);
        Assert.Equal(90D, connector.Y, 6);
        Assert.Equal(48D, connector.Shape.Width, 6);
        Assert.Equal(0D, connector.Shape.Height, 6);
        Assert.Equal(OfficeLineMarkerKind.Triangle,
            connector.Shape.StrokeEndMarker?.Kind);
    }

    [Theory]
    [InlineData(OfficeDiagramKind.Process)]
    [InlineData(OfficeDiagramKind.Hierarchy)]
    [InlineData(OfficeDiagramKind.Cycle)]
    [InlineData(OfficeDiagramKind.List)]
    [InlineData(OfficeDiagramKind.Matrix)]
    [InlineData(OfficeDiagramKind.Pyramid)]
    [InlineData(OfficeDiagramKind.Relationship)]
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
        Assert.True(drawing.Shapes.Count >= snapshot.Nodes.Count + 1);
        Assert.True(OfficePngReader.TryDecode(png,
            out OfficeRasterImage? raster));
        Assert.NotNull(raster);
        Assert.Equal(320, raster!.Width);
        Assert.Equal(180, raster.Height);
        Assert.Contains(raster.GetPixels(), channel => channel != 255);
    }
}
