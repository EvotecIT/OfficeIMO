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

    [Fact]
    public void OfficeDiagramDrawingRenderer_CanOmitStandaloneCanvasBackground() {
        var snapshot = new OfficeDiagramSnapshot("Embedded process",
            OfficeDiagramKind.Process, new[] { "Start", "Finish" }, 320D, 180D);

        OfficeDrawing drawing = OfficeDiagramDrawingRenderer.Render(snapshot,
            includeBackground: false);

        Assert.DoesNotContain(drawing.Shapes, shape =>
            shape.X == 0D && shape.Y == 0D
            && shape.Shape.Width == snapshot.WidthPoints
            && shape.Shape.Height == snapshot.HeightPoints
            && shape.Shape.FillColor == OfficeColor.White);
    }

    [Fact]
    public void OfficeDiagramDrawingRenderer_HierarchyConnectsEveryChildToRoot() {
        var snapshot = new OfficeDiagramSnapshot("Organization",
            OfficeDiagramKind.Hierarchy,
            new[] { "Executive", "A", "B", "C", "D", "E", "F" },
            320D, 180D);

        OfficeDrawing drawing = OfficeDiagramDrawingRenderer.Render(snapshot,
            includeBackground: false);
        OfficeDrawingShape[] connectors = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
            .ToArray();
        OfficeDrawingShape[] nodes = drawing.Shapes
            .Where(shape => shape.Shape.Kind != OfficeShapeKind.Line)
            .ToArray();
        OfficeDrawingShape root = nodes[0];

        Assert.Equal(snapshot.Nodes.Count - 1, connectors.Length);
        Assert.Equal(snapshot.Nodes.Count, nodes.Length);
        foreach (OfficeDrawingShape connector in connectors) {
            OfficePoint start = connector.Shape.Points[0];
            double startX = connector.X + start.X;
            double startY = connector.Y + start.Y;
            Assert.InRange(startX, root.X, root.X + root.Shape.Width);
            Assert.InRange(startY, root.Y, root.Y + root.Shape.Height);
        }
    }
}
