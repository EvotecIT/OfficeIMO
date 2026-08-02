using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeDiagramDrawingRenderer_ProcessDoesNotInventEdges() {
        var snapshot = new OfficeDiagramSnapshot("Delivery",
            OfficeDiagramKind.Process, new[] { "Discover", "Deliver" },
            320D, 180D);

        OfficeDrawing drawing = OfficeDiagramDrawingRenderer.Render(snapshot);

        Assert.DoesNotContain(drawing.Shapes,
            shape => shape.Shape.Kind == OfficeShapeKind.Line);
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
    public void OfficeDiagramDrawingRenderer_HierarchyDoesNotInventEdges() {
        var snapshot = new OfficeDiagramSnapshot("Organization",
            OfficeDiagramKind.Hierarchy,
            new[] { "Executive", "A", "B", "C", "D", "E", "F" },
            320D, 180D);

        OfficeDrawing drawing = OfficeDiagramDrawingRenderer.Render(snapshot,
            includeBackground: false);
        OfficeDrawingShape[] connectors = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
            .ToArray();
        Assert.Empty(connectors);
    }

    [Fact]
    public void OfficeDiagramDrawingRenderer_RelationshipDoesNotInventEdges() {
        var snapshot = new OfficeDiagramSnapshot("Relationships",
            OfficeDiagramKind.Relationship,
            new[] { "Center", "A", "B", "C" }, 320D, 180D);

        OfficeDrawing drawing = OfficeDiagramDrawingRenderer.Render(snapshot,
            includeBackground: false);
        OfficeDrawingShape[] connectors = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
            .ToArray();

        Assert.Empty(connectors);
    }

    [Fact]
    public void OfficeDiagramDrawingRenderer_PyramidUsesAuthoredNormalizedGeometry() {
        var snapshot = new OfficeDiagramSnapshot("Priorities",
            OfficeDiagramKind.Pyramid,
            new[] { "A", "B", "C", "D" }, 320D, 180D);

        OfficeDrawing drawing = OfficeDiagramDrawingRenderer.Render(snapshot,
            includeBackground: false);
        OfficeDrawingShape[] nodes = drawing.Shapes
            .Where(shape => shape.Shape.Kind != OfficeShapeKind.Line)
            .ToArray();

        Assert.Equal(4, nodes.Length);
        for (int index = 0; index < nodes.Length; index++) {
            double progress = index / 3D;
            double expectedWidth = 320D * (0.28D + 0.5D * progress);
            double cellHeight = 180D * 0.82D / 4D;
            double expectedHeight = Math.Min(180D * 0.19D,
                cellHeight * 0.86D);
            double expectedCenterY = 180D * 0.09D
                + (index + 0.5D) * cellHeight;
            Assert.Equal(expectedWidth, nodes[index].Shape.Width, 6);
            Assert.Equal(expectedHeight, nodes[index].Shape.Height, 6);
            Assert.Equal((320D - expectedWidth) / 2D,
                nodes[index].X, 6);
            Assert.Equal(expectedCenterY - expectedHeight / 2D,
                nodes[index].Y, 6);
        }
    }
}
