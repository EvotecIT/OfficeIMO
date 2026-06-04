using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingSceneTests {
    [Fact]
    public void OfficeDrawingStoresReusablePositionedShapesInPaintOrder() {
        var background = OfficeShape.Rectangle(120, 60);
        background.FillColor = OfficeColor.WhiteSmoke;

        var marker = OfficeShape.Polygon(
            new OfficePoint(0, 30),
            new OfficePoint(40, 0),
            new OfficePoint(80, 30));
        marker.FillColor = OfficeColor.SteelBlue;
        marker.StrokeColor = OfficeColor.Black;
        marker.StrokeWidth = 1.25;

        var drawing = new OfficeDrawing(120, 60)
            .AddShape(background, 0, 0)
            .AddShape(marker, 20, 15);

        var clone = drawing.Clone();
        background.Width = 10;

        Assert.Equal(120, clone.Width);
        Assert.Equal(60, clone.Height);
        Assert.Equal(2, clone.Shapes.Count);
        Assert.Equal(0, clone.Shapes[0].X);
        Assert.Equal(0, clone.Shapes[0].Y);
        Assert.Equal(OfficeShapeKind.Rectangle, clone.Shapes[0].Shape.Kind);
        Assert.Equal(120, clone.Shapes[0].Shape.Width);
        Assert.Equal(20, clone.Shapes[1].X);
        Assert.Equal(15, clone.Shapes[1].Y);
        Assert.Equal(OfficeShapeKind.Polygon, clone.Shapes[1].Shape.Kind);
        Assert.Equal(OfficeColor.SteelBlue, clone.Shapes[1].Shape.FillColor);
        Assert.Equal(OfficeColor.Black, clone.Shapes[1].Shape.StrokeColor);
    }

    [Fact]
    public void OfficeDrawingRejectsShapesOutsideCanvas() {
        var shape = OfficeShape.Rectangle(40, 20);
        var drawing = new OfficeDrawing(60, 30);

        Assert.Throws<ArgumentOutOfRangeException>(() => drawing.AddShape(shape, 25, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => drawing.AddShape(shape, 0, 15));
    }

    [Fact]
    public void OfficeDrawingQualityAnalyzerReportsCleanDrawing() {
        var drawing = new OfficeDrawing(160, 80)
            .AddText("Title", 8, 8, 80, 14)
            .AddText("Legend", 8, 28, 80, 14)
            .AddShape(OfficeShape.Rectangle(32, 18), 112, 24);

        OfficeDrawingQualityReport report = OfficeDrawingQualityAnalyzer.Analyze(drawing);

        Assert.False(report.HasIssues);
        Assert.Empty(report.Issues);
    }

    [Fact]
    public void OfficeDrawingQualityAnalyzerReportsOverlappingTextBoxes() {
        var drawing = new OfficeDrawing(160, 80)
            .AddText("Revenue", 12, 12, 70, 16)
            .AddText("Target", 48, 16, 70, 16);

        OfficeDrawingQualityReport report = OfficeDrawingQualityAnalyzer.Analyze(drawing);

        OfficeDrawingQualityIssue issue = Assert.Single(report.Issues);
        Assert.True(report.HasIssues);
        Assert.Equal(OfficeDrawingQualityIssueKind.TextOverlap, issue.Kind);
        Assert.Equal(0, issue.ElementIndex);
        Assert.Equal(1, issue.RelatedElementIndex);
        Assert.Contains("Revenue", issue.Message);
        Assert.Contains("Target", issue.Message);
    }
}
