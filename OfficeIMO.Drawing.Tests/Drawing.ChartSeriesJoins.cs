using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartSeriesJoinTests {
    private static readonly OfficeColor Ink = OfficeColor.Parse("#D900AA");

    [Theory]
    [InlineData(OfficeChartType.Line)]
    [InlineData(OfficeChartType.Scatter)]
    public void ExplicitJoinConnectsVerticesInSvgAndChangesCornerPixels(OfficeChartType kind) {
        var images = new List<byte[]>();
        foreach (OfficeStrokeLineJoin join in new[] { OfficeStrokeLineJoin.Bevel, OfficeStrokeLineJoin.Round, OfficeStrokeLineJoin.Miter }) {
            OfficeDrawing drawing = Render(kind, new[] { 2d, 8d, 2d, 6d }, join);
            OfficeDrawingShape path = Assert.Single(SeriesInk(drawing));
            Assert.Equal(OfficeShapeKind.Path, path.Shape.Kind);
            Assert.Equal(3, path.Shape.PathCommands.Count(command => command.Kind == OfficePathCommandKind.LineTo));
            Assert.DoesNotContain(path.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.Close);
            Assert.Equal(join, path.Shape.StrokeLineJoin);
            string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
            Assert.Contains("stroke-linejoin=\"" + join.ToString().ToLowerInvariant() + "\"", svg);

            // Compare only the first real corner, independently of text/axes and endpoint caps.
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, 2, OfficeColor.White);
            OfficePoint corner = path.Shape.PathCommands[1].Point;
            int cx = (int)Math.Round((path.X + corner.X) * 2), cy = (int)Math.Round((path.Y + corner.Y) * 2);
            var pixels = new List<byte>();
            for (int y = cy - 24; y <= cy + 24; y++)
                for (int x = cx - 24; x <= cx + 24; x++) pixels.Add(image.GetPixel(x, y).G);
            images.Add(pixels.ToArray());
        }
        Assert.False(images[0].SequenceEqual(images[1]), "Bevel and round must paint different corner pixels.");
        Assert.False(images[1].SequenceEqual(images[2]), "Round and miter must paint different corner pixels.");

    }

    [Theory]
    [InlineData(OfficeChartType.Line)]
    [InlineData(OfficeChartType.Scatter)]
    public void JoinedRunsStopAtGapsAndOutsideExcursions(OfficeChartType kind) {
        double[] values = { 2, 8, 2, double.NaN, 2, 8, 2, 20, 20, 2, 8, 2 };
        OfficeDrawing drawing = Render(kind, values, OfficeStrokeLineJoin.Round, bounded: true);
        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>(), candidate => SeriesInk(candidate.Drawing).Any());
        OfficeDrawingShape[] paths = SeriesInk(group.Drawing).ToArray();
        Assert.Equal(3, paths.Length);
        Assert.All(paths, path => {
            Assert.Equal(OfficeShapeKind.Path, path.Shape.Kind);
            Assert.DoesNotContain(path.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.Close);
            foreach (OfficePathCommand command in path.Shape.PathCommands) {
                Assert.InRange(path.X + command.Point.X, group.X - .001, group.X + group.ClipPath.Width + .001);
                Assert.InRange(path.Y + command.Point.Y, group.Y - .001, group.Y + group.ClipPath.Height + .001);
            }
        });
        Assert.Equal(new[] { 2, 3, 3 }, paths.Select(path => path.Shape.PathCommands.Count - 1));
        Assert.NotEmpty(OfficeDrawingRasterRenderer.ToPng(drawing));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void JoinedScatterSafelyClipsOppositeFiniteExtremes(bool diagonal) {
        var series = new OfficeSeries("Series", diagonal ? new[] { -double.MaxValue, double.MaxValue } : new[] { 5d, 5d },
            new[] { -double.MaxValue, double.MaxValue }, Ink, null, showMarkers: false)
            .WithSeriesStyle(Paint(OfficeStrokeLineJoin.Bevel));
        OfficeDrawing drawing = Render(OfficeChartType.Scatter, series, 2, new OfficeChartLayout(showLegend: false,
            horizontalAxisMinimum: 0, horizontalAxisMaximum: 10, verticalAxisMinimum: 0, verticalAxisMaximum: 10));
        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>(), candidate => SeriesInk(candidate.Drawing).Any());
        OfficeDrawingShape path = Assert.Single(SeriesInk(group.Drawing));
        Assert.Equal(2, path.Shape.PathCommands.Count);
        Assert.True(path.Shape.Width >= group.ClipPath.Width * .99);
        Assert.NotEmpty(OfficeDrawingRasterRenderer.ToPng(drawing));
    }

    [Theory]
    [InlineData(OfficeChartType.Area, false)]
    [InlineData(OfficeChartType.Area, true)]
    [InlineData(OfficeChartType.Radar, false)]
    [InlineData(OfficeChartType.Radar, true)]
    public void ClosedSeriesKeepsClosingJoinAndIndependentFill(OfficeChartType kind, bool noFill) {
        OfficeDrawing drawing = Render(kind, new[] { 2d, 8d, 2d, 6d }, OfficeStrokeLineJoin.Bevel, noFill: noFill);
        OfficeDrawingShape outline = Assert.Single(SeriesInk(drawing));
        Assert.Equal(OfficeStrokeLineJoin.Bevel, outline.Shape.StrokeLineJoin);
        Assert.True(
            (outline.Shape.Kind == OfficeShapeKind.Polygon && outline.Shape.Points.Count >= 3) ||
            (outline.Shape.Kind == OfficeShapeKind.Path &&
             outline.Shape.PathCommands.Count(command => command.Kind == OfficePathCommandKind.MoveTo) == 1 &&
             outline.Shape.PathCommands.Count(command => command.Kind == OfficePathCommandKind.LineTo) >= 2 &&
             outline.Shape.PathCommands.Last().Kind == OfficePathCommandKind.Close),
            "The series boundary must be closed so its first vertex receives a closing join.");
        Assert.Equal(!noFill, Shapes(drawing).Any(shape => shape.Shape.FillColor == Ink));

        var series = new OfficeSeries("Series", new[] { 2d, 8d, 2d, 6d }, null, Ink, null, false)
            .WithSeriesStyle(new OfficeChartPointStyle(fillColor: Ink, showOutline: false, outlineJoin: OfficeStrokeLineJoin.Bevel));
        OfficeDrawing filled = Render(kind, series, 4, new OfficeChartLayout(showLegend: false));
        Assert.Empty(SeriesInk(filled));
        Assert.Contains(Shapes(filled), shape => shape.Shape.FillColor == Ink);
    }

    [Fact]
    public void ClippedAreaJoinsOriginalLowerCornersWithoutInventingAnUpperEdge() {
        OfficeDrawing drawing = Render(OfficeChartType.Area, new[] { 20d, 20d }, OfficeStrokeLineJoin.Bevel, bounded: true);
        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>(), candidate => SeriesInk(candidate.Drawing).Any());
        OfficeDrawingShape outline = Assert.Single(SeriesInk(group.Drawing));
        Assert.Equal(3, outline.Shape.PathCommands.Count(command => command.Kind == OfficePathCommandKind.LineTo));
        Assert.DoesNotContain(outline.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.Close);
        OfficePoint[] points = outline.Shape.PathCommands.Select(command => command.Point).ToArray();
        Assert.Equal(points[0].Y, points[points.Length - 1].Y);
        Assert.True(points[1].Y > points[0].Y);
        Assert.True(points[2].Y > points[0].Y);
        Assert.NotEmpty(OfficeDrawingRasterRenderer.ToPng(drawing));
    }

    [Theory]
    [InlineData(OfficeStrokeLineJoin.Round)]
    [InlineData(null)]
    public void RadarGapKeepsSeamEdgeWithoutBridgingMissingPoint(OfficeStrokeLineJoin? join) {
        OfficeDrawing drawing = Render(OfficeChartType.Radar, new[] { 4d, 6d, 5d, double.NaN, 7d }, join);
        OfficeDrawingShape[] paths = SeriesInk(drawing).ToArray();
        Assert.Equal(3, paths.Sum(path => path.Shape.Kind == OfficeShapeKind.Line ? 1 :
            path.Shape.PathCommands.Count(command => command.Kind == OfficePathCommandKind.LineTo)));
        Assert.All(paths, path => Assert.DoesNotContain(path.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.Close));
    }

    private static OfficeChartPointStyle Paint(OfficeStrokeLineJoin? join, bool noFill = true) =>
        new(noFill: noFill, outlineColor: Ink, outlineWidth: global::OfficeIMO.OfficeLength.Points(16), showOutline: true, outlineJoin: join);

    private static OfficeDrawing Render(OfficeChartType kind, double[] values, OfficeStrokeLineJoin? join,
        bool bounded = false, bool noFill = true) {
        var series = new OfficeSeries("Series", values, Enumerable.Range(0, values.Length).Select(index => (double)index),
            Ink, null, showMarkers: false).WithSeriesStyle(Paint(join, noFill));
        var layout = bounded ? new OfficeChartLayout(showLegend: false, horizontalAxisMinimum: 0,
            horizontalAxisMaximum: values.Length - 1, verticalAxisMinimum: 0, verticalAxisMaximum: 10)
            : new OfficeChartLayout(showLegend: false);
        return Render(kind, series, values.Length, layout);
    }

    private static OfficeDrawing Render(OfficeChartType kind, OfficeSeries series, int count, OfficeChartLayout layout) =>
        OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null, kind,
            new OfficeChartData(Enumerable.Range(0, count).Select(index => index.ToString()), new[] { series }),
            global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(300), layout: layout));

    private static IEnumerable<OfficeDrawingShape> SeriesInk(OfficeDrawing drawing) =>
        Shapes(drawing).Where(shape => shape.Shape.StrokeColor == Ink);

    private static IEnumerable<OfficeDrawingShape> Shapes(OfficeDrawing drawing) => drawing.Shapes.Concat(
        drawing.Elements.OfType<OfficeDrawingGroup>().SelectMany(group => Shapes(group.Drawing)));
}
