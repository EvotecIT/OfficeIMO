using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Pins category geometry to what Office draws for the same native chart settings.</summary>
public sealed class DrawingChartCategoryGeometryTests {
    private static readonly OfficeColor First = OfficeColor.Parse("#D900AA");
    private static readonly OfficeColor Second = OfficeColor.Parse("#00AAD9");

    [Fact]
    public void HorizontalBarsRunBottomUpWithFirstSeriesLowest() {
        // Office writes an explicit minMax orientation and draws the first category at the bottom.
        OfficeDrawing drawing = Render(OfficeChartType.BarClustered,
            new OfficeChartLayout(showLegend: false, categoryAxisOrientationSpecified: true));

        // First has 2 then 3, Second 4 then 1, so bar lengths identify each category.
        OfficeDrawingShape firstA = Bars(drawing, First).OrderBy(shape => shape.Shape.Width).First();
        OfficeDrawingShape firstB = Bars(drawing, First).OrderBy(shape => shape.Shape.Width).Last();
        OfficeDrawingShape secondA = Bars(drawing, Second).OrderBy(shape => shape.Shape.Width).Last();
        Assert.True(firstA.Y > firstB.Y, "The first category should sit below the second.");
        Assert.True(firstA.Y > secondA.Y, "The first series should sit below the second within its category.");
    }

    [Fact]
    public void ReversedHorizontalBarsRunTopDown() {
        OfficeDrawing drawing = Render(OfficeChartType.BarClustered,
            new OfficeChartLayout(showLegend: false, reverseCategoryAxis: true, categoryAxisOrientationSpecified: true));

        OfficeDrawingShape firstA = Bars(drawing, First).OrderBy(shape => shape.Shape.Width).First();
        OfficeDrawingShape firstB = Bars(drawing, First).OrderBy(shape => shape.Shape.Width).Last();
        Assert.True(firstA.Y < firstB.Y);
    }

    [Fact]
    public void GapWidthAndOverlapPlaceClusteredColumns() {
        OfficeDrawing standard = Render(OfficeChartType.ColumnClustered, new OfficeChartLayout(showLegend: false));
        OfficeDrawing narrow = Render(OfficeChartType.ColumnClustered,
            new OfficeChartLayout(showLegend: false).WithBarGeometry(gapWidthPercent: 50, overlapPercent: 50));

        // Two bars, a 150% gap: each bar is a 3.5th of the slot. A 50% gap with 50% overlap: a 2nd.
        double slot = Bars(standard, First)[1].X - Bars(standard, First)[0].X;
        Assert.Equal(slot / 3.5D, Bars(standard, First)[0].Shape.Width, 3);
        double narrowSlot = Bars(narrow, First)[1].X - Bars(narrow, First)[0].X;
        OfficeDrawingShape bar = Bars(narrow, First)[0];
        Assert.Equal(narrowSlot / 2D, bar.Shape.Width, 3);
        Assert.Equal(bar.Shape.Width / 2D, Bars(narrow, Second)[0].X - bar.X, 3);
    }

    [Theory]
    [InlineData(OfficeChartType.Line)]
    [InlineData(OfficeChartType.Area)]
    public void LineAndAreaPointsSitInTheMiddleOfTheirCategoryUnlessOnTickMarks(OfficeChartType kind) {
        double between = LeftmostPoint(Render(kind, new OfficeChartLayout(showLegend: false)), out double plotLeft, out double plotWidth);
        Assert.Equal(plotLeft + plotWidth / 4D, between, 1);

        double onTicks = LeftmostPoint(Render(kind, new OfficeChartLayout(showLegend: false).WithCategoryPointsOnTickMarks(true)),
            out plotLeft, out _);
        Assert.Equal(plotLeft, onTicks, 1);
    }

    [Fact]
    public void HighLowLinesJoinTheExtremesOfEachCategory() {
        OfficeColor rule = OfficeColor.Parse("#123456");
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("High", new[] { 40d, 30d }, null, First, null, showMarkers: false, connectLine: false),
            new OfficeSeries("Low", new[] { 10d, 20d }, null, Second, null, showMarkers: false, connectLine: false)
        });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null, OfficeChartType.Line, data,
            360, 240, layout: new OfficeChartLayout(showLegend: false, verticalAxisMinimum: 0, verticalAxisMaximum: 50)
                .WithHighLowLines(true, rule, 1.5D)));

        OfficeDrawingShape[] rules = drawing.Shapes.Concat(drawing.Elements.OfType<OfficeDrawingGroup>().SelectMany(group => group.Drawing.Shapes))
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == rule).OrderBy(shape => shape.X).ToArray();
        Assert.Equal(2, rules.Length);
        Assert.All(rules, line => Assert.Equal(1.5D, line.Shape.StrokeWidth));
        Assert.Equal(3D, rules[0].Shape.Height / rules[1].Shape.Height, 3);
    }

    [Fact]
    public void DateCategoryLabelsUseTheirSpreadsheetDateFormat() {
        var data = new OfficeChartData(new[] { "37377", "37408" }, new[] {
            new OfficeSeries("Series", new[] { 1d, 2d }, null, First)
        });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null, OfficeChartType.Area, data,
            360, 240, layout: new OfficeChartLayout(showLegend: false, categoryAxisNumberFormat: "m/d/yyyy")));

        string[] labels = drawing.Elements.OfType<OfficeDrawingText>().Select(text => text.Text).ToArray();
        Assert.Contains("5/1/2002", labels);
        Assert.Contains("6/1/2002", labels);
    }

    private static OfficeDrawing Render(OfficeChartType kind, OfficeChartLayout layout) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("First", new[] { 2d, 3d }, null, First),
            new OfficeSeries("Second", new[] { 4d, 1d }, null, Second)
        });
        return OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null, kind, data, 360, 240, layout: layout));
    }

    private static OfficeDrawingShape[] Bars(OfficeDrawing drawing, OfficeColor color) =>
        drawing.Shapes.Where(shape => shape.Shape.Kind == OfficeShapeKind.Rectangle && shape.Shape.FillColor == color)
            .OrderBy(shape => shape.X).ToArray();

    private static double LeftmostPoint(OfficeDrawing drawing, out double plotLeft, out double plotWidth) {
        OfficeDrawingShape axis = drawing.Shapes.Where(shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.Height == 0D)
            .OrderByDescending(shape => shape.Shape.Width).First();
        plotLeft = axis.X;
        plotWidth = axis.Shape.Width;
        return drawing.Shapes.Where(shape => shape.Shape.StrokeColor == First || shape.Shape.FillColor == First)
            .Where(shape => shape.Shape.Kind is OfficeShapeKind.Line or OfficeShapeKind.Polygon)
            .Min(shape => shape.X);
    }
}
