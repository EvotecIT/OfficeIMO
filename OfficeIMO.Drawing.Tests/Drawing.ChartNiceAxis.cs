using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartNiceAxisTests {
    private static OfficeDrawing RenderColumns(OfficeChartKind kind, OfficeChartLayout? layout, params double[][] series) {
        string[] categories = Enumerable.Range(1, series[0].Length).Select(i => "C" + i).ToArray();
        return OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Axis", null, kind,
            new OfficeChartData(categories, series.Select((values, index) => new OfficeChartSeries("S" + index, values)).ToArray()),
            widthPoints: 360D, heightPoints: 240D, layout: layout ?? new OfficeChartLayout(showLegend: false)));
    }

    private static List<string> Texts(OfficeDrawing drawing) => drawing.Elements.OfType<OfficeDrawingText>().Select(text => text.Text).ToList();

    [Fact]
    public void AutomaticValueAxisUsesRoundStepsWithHeadroom() {
        // Word draws this chart on 0-250 in steps of 50.
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.ColumnClustered, null,
            new[] { 10D, 35D, 18D, 23D }, new[] { 15D, 30D, 8D, 18D }, new[] { 13D, 20D, 230D, 150D });
        List<string> texts = Texts(drawing);
        foreach (string tick in new[] { "0", "50", "100", "150", "200", "250" }) Assert.Contains(tick, texts);
        Assert.DoesNotContain("57.5", texts);
        Assert.DoesNotContain("230", texts);
    }

    [Theory]
    [InlineData(81D, new[] { "0", "20", "40", "60", "80", "100" })]
    [InlineData(92.5D, new[] { "0", "20", "40", "60", "80", "100" })]
    [InlineData(32D, new[] { "0", "10", "20", "30", "40" })]
    [InlineData(27D, new[] { "0", "5", "10", "15", "20", "25", "30" })]
    public void AutomaticValueAxisPicksFourToSixRoundIntervals(double maximum, string[] expected) {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.ColumnClustered, null, new[] { maximum / 3D, maximum });
        List<string> texts = Texts(drawing);
        foreach (string tick in expected) Assert.Contains(tick, texts);
        Assert.DoesNotContain(maximum.ToString(System.Globalization.CultureInfo.InvariantCulture), texts);
    }

    [Fact]
    public void TallestColumnStaysBelowTopOfPlot() {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.ColumnClustered, null, new[] { 40D, 80D });
        OfficeDrawingText top = drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "100");
        OfficeDrawingText eighty = drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "80");
        OfficeDrawingShape tallest = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Rectangle && shape.Shape.Width < 200D && shape.Shape.Height > 20D)
            .OrderBy(shape => shape.Y).First();
        // The 80 column ends on the 80 tick, one step below the top of the axis.
        Assert.Equal(eighty.Y + eighty.Height / 2D, tallest.Y, precision: 0);
        Assert.True(tallest.Y > top.Y + top.Height, "The tallest column must leave headroom below the top of the axis.");
    }

    [Fact]
    public void PercentStackedAxisStopsAtOneHundredPercent() {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.ColumnStacked100, null, new[] { 1D, 2D }, new[] { 3D, 1D });
        List<string> texts = Texts(drawing);
        Assert.Contains("100%", texts);
        Assert.DoesNotContain("120%", texts);
    }

    [Fact]
    public void ExplicitBoundsAndUnitsAreKept() {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.ColumnClustered,
            new OfficeChartLayout(showLegend: false, verticalAxisMinimum: 0D, verticalAxisMaximum: 90D, verticalAxisMajorUnit: 30D),
            new[] { 10D, 81D });
        List<string> texts = Texts(drawing);
        foreach (string tick in new[] { "0", "30", "60", "90" }) Assert.Contains(tick, texts);
        Assert.DoesNotContain("100", texts);
    }

    [Fact]
    public void ExplicitMaximumWithAutomaticUnitUsesRoundTicks() {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.ColumnClustered,
            new OfficeChartLayout(showLegend: false, verticalAxisMaximum: 92.5D), new[] { 10D, 81D });
        List<string> texts = Texts(drawing);
        Assert.Contains("80", texts);
        Assert.DoesNotContain("69.375", texts);
    }

    [Fact]
    public void NegativeValuesGetRoundBoundsOnBothSides() {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.ColumnClustered, null, new[] { -10D, 50D });
        List<string> texts = Texts(drawing);
        foreach (string tick in new[] { "-20", "0", "20", "40", "60" }) Assert.Contains(tick, texts);
    }

    [Fact]
    public void VerticalAxisTitleIsRotatedBesideTheTickLabels() {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.ColumnClustered,
            new OfficeChartLayout(showLegend: false, valueAxisTitle: "Revenue", verticalAxisDisplayUnitLabel: "Thousands USD"),
            new[] { 10D, 35D });
        OfficeDrawingText title = drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "Revenue");
        Assert.Equal(-90D, title.RotationDegrees);
        OfficeDrawingText[] tickLabels = drawing.Elements.OfType<OfficeDrawingText>()
            .Where(text => text.Text is "0" or "10" or "20" or "30" or "40").ToArray();
        Assert.NotEmpty(tickLabels);
        // The rotated title occupies a band of its own height left of every tick label.
        // Turned a quarter turn about a point half its length in from its corner, the title's
        // layout box lands in a band of its own height that starts at its X position.
        Assert.Equal(title.X + title.Width / 2D, title.RotationCenterX, precision: 6);
        Assert.Equal(title.Y + title.Width / 2D, title.RotationCenterY, precision: 6);
        double titleRight = title.X + title.Height;
        Assert.All(tickLabels, label => Assert.True(label.X >= titleRight - 0.01D));
        OfficeDrawingText units = drawing.Elements.OfType<OfficeDrawingText>().Single(text => text.Text == "Thousands USD");
        Assert.True(units.X >= titleRight - 0.01D);
    }

    [Fact]
    public void LinearTrendlineIsDrawnWithEquationAndRSquared() {
        var series = new OfficeChartSeries("Sales", new[] { 3D, 5D, 7D, 9D })
            .WithTrendline(new OfficeChartTrendline(OfficeChartTrendlineType.Linear, displayEquation: true, displayRSquared: true));
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Trend", null, OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "Q1", "Q2", "Q3", "Q4" }, new[] { series }), 360D, 240D,
            layout: new OfficeChartLayout(showLegend: false)));
        Assert.Contains(Texts(drawing), text => text == "y = 2x + 1   R\u00B2 = 1");
    }

    [Fact]
    public void TopLegendStartsAtTheLeftWithPackedEntries() {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.ColumnClustered,
            new OfficeChartLayout(legendPosition: OfficeChartLegendPosition.Top), new[] { 1D, 2D }, new[] { 2D, 3D }, new[] { 3D, 1D });
        OfficeDrawingText[] entries = drawing.Elements.OfType<OfficeDrawingText>().Where(text => text.Text is "S0" or "S1" or "S2")
            .OrderBy(text => text.X).ToArray();
        Assert.Equal(3, entries.Length);
        Assert.True(entries[0].X < 40D, "The first legend entry starts at the left edge.");
        Assert.True(entries[2].X - entries[0].X < 120D, "Legend entries sit together instead of spreading across the chart.");
    }

    [Fact]
    public void LineSeriesLegendUsesAStrokeSwatch() {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.Line,
            new OfficeChartLayout(legendPosition: OfficeChartLegendPosition.Top), new[] { 1D, 2D }, new[] { 2D, 3D });
        OfficeDrawingText entry = drawing.Elements.OfType<OfficeDrawingText>().First(text => text.Text == "S0");
        Assert.Contains(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.Height == 0D &&
            shape.X < entry.X && shape.X > entry.X - 20D && Math.Abs(shape.Y - (entry.Y + entry.Height / 2D)) < 6D);
    }

    [Fact]
    public void NeighbouringDataLabelsDoNotOverlap() {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.ColumnClustered,
            new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true, dataLabelPosition: OfficeChartDataLabelPosition.OutsideEnd),
            new[] { 10D, 20D }, new[] { 12D, 21D }, new[] { 11D, 20.5D });
        var labels = drawing.Elements.OfType<OfficeDrawingText>()
            .Where(text => text.X > 30D && text.Text is "10" or "12" or "11" or "20" or "21" or "20.5").ToList();
        Assert.Equal(6, labels.Count);
        for (int i = 0; i < labels.Count; i++) {
            for (int j = i + 1; j < labels.Count; j++) {
                double overlapX = Math.Min(labels[i].X + labels[i].Width, labels[j].X + labels[j].Width) - Math.Max(labels[i].X, labels[j].X);
                bool sameRow = Math.Abs(labels[i].Y - labels[j].Y) < 0.5D * labels[i].Height;
                // Boxes may touch horizontally, but two labels on one row never share their text area.
                Assert.False(sameRow && overlapX > Math.Max(labels[i].Width, labels[j].Width) * 0.5D,
                    labels[i].Text + " overlaps " + labels[j].Text);
            }
        }
    }

    [Fact]
    public void AreaLabelsAreDrawnAboveEveryAreaFill() {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.Area,
            new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true), new[] { 100D, 1D, 18D }, new[] { 13D, 20D, 230D });
        List<object> elements = drawing.Elements.Cast<object>().ToList();
        int lastFill = elements.FindLastIndex(element => element is OfficeDrawingShape shape && shape.Shape.FillColor.HasValue &&
            shape.Shape.Kind == OfficeShapeKind.Polygon);
        int firstLabel = elements.FindIndex(element => element is OfficeDrawingText text && text.Text == "100");
        Assert.True(lastFill >= 0 && firstLabel > lastFill, "Data labels of the first series must follow the last area fill.");
    }

    [Fact]
    public void CoincidentDataLabelsAreMovedApart() {
        OfficeDrawing drawing = RenderColumns(OfficeChartKind.Line,
            new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true, dataLabelPosition: OfficeChartDataLabelPosition.Top),
            new[] { 12D, 18D }, new[] { 12D, 18D });
        OfficeDrawingText[] twelves = drawing.Elements.OfType<OfficeDrawingText>().Where(text => text.Text == "12" && text.X > 30D).ToArray();
        Assert.Equal(2, twelves.Length);
        Assert.True(Math.Abs(twelves[0].Y - twelves[1].Y) >= 8D, "Two labels for the same point position must not be drawn on top of each other.");
    }
}
