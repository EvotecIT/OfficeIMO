using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingChartAxisLayoutTests {
    [Fact]
    public void OfficeChartDrawingRenderer_BoundsCategoryLabelMeasurementWork() {
        const int categoryCount = 5_000;
        string oversizedLabel = new string('W', 100_000);
        string[] categories = Enumerable.Repeat(oversizedLabel, categoryCount).ToArray();
        double[] values = Enumerable.Repeat(1D, categoryCount).ToArray();

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Bounded categories",
            null,
            OfficeChartKind.BarClustered,
            new OfficeChartData(categories, new[] { new OfficeChartSeries("Series", values) }),
            widthPoints: 320D,
            heightPoints: 190D,
            layout: new OfficeChartLayout(maximumHorizontalCategoryAxisLabels: 8)));

        Assert.NotEmpty(drawing.Elements);
        Assert.True(drawing.Elements.OfType<OfficeDrawingText>().Count() < 100,
            "A large category collection should still render a bounded number of labels.");
    }

    [Fact]
    public void OfficeChartDrawingRenderer_ReservesMeasuredVerticalValueAxisLabelBand() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Measured vertical axis",
            "Measured Vertical Axis",
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(
                new[] { "Q1", "Q2" },
                new[] {
                    new OfficeChartSeries("Actual", new[] { 1200000D, 1800000D })
                }),
            widthPoints: 320D,
            heightPoints: 190D,
            layout: new OfficeChartLayout(verticalAxisNumberFormat: "$#,##0.00")));

        OfficeDrawingText valueAxisLabel = drawing.Elements
            .OfType<OfficeDrawingText>()
            .First(label => label.Text == "$1,800,000.00");

        Assert.True(valueAxisLabel.Width > 60D, "Long formatted value-axis labels should reserve measured width instead of using the legacy fixed band.");
        Assert.True(valueAxisLabel.X >= 0D);
    }

    [Fact]
    public void OfficeChartDrawingRenderer_ReservesMeasuredHorizontalValueAxisLabelWidth() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Measured horizontal axis",
            "Measured Horizontal Axis",
            OfficeChartKind.BarClustered,
            new OfficeChartData(
                new[] { "Q1", "Q2" },
                new[] {
                    new OfficeChartSeries("Actual", new[] { 1200000D, 1800000D })
                }),
            widthPoints: 320D,
            heightPoints: 190D,
            layout: new OfficeChartLayout(horizontalAxisNumberFormat: "$#,##0.00")));

        OfficeDrawingText valueAxisLabel = drawing.Elements
            .OfType<OfficeDrawingText>()
            .First(label => label.Text == "$1,800,000.00");

        Assert.True(valueAxisLabel.Width > 60D, "Long formatted horizontal value-axis labels should not be forced into the legacy 34-point box.");
        Assert.True(valueAxisLabel.X >= 0D);
        Assert.True(valueAxisLabel.X + valueAxisLabel.Width <= drawing.Width);
    }

    [Fact]
    public void OfficeChartDrawingRenderer_AppliesCategoryAxisNumberFormatToNumericLabels() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Formatted categories",
            "Formatted Categories",
            OfficeChartKind.BarClustered,
            new OfficeChartData(
                new[] { "1", "2" },
                new[] {
                    new OfficeChartSeries("Actual", new[] { 12D, 18D })
                }),
            widthPoints: 320D,
            heightPoints: 190D,
            layout: new OfficeChartLayout(categoryAxisNumberFormat: "0.0")));

        string[] labels = drawing.Elements
            .OfType<OfficeDrawingText>()
            .Select(label => label.Text)
            .ToArray();

        Assert.Contains("1.0", labels);
        Assert.Contains("2.0", labels);
    }

    [Fact]
    public void OfficeChartDrawingRenderer_ScatterAxesIgnoreSkippedPointPairs() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Scatter skipped pairs",
            "Scatter Skipped Pairs",
            OfficeChartKind.Scatter,
            new OfficeChartData(
                new[] { "1", "999" },
                new[] {
                    new OfficeChartSeries("Points", new[] { 10D, double.NaN }, new[] { 1D, 999D })
                }),
            widthPoints: 320D,
            heightPoints: 190D,
            layout: new OfficeChartLayout(showLegend: false)));

        string[] labels = drawing.Elements
            .OfType<OfficeDrawingText>()
            .Select(label => label.Text)
            .ToArray();

        Assert.DoesNotContain("999", labels);
        Assert.DoesNotContain("NaN", labels);
    }

    [Fact]
    public void OfficeChartDrawingRenderer_ClampsColumnBaselineToExplicitVisibleAxisRange() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Visible axis minimum",
            "Visible Axis Minimum",
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(
                new[] { "Q1", "Q2" },
                new[] {
                    new OfficeChartSeries("Actual", new[] { 10D, 12D })
                }),
            widthPoints: 320D,
            heightPoints: 190D,
            layout: new OfficeChartLayout(verticalAxisMinimum: 5D, verticalAxisMaximum: 15D, showLegend: false)));

        OfficeDrawingShape[] filledRectangles = drawing.Elements
            .OfType<OfficeDrawingShape>()
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Rectangle && shape.Shape.FillColor.HasValue)
            .ToArray();

        Assert.NotEmpty(filledRectangles);
        Assert.All(filledRectangles, shape => Assert.True(
            shape.Y + shape.Shape.Height <= drawing.Height,
            "Column bars should stay inside the drawing when the visible value-axis range excludes zero."));
    }

    [Fact]
    public void OfficeChartDrawingRenderer_RendersMixedBarAndColumnSeriesWithAxisAlignedColumns() {
        OfficeColor columnColor = OfficeColor.ParseHex("#2563EB");
        OfficeColor barColor = OfficeColor.ParseHex("#DC2626");
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Mixed bar column",
            "Mixed Bar Column",
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(
                new[] { "Q1", "Q2" },
                new[] {
                    new OfficeChartSeries("Columns", new[] { 10D, 12D }, null, columnColor, null, true, renderKind: OfficeChartKind.ColumnClustered),
                    new OfficeChartSeries("Bars", new[] { 8D, 14D }, null, barColor, null, true, renderKind: OfficeChartKind.BarClustered)
                }),
            widthPoints: 340D,
            heightPoints: 220D,
            layout: new OfficeChartLayout(showLegend: false)));

        OfficeDrawingShape[] columnBars = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Rectangle && shape.Shape.FillColor == columnColor)
            .ToArray();
        OfficeDrawingShape[] axisAlignedBars = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Rectangle && shape.Shape.FillColor == barColor)
            .ToArray();

        Assert.NotEmpty(columnBars);
        Assert.NotEmpty(axisAlignedBars);
        Assert.All(columnBars, shape => Assert.True(shape.Shape.Height > shape.Shape.Width));
        Assert.All(axisAlignedBars, shape => Assert.True(shape.Shape.Height > shape.Shape.Width));
    }

    [Fact]
    public void OfficeChartDrawingRenderer_FiltersUnsupportedMixedScatterSeriesFromLegend() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            null,
            "Mixed Scatter Category Guard",
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(
                new[] { "Q1", "Q2" },
                new[] {
                    new OfficeChartSeries("Columns", new[] { 10D, 12D }, null, null, null, true, renderKind: OfficeChartKind.ColumnClustered),
                    new OfficeChartSeries("Scatter", new[] { 8D, 14D }, new[] { 1D, 2D }, null, null, true, renderKind: OfficeChartKind.Scatter)
                }),
            widthPoints: 340D,
            heightPoints: 220D));

        string[] text = drawing.Elements.OfType<OfficeDrawingText>().Select(label => label.Text).ToArray();

        Assert.Contains("Columns", text);
        Assert.DoesNotContain("Scatter", text);
    }

    [Fact]
    public void OfficeChartDrawingRenderer_UsesIndependentSecondaryValueAxisRange() {
        OfficeColor columnColor = OfficeColor.ParseHex("#2563EB");
        OfficeColor lineColor = OfficeColor.ParseHex("#DC2626");
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Secondary axis",
            "Secondary Axis",
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new OfficeChartSeries("Revenue", new[] { 100D, 150D, 200D }, null, columnColor, null,
                        showMarkers: false, renderKind: OfficeChartKind.ColumnClustered),
                    new OfficeChartSeries("Margin", new[] { 1D, 1.5D, 2D }, null, lineColor, null,
                        showMarkers: true, renderKind: OfficeChartKind.Line,
                        axisGroup: OfficeChartAxisGroup.Secondary)
                }),
            widthPoints: 360D,
            heightPoints: 220D,
            layout: new OfficeChartLayout(showLegend: false)));

        OfficeDrawingShape[] marginLines = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == lineColor)
            .ToArray();
        Assert.NotEmpty(marginLines);
        double verticalSpan = marginLines.Max(shape => shape.Y + shape.Shape.Height) -
                              marginLines.Min(shape => shape.Y);
        Assert.True(verticalSpan > 40D,
            "Secondary-axis line values should use their own 0-2 range rather than the primary 0-200 range.");
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), label => label.Text == "200");
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), label => label.Text == "2");
    }

    [Fact]
    public void OfficeChartDrawingRenderer_UsesIndependentSecondaryRangeForHorizontalBars() {
        OfficeColor primaryColor = OfficeColor.ParseHex("#2563EB");
        OfficeColor secondaryColor = OfficeColor.ParseHex("#DC2626");
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Horizontal secondary axis",
            "Horizontal Secondary Axis",
            OfficeChartKind.BarClustered,
            new OfficeChartData(
                new[] { "North", "South" },
                new[] {
                    new OfficeChartSeries("Primary", new[] { 100D, 200D }, null, primaryColor, null,
                        showMarkers: false, renderKind: OfficeChartKind.BarClustered),
                    new OfficeChartSeries("Secondary", new[] { 1D, 2D }, null, secondaryColor, null,
                        showMarkers: false, renderKind: OfficeChartKind.BarClustered,
                        axisGroup: OfficeChartAxisGroup.Secondary)
                }),
            widthPoints: 360D,
            heightPoints: 220D,
            layout: new OfficeChartLayout(showLegend: false)));

        OfficeDrawingShape[] secondaryBars = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                            shape.Shape.FillColor == secondaryColor)
            .ToArray();
        Assert.Equal(2, secondaryBars.Length);
        Assert.True(secondaryBars.Max(shape => shape.Shape.Width) > 100D,
            "Secondary horizontal bars should use their own 0-2 scale rather than the primary 0-200 scale.");
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), label => label.Text == "200");
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), label => label.Text == "2");
    }

    [Fact]
    public void OfficeChartDrawingRenderer_UsesAssignedAxisRangeForScatterSeries() {
        OfficeColor primaryColor = OfficeColor.ParseHex("#2563EB");
        OfficeColor secondaryColor = OfficeColor.ParseHex("#DC2626");
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Scatter axes",
            "Scatter Axes",
            OfficeChartKind.Scatter,
            new OfficeChartData(
                new[] { "1", "2" },
                new[] {
                    new OfficeChartSeries("Primary", new[] { 0D, 100D }, new[] { 1D, 2D },
                        primaryColor, null, showMarkers: true, renderKind: OfficeChartKind.Scatter),
                    new OfficeChartSeries("Secondary", new[] { 0D, 1D }, new[] { 1D, 2D },
                        secondaryColor, null, showMarkers: true, renderKind: OfficeChartKind.Scatter,
                        axisGroup: OfficeChartAxisGroup.Secondary)
                }),
            widthPoints: 360D,
            heightPoints: 220D,
            layout: new OfficeChartLayout(showLegend: false)));

        OfficeDrawingShape[] secondaryMarkers = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Ellipse &&
                            shape.Shape.FillColor == secondaryColor)
            .ToArray();
        Assert.Equal(2, secondaryMarkers.Length);
        double verticalSpan = secondaryMarkers.Max(shape => shape.Y) - secondaryMarkers.Min(shape => shape.Y);
        Assert.True(verticalSpan > 40D,
            "Secondary scatter values should use their own 0-1 range instead of the primary 0-100 range.");
    }
}
