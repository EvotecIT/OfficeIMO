using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingChartAxisLayoutTests {
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
}
