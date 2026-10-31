using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartThemeMarkerTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AutomaticLineMarkersKeepIsolatedFinitePointsVisible(bool threeIslands) {
        double[] values = threeIslands
            ? new[] { double.NaN, 1D, double.PositiveInfinity, 2D, double.NegativeInfinity, 3D, double.NaN }
            : new[] { double.NaN, 5D, double.NaN };
        OfficeColor ink = OfficeColor.Parse("#55AA77");
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Sales", null,
            OfficeChartType.Line, new OfficeChartData(values.Select((_, index) => index.ToString()),
                new[] { new OfficeSeries("Sales", values) }), 320, 240,
            new OfficeChartStyle(palette: new[] { ink }), new OfficeChartLayout(showLegend: false)));
        Assert.Equal(threeIslands ? 3 : 1, drawing.Shapes.Count(shape => shape.Shape.FillColor == ink));
    }

    [Theory]
    [InlineData(3, null, 0)]
    [InlineData(2, null, 2)]
    [InlineData(3, true, 3)]
    [InlineData(2, false, 0)]
    public void LinePaintingUsesAutomaticMarkersOnlyForSparseSeries(int count, bool? visible, int expected) {
        OfficeColor ink = OfficeColor.Parse("#55AA77");
        double[] values = Enumerable.Range(1, count).Select(value => (double)value).ToArray();
        OfficeSeries series = visible.HasValue
            ? new OfficeSeries("Sales", values, null, null, null, visible.Value)
            : new OfficeSeries("Sales", values).WithLegendVisibility(true);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Sales", null,
            OfficeChartType.Line, new OfficeChartData(Enumerable.Range(1, count).Select(value => value.ToString()),
                new[] { series }), 320, 240, new OfficeChartStyle(palette: new[] { ink }),
            new OfficeChartLayout(showLegend: false)));

        Assert.Equal(expected, drawing.Shapes.Count(shape => shape.Shape.FillColor == ink));
    }
}
