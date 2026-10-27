using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartThemeDefaultsTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void PointPaintUpdatesKeepNativeMarkerChoices(bool pointStyles, bool visible) {
        var original = new OfficeChartData(new[] { "A", "B", "C" },
            new[] { new OfficeChartSeries("Sales", new[] { 1D, 2D, 3D }) });
        using var document = PowerPointPresentation.Create();
        PowerPointChart chart = document.AddSlide().AddChart(OfficeChartKind.Line, original);
        chart.SetSeriesMarker(0, visible ? OfficeChartMarkerShape.Diamond : OfficeChartMarkerShape.None);
        var series = pointStyles
            ? new OfficeChartSeries("Sales", new[] { 3D, 4D, 5D }).WithPointStyles(
                new OfficeChartPointStyle?[] { new(fillColor: OfficeColor.Green), null, null })
            : new OfficeChartSeries("Sales", new[] { 3D, 4D, 5D }, null, null,
                new OfficeColor?[] { OfficeColor.Green, null, null });
        var updated = new OfficeChartData(original.Categories, new[] { series });
        chart.UpdateData(updated);
        using var stream = new MemoryStream(document.ToBytes());
        using PresentationDocument saved = PresentationDocument.Open(stream, false);
        C.Marker marker = saved.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace!
            .Descendants<C.Marker>().Single();
        Assert.Equal(visible ? C.MarkerStyleValues.Diamond : C.MarkerStyleValues.None, marker.Symbol!.Val!.Value);
        C.DataPoint point = marker.Parent!.Elements<C.DataPoint>().Single();
        Assert.Equal("008000", point.Descendants<A.RgbColorModelHex>().First().Val!.Value);
    }

    [Theory]
    [InlineData(false, 3, null, false)]
    [InlineData(false, 2, null, true)]
    [InlineData(true, 3, null, true)]
    [InlineData(false, 3, true, true)]
    [InlineData(true, 3, false, false)]
    public void LineMarkersFollowTheThemeUnlessExplicitlyChosen(bool classic, int count, bool? visible, bool expected) {
        OfficeTheme theme = classic ? OfficeTheme.Classic : OfficeTheme.Modern;
        double[] values = new[] { 1D, 2D, 3D }.Take(count).ToArray();
        OfficeChartSeries series = visible.HasValue
            ? new OfficeChartSeries("Sales", values, null, null, null, visible.Value)
            : new OfficeChartSeries("Sales", values).WithLegendVisibility(true);
        using var presentation = PowerPointPresentation.Create(new PowerPointCreateOptions { Theme = theme });
        PowerPointChart chart = presentation.AddSlide().AddChart(OfficeChartKind.Line,
            new OfficeChartData(new[] { "A", "B", "C" }.Take(count), new[] { series }));
        if (visible.HasValue) chart.UpdateData(new OfficeChartData(new[] { "A", "B", "C" }.Take(count),
            new[] { new OfficeChartSeries("Sales", values.Select(value => value + 1D)) }));
        using var stream = new MemoryStream(presentation.ToBytes());
        using PresentationDocument saved = PresentationDocument.Open(stream, false);
        C.Marker marker = saved.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace!
            .Descendants<C.LineChartSeries>().Single().GetFirstChild<C.Marker>()!;
        Assert.Equal(!expected, marker.Symbol!.Val!.Value == C.MarkerStyleValues.None);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(expected, snapshot.Data.Series.Single().ShowMarkers);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NewChartsUseThePresentationThemeAndSaveValidNativeParts(bool classic) {
        OfficeTheme theme = classic ? OfficeTheme.Classic : OfficeTheme.Modern;
        using var presentation = PowerPointPresentation.Create(new PowerPointCreateOptions { Theme = theme });
        PowerPointChart chart = presentation.AddSlide().AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "East", "West" }, new[] { new OfficeChartSeries("Sales", new[] { 30D, 42D }) }));
        chart.SetTitle("Sales by region");
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(theme.ChartStyle.AxisFont.Family, snapshot.Style.FontFamily);
        Assert.Equal(theme.Palette.Accent1, snapshot.Data.Series.Single().Color);
        using var stream = new MemoryStream(presentation.ToBytes());
        using PresentationDocument saved = PresentationDocument.Open(stream, false);
        C.ChartSpace space = saved.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace!;
        Assert.Equal(theme.ChartStyle.BarGapPercent, space.Descendants<C.GapWidth>().Single().Val!.Value);
        Assert.Equal(theme.ChartStyle.TitleFont.Family,
            space.Descendants<C.Title>().Single().GetFirstChild<C.TextProperties>()!.Descendants<A.LatinFont>().Single().Typeface!.Value);
        C.ChartShapeProperties grid = space.Descendants<C.MajorGridlines>().Single().GetFirstChild<C.ChartShapeProperties>()!;
        Assert.Equal(theme.ChartStyle.GridlineColor.Resolve(theme).ToRgbHex(), grid.Descendants<A.RgbColorModelHex>().Single().Val!.Value);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(saved));
    }
}
