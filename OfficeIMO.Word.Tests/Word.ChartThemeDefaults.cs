using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartThemeDefaultsTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void PointPaintUpdatesKeepNativeMarkerChoices(bool pointStyles, bool visible) {
        var original = new OfficeChartData(new[] { "A", "B", "C" },
            new[] { new OfficeSeries("Sales", new[] { 1D, 2D, 3D }) });
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartType.Line, original);
        chart.ChartPart!.ChartSpace!.Descendants<C.Marker>().Single().Symbol!.Val = visible
            ? C.MarkerStyleValues.Diamond : C.MarkerStyleValues.None;
        var series = pointStyles
            ? new OfficeSeries("Sales", new[] { 3D, 4D, 5D }).WithPointStyles(
                new OfficeChartPointStyle?[] { new(fillColor: OfficeColor.Green), null, null })
            : new OfficeSeries("Sales", new[] { 3D, 4D, 5D }, null, null,
                new OfficeColor?[] { OfficeColor.Green, null, null });
        var updated = new OfficeChartData(original.Categories, new[] { series });
        chart.SetData(OfficeChartType.Line, updated);
        using var stream = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(stream, false);
        C.Marker marker = saved.MainDocumentPart!.ChartParts.Single().ChartSpace!.Descendants<C.Marker>().Single();
        Assert.Equal(visible ? C.MarkerStyleValues.Diamond : C.MarkerStyleValues.None, marker.Symbol!.Val!.Value);
        C.DataPoint point = marker.Parent!.Elements<C.DataPoint>().Single();
        Assert.Equal("008000", point.Descendants<A.RgbColorModelHex>().First().Val!.Value);
    }

    [Fact]
    public void NewChartsWriteSharedThemeFontsAndGeometryIntoTheSavedDocument() {
        using var document = WordDocument.Create();
        document.AddChart(OfficeChartType.ColumnClustered,
            new OfficeChartData(new[] { "East", "West" }, new[] { new OfficeSeries("Sales", new[] { 30D, 42D }) }))
            .SetTitle("Sales by region");
        using var stream = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(stream, false);
        C.ChartSpace space = saved.MainDocumentPart!.ChartParts.Single().ChartSpace!;
        A.DefaultRunProperties font = space.GetFirstChild<C.TextProperties>()!.Descendants<A.DefaultRunProperties>().Single();
        Assert.Equal(OfficeTheme.Default.ChartStyle.AxisFont.Family, font.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
        Assert.Equal(OfficeTheme.Default.ChartStyle.BarGapPercent, space.Descendants<C.GapWidth>().Single().Val!.Value);
        Assert.Equal((int)(OfficeTheme.Default.ChartStyle.TitleFontSize.ToPoints() * 100D),
            space.Descendants<C.Title>().Single().GetFirstChild<C.TextProperties>()!.Descendants<A.DefaultRunProperties>().Single().FontSize!.Value);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(saved));
    }
}
