using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartPhysicalLengthsTests {
    [Fact]
    public void ChartLengthsRetainNativePrecisionAndUnspecifiedStylesAfterReload() {
        using var bytes = new MemoryStream();
        using (var presentation = PowerPointDocument.Create()) {
            var chart = presentation.AddSlide().AddChart(OfficeChartType.Line,
                new OfficeChartData(new[] { "First", "Next" }, new[] {
                    new OfficeSeries("Revenue", new[] { 12.5, 19.75 }) }));
            chart.Width = OfficeLength.Emus(4750001);
            chart.Height = OfficeLength.Points(150.25);
            Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot initialSnapshot, out string? reason), reason);
            Assert.Equal(4750001, initialSnapshot.Width.ToEmus());
            Assert.Equal(1908175, initialSnapshot.Height.ToEmus());
            chart.SetTitle("Revenue")
                .SetTitleTextStyle(OfficeLength.Points(12.345), bold: true, fontName: "Aptos")
                .SetLegend(OfficeChartLegendPosition.Right)
                .SetLegendTextStyle(OfficeLength.Points(9.125), italic: true)
                .SetCategoryAxisTitle("Period")
                .SetCategoryAxisTitleTextStyle(OfficeLength.Points(10.125))
                .SetValueAxisLabelTextStyle(OfficeLength.Points(8.125))
                .SetValueAxisGridlines(lineWidth: OfficeLength.Emus(9526))
                .SetSeriesLineColor(0, "123456", OfficeLength.Emus(19051))
                .SetSeriesMarker(0, OfficeChartMarkerShape.Circle, OfficeLength.Points(8.5),
                    lineColor: "654321", lineWidth: OfficeLength.Emus(6351))
                .SetSeriesDataLabelTemplate("Revenue", new PowerPointChartDataLabelTemplate {
                    ShowValue = true,
                    FontSize = OfficeLength.Points(11.125),
                    LineColor = "234567",
                    LineWidth = OfficeLength.Emus(12701),
                    ShowLeaderLines = true,
                    LeaderLineWidth = OfficeLength.Emus(9526)
                });
            var native = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace;
            var series = native.Descendants<C.LineChartSeries>().Single();
            var outline = series.ChartShapeProperties!.GetFirstChild<A.Outline>()!;
            outline.CapType = A.LineCapValues.Round;
            outline.Append(new A.PresetDash { Val = A.PresetLineDashValues.Dash });
            presentation.Save(bytes);
        }

        bytes.Position = 0;
        using (var loaded = PowerPointDocument.Load(bytes)) {
            var chart = loaded.Slides[0].Charts.Single();
            Assert.Equal(4750001, chart.Width.ToEmus());
            Assert.Equal(1908175, chart.Height.ToEmus());
            chart.SetTitleTextStyle(color: "ABCDEF")
                .SetSeriesLineColor("Revenue", "ABCDEF")
                .SetSeriesMarker("Revenue", OfficeChartMarkerShape.Diamond);
            using var savedBytes = new MemoryStream();
            loaded.Save(savedBytes);
            savedBytes.Position = 0;
            using var saved = PresentationDocument.Open(savedBytes, false);
            var native = saved.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace;
            var title = native.Descendants<C.Title>().First();
            var titleStyle = title.Descendants<A.RunProperties>().Single();
            Assert.Equal(1235, titleStyle.FontSize!.Value);
            Assert.True(titleStyle.Bold!.Value);
            Assert.Equal("Aptos", titleStyle.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.Equal(913, native.Descendants<C.Legend>().Single()
                .Descendants<A.DefaultRunProperties>().Single().FontSize!.Value);
            Assert.Contains(native.Descendants<C.CategoryAxis>().Single().Descendants<A.RunProperties>(),
                props => props.FontSize?.Value == 1013);
            Assert.Contains(native.Descendants<C.ValueAxis>().Single().Descendants<A.DefaultRunProperties>(),
                props => props.FontSize?.Value == 813);
            Assert.Equal(9526, native.Descendants<C.MajorGridlines>().Single()
                .Descendants<A.Outline>().Single().Width!.Value);
            var series = native.Descendants<C.LineChartSeries>().Single();
            var outline = series.ChartShapeProperties!.GetFirstChild<A.Outline>()!;
            Assert.Equal(19051, outline.Width!.Value);
            Assert.Equal(A.LineCapValues.Round, outline.CapType!.Value);
            Assert.Equal(A.PresetLineDashValues.Dash, outline.GetFirstChild<A.PresetDash>()!.Val!.Value);
            var marker = series.GetFirstChild<C.Marker>()!;
            Assert.Equal((byte)9, marker.Size!.Val!.Value);
            Assert.Equal(C.MarkerStyleValues.Diamond, marker.Symbol!.Val!.Value);
            Assert.Equal(6351, marker.ChartShapeProperties!.GetFirstChild<A.Outline>()!.Width!.Value);
            var labels = series.GetFirstChild<C.DataLabels>()!;
            Assert.Equal(1113, labels.Descendants<A.DefaultRunProperties>().Single().FontSize!.Value);
            Assert.Equal(12701, labels.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!.Width!.Value);
            Assert.Equal(9526, labels.GetFirstChild<C.LeaderLines>()!.Descendants<A.Outline>().Single().Width!.Value);
        }
    }

    [Fact]
    public void TypedChartStyleGuardsPreserveOmittedValuesAndRejectInvalidLengths() {
        using var presentation = PowerPointDocument.Create();
        var chart = presentation.AddSlide().AddChart(OfficeChartType.Line,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Revenue", new[] { 1d }) }));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetTitleTextStyle(OfficeLength.Zero));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetLegendTextStyle(OfficeLength.Points(-1)));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetDataLabelTextStyle(OfficeLength.Points(4001)));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetValueAxisGridlines(lineWidth: OfficeLength.Zero));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetDataLabelLeaderLines(lineWidth: OfficeLength.Emus(-1)));
        Assert.Throws<ArgumentException>(() => chart.SetDataLabelShapeStyle(lineWidth: OfficeLength.Points(1), noLine: true));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetSeriesMarker(0, OfficeChartMarkerShape.Circle, OfficeLength.Zero));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetSeriesMarker("Revenue", OfficeChartMarkerShape.Circle, OfficeLength.Points(73)));
        chart.SetTitleTextStyle(fontSize: null).SetSeriesLineColor(0, "123456", OfficeLength.Zero);
        var outline = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace
            .Descendants<C.LineChartSeries>().Single().ChartShapeProperties!.GetFirstChild<A.Outline>()!;
        Assert.Equal(0, outline.Width!.Value);
    }
}
