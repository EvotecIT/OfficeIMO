using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartProjectionCompletionTests {
    [Fact]
    public void SeparateStackedLayersSharingAxesDoNotFlattenIntoOneStack() {
        using var presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChart(OfficeChartKind.ColumnStacked,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("First", new[] { 1d, 2d })
            }));
        C.PlotArea plot = GetChartPart(presentation).ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        C.BarChart first = plot.GetFirstChild<C.BarChart>()!;
        C.BarChart second = (C.BarChart)first.CloneNode(true);
        second.Elements<C.BarChartSeries>().Single().Index!.Val = 1;
        second.Elements<C.BarChartSeries>().Single().Order!.Val = 1;
        plot.InsertAfter(second, first);

        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UncachedFormulaTitleDoesNotDisappearFromStaticProjection(bool axisTitle) {
        using var presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChart();
        ChartPart part = GetChartPart(presentation);
        C.Title title = new(new C.ChartText(new C.StringReference(
            new C.Formula("Sheet1!$A$1"))));
        if (axisTitle) part.ChartSpace!.Descendants<C.ValueAxis>().Single().AddChild(title, true);
        else part.ChartSpace!.GetFirstChild<C.Chart>()!.AddChild(title, true);

        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void NativeAxisVisibilityAndLabelsSurviveStaticProjection() {
        using var presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChart();
        C.ChartSpace native = GetChartPart(presentation).ChartSpace!;
        C.CategoryAxis category = native.Descendants<C.CategoryAxis>().Single();
        C.ValueAxis value = native.Descendants<C.ValueAxis>().Single();
        category.AddChild(new C.Delete { Val = true }, true);
        value.AddChild(new C.TickLabelPosition { Val = C.TickLabelPositionValues.None }, true);

        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.False(snapshot.Layout.ShowCategoryAxis);
        Assert.False(snapshot.Layout.ShowValueAxisLabels);
    }

    [Fact]
    public void RadarGridlinesLabelsAndFillFollowNativeAxes() {
        using var presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChart(OfficeChartKind.Radar,
            new OfficeChartData(new[] { "A", "B", "C" }, new[] {
                new OfficeChartSeries("Values", new[] { 1d, 2d, 3d })
            }));
        C.ChartSpace native = GetChartPart(presentation).ChartSpace!;
        C.CategoryAxis category = native.Descendants<C.CategoryAxis>().Single();
        C.ValueAxis value = native.Descendants<C.ValueAxis>().Single();
        category.GetFirstChild<C.MajorGridlines>()?.Remove();
        category.AddChild(new C.TickLabelPosition { Val = C.TickLabelPositionValues.None }, true);
        value.AddChild(new C.MajorGridlines(new C.ChartShapeProperties(
            new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "BADA55" })))), true);
        native.Descendants<C.RadarChart>().Single().RadarStyle!.Val = C.RadarStyleValues.Filled;

        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.False(snapshot.Layout.ShowCategoryAxisLabels);
        Assert.True(snapshot.Layout.FillRadarSeries);
        Assert.False(snapshot.Style.ShowCategoryGridLines);
        Assert.True(snapshot.Style.ShowValueGridLines);
        Assert.Equal(OfficeColor.FromRgb(0xBA, 0xDA, 0x55), snapshot.Style.ValueGridLineColor);
    }

    private static ChartPart GetChartPart(PowerPointPresentation presentation) =>
        presentation.Slides.Single().SlidePart.ChartParts.Single();
}
