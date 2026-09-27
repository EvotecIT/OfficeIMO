using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartAxisBindingsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ScatterUpdates_RejectConflictingLayerAxesBeforeNativeOrWorkbookMutation(bool shared) {
        using var document = PowerPointPresentation.Create();
        var data = new OfficeChartData(new[] { "1", "2" }, new[] {
            new OfficeChartSeries("First", new[] { 1d, 2d }, new[] { 1d, 2d }), new OfficeChartSeries("Second", new[] { 3d, 4d }, new[] { 1d, 2d }) });
        var chart = document.AddSlide().AddChart(OfficeChartKind.Scatter, data);
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        var plot = part.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var first = plot.GetFirstChild<C.ScatterChart>()!;
        var second = (C.ScatterChart)first.CloneNode(true);
        first.Elements<C.ScatterChartSeries>().Last().Remove(); second.Elements<C.ScatterChartSeries>().First().Remove();
        var horizontal = (C.ValueAxis)plot.Elements<C.ValueAxis>().First().CloneNode(true);
        var vertical = (C.ValueAxis)plot.Elements<C.ValueAxis>().Last().CloneNode(true);
        horizontal.AxisId!.Val = 400001; horizontal.CrossingAxis!.Val = 400002;
        vertical.AxisId!.Val = 400002; vertical.CrossingAxis!.Val = 400001;
        second.Elements<C.AxisId>().First().Val = 400001; second.Elements<C.AxisId>().Last().Val = 400002;
        plot.InsertAfter(second, first); plot.Append(horizontal, vertical);
        Assert.Empty(document.ValidateDocument());
        string before = part.ChartSpace.OuterXml;
        byte[] Workbook() { using var output = new MemoryStream(); using (var stream = part.GetPartsOfType<EmbeddedPackagePart>().Single().GetStream()) stream.CopyTo(output); return output.ToArray(); }
        byte[] workbookBefore = Workbook();
        Assert.Throws<NotSupportedException>(() => {
            if (shared) chart.UpdateData(data);
            else chart.UpdateData(new PowerPointScatterChartData(data.Series.Select(series => new PowerPointScatterChartSeries(series.Name, series.XValues!, series.Values))));
        });
        Assert.Equal(before, part.ChartSpace.OuterXml);
        Assert.Equal(workbookBefore, Workbook());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CategoryUpdate_PreservesLayerFormattingWithRepositionedPrimaryAxis(bool top) {
        using var document = PowerPointPresentation.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 1d, 2d }) });
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, data);
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        var plot = part.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var position = top ? C.AxisPositionValues.Top : C.AxisPositionValues.Right;
        plot.GetFirstChild<C.LineChart>()!.AddChild(new C.DataLabels(new C.ShowValue { Val = true }), true);
        plot.GetFirstChild<C.ValueAxis>()!.AxisPosition!.Val = position;
        chart.UpdateData(data);
        plot = part.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!;
        Assert.True(plot.GetFirstChild<C.LineChart>()!.GetFirstChild<C.DataLabels>()?.GetFirstChild<C.ShowValue>()?.Val?.Value);
        Assert.Equal(position, plot.GetFirstChild<C.ValueAxis>()!.AxisPosition!.Val!.Value);
        Assert.Empty(document.ValidateDocument());
    }
}
