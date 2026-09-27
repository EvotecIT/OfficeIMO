using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Tests;

public sealed class WordChartImportedPreservationTests {
    [Fact]
    public void SharedUpdate_RejectsRepeatedLayersWithDifferentAxisPairsBeforeMutation() {
        using var document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("First", new[] { 1d, 2d }), new OfficeChartSeries("Second", new[] { 3d, 4d }) });
        var chart = document.AddChart(OfficeChartKind.Line, data);
        var part = chart.ChartPart!;
        var plot = part.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var first = plot.GetFirstChild<C.LineChart>()!;
        var second = (C.LineChart)first.CloneNode(true);
        first.Elements<C.LineChartSeries>().Last().Remove(); second.Elements<C.LineChartSeries>().First().Remove();
        var category = (C.CategoryAxis)plot.Elements<C.CategoryAxis>().Single().CloneNode(true);
        var value = (C.ValueAxis)plot.Elements<C.ValueAxis>().Single().CloneNode(true);
        category.AxisId!.Val = 400001; category.CrossingAxis!.Val = 400002;
        value.AxisId!.Val = 400002; value.CrossingAxis!.Val = 400001;
        second.Elements<C.AxisId>().First().Val = 400001; second.Elements<C.AxisId>().Last().Val = 400002;
        plot.InsertAfter(second, first); plot.Append(category, value);
        Assert.Empty(document.ValidateDocument());
        string before = part.ChartSpace.OuterXml;
        byte[] Workbook() {
            using var output = new System.IO.MemoryStream();
            using (var stream = part.GetPartsOfType<EmbeddedPackagePart>().Single().GetStream()) stream.CopyTo(output);
            return output.ToArray();
        }
        byte[] workbookBefore = Workbook();
        Assert.Throws<NotSupportedException>(() => chart.SetData(OfficeChartKind.Line, data));
        Assert.Equal(before, part.ChartSpace.OuterXml);
        Assert.Equal(workbookBefore, Workbook());
    }

    [Fact]
    public void SharedUpdate_PreservesAxesByTheirChartReferencesWhenXmlOrderChanges() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, Combo());
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var primary = plot.Elements<C.ValueAxis>().Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Left);
        var secondary = plot.Elements<C.ValueAxis>().Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right);
        primary.AddChild(Title("Primary"), true); secondary.AddChild(Title("Secondary"), true);
        primary.Scaling!.AddChild(new C.MaxAxisValue { Val = 10 }, true);
        secondary.Scaling!.AddChild(new C.MaxAxisValue { Val = 100 }, true);
        secondary.Remove(); plot.InsertBefore(secondary, primary);
        var categories = plot.Elements<C.CategoryAxis>().ToArray();
        categories[1].Remove(); plot.InsertBefore(categories[1], categories[0]);
        Assert.Empty(document.ValidateDocument());
        chart.SetData(OfficeChartKind.ColumnClustered, Combo());
        plot = chart.ChartPart.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!;
        uint secondaryId = plot.GetFirstChild<C.LineChart>()!.Elements<C.AxisId>().Last().Val!.Value;
        secondary = plot.Elements<C.ValueAxis>().Single(axis => axis.AxisId!.Val!.Value == secondaryId);
        Assert.Equal("Secondary", secondary.GetFirstChild<C.Title>()!.InnerText);
        Assert.Equal(100d, secondary.Scaling!.GetFirstChild<C.MaxAxisValue>()!.Val!.Value);
        Assert.Equal(C.AxisPositionValues.Right, secondary.AxisPosition!.Val!.Value);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SharedUpdate_PreservesSecondaryLayerWhenItsCategoryAxisIsVisible() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, Combo());
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        plot.Elements<C.CategoryAxis>().Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Top).Delete!.Val = false;
        plot.GetFirstChild<C.LineChart>()!.AddChild(new C.DataLabels(new C.ShowValue { Val = true }), true);
        chart.SetData(OfficeChartKind.ColumnClustered, Combo());
        Assert.True(chart.ChartPart.ChartSpace.Descendants<C.LineChart>().Single().GetFirstChild<C.DataLabels>()!.GetFirstChild<C.ShowValue>()!.Val!.Value);
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void SharedUpdate_PreservesSeparateNativeLayersSharingFamilyAndAxes(int updatedSeriesCount) {
        using var document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("First", new[] { 1d, 2d }), new OfficeChartSeries("Second", new[] { 3d, 4d }) });
        var chart = document.AddChart(OfficeChartKind.Line, data);
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var first = plot.GetFirstChild<C.LineChart>()!;
        var second = (C.LineChart)first.CloneNode(true);
        first.Elements<C.LineChartSeries>().Last().Remove(); second.Elements<C.LineChartSeries>().First().Remove();
        first.AddChild(new C.DataLabels(new C.ShowValue { Val = true }), true);
        second.AddChild(new C.DataLabels(new C.ShowValue { Val = false }), true);
        plot.InsertAfter(second, first);
        Assert.Empty(document.ValidateDocument());
        var updated = new OfficeChartData(data.Categories, Enumerable.Range(0, updatedSeriesCount).Select(index =>
            new OfficeChartSeries(index == 1 ? "Second" : "Series " + index, new[] { index + 10d, index + 20d })));
        chart.SetData(OfficeChartKind.Line, updated);
        var layers = chart.ChartPart.ChartSpace.Descendants<C.LineChart>().ToArray();
        Assert.Equal(Math.Min(2, updatedSeriesCount), layers.Length);
        Assert.True(layers[0].GetFirstChild<C.DataLabels>()!.GetFirstChild<C.ShowValue>()!.Val!.Value);
        if (updatedSeriesCount > 1) {
            Assert.False(layers[1].GetFirstChild<C.DataLabels>()!.GetFirstChild<C.ShowValue>()!.Val!.Value);
            Assert.Equal("Second", layers[1].Descendants<C.SeriesText>().First().Descendants<C.StringPoint>().Single().NumericValue!.Text);
        }
        Assert.Equal(updatedSeriesCount, layers.Sum(layer => layer.Elements<C.LineChartSeries>().Count()));
        Assert.Empty(document.ValidateDocument());
        using var bytes = new System.IO.MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var reopened = WordDocument.Load(bytes);
        Assert.Equal(layers.Length, reopened.Charts.Single().ChartPart!.ChartSpace!.Descendants<C.LineChart>().Count());
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData("standard")]
    [InlineData("filled")]
    public void SharedUpdate_PreservesImportedRadarStyle(string style) {
        using var document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Series", new[] { 1d, 2d }) });
        var chart = document.AddChart(OfficeChartKind.Radar, data);
        chart.ChartPart!.ChartSpace!.Descendants<C.RadarStyle>().Single().Val = style == "filled" ? C.RadarStyleValues.Filled : C.RadarStyleValues.Standard;
        chart.SetData(OfficeChartKind.Radar, data);
        Assert.Equal(style, chart.ChartPart.ChartSpace.Descendants<C.RadarStyle>().Single().Val!.InnerText);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SharedUpdate_RejectsAnExternallyLinkedWorkbookBeforeChangingNativeData() {
        using var document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Series", new[] { 1d, 2d }) });
        var chart = document.AddChart(OfficeChartKind.Line, data);
        var part = chart.ChartPart!;
        part.DeletePart(part.GetPartsOfType<EmbeddedPackagePart>().Single());
        var link = part.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package", new Uri("https://example.test/data.xlsx"));
        part.ChartSpace!.GetFirstChild<C.ExternalData>()!.Id = link.Id;
        string before = part.ChartSpace.OuterXml;
        Assert.Throws<NotSupportedException>(() => chart.SetData(OfficeChartKind.Line, data));
        Assert.Equal(before, part.ChartSpace.OuterXml);
        Assert.Empty(part.GetPartsOfType<EmbeddedPackagePart>());
        Assert.Equal(link.Id, part.ChartSpace.GetFirstChild<C.ExternalData>()!.Id!.Value);
    }

    [Fact]
    public void SharedUpdate_PreservesFormattedLegendEntriesWhileUpdatingDeletion() {
        using var document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Series", new[] { 1d, 2d }) });
        var chart = document.AddChart(OfficeChartKind.Line, data);
        var legend = chart.ChartPart!.ChartSpace!.Descendants<C.Legend>().Single();
        var text = new C.TextProperties(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties { FontSize = 1800 })));
        legend.AddChild(new C.LegendEntry(new C.Index { Val = 0 }, text), true);
        Assert.Empty(document.ValidateDocument());
        chart.SetData(OfficeChartKind.Line, data);
        var entry = chart.ChartPart.ChartSpace.Descendants<C.LegendEntry>().Single();
        Assert.Equal(text.OuterXml, entry.GetFirstChild<C.TextProperties>()!.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }

    private static C.Title Title(string value) => new C.Title(new C.ChartText(new C.RichText(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text(value))))));
    private static OfficeChartData Combo() => new OfficeChartData(new[] { "A", "B" }, new[] {
        new OfficeChartSeries("Counts", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.ColumnClustered),
        new OfficeChartSeries("Rates", new[] { 10d, 20d }, null, null, null, true, renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary) });
}
