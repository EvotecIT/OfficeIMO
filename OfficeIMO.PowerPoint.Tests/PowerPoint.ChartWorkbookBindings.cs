using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartWorkbookBindingsTests {
    [Theory]
    [InlineData(OfficeChartType.ColumnClustered)]
    [InlineData(OfficeChartType.Line)]
    [InlineData(OfficeChartType.Area)]
    [InlineData(OfficeChartType.Doughnut)]
    public void LegacyGrowth_DoesNotCopyPointMetadataToAddedCategorySeries(OfficeChartType kind) {
        using var document = PowerPointDocument.Create();
        var chart = document.AddSlide().AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Original", new[] { 1d, 2d }) }));
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        var original = (DocumentFormat.OpenXml.OpenXmlCompositeElement)part.ChartSpace!.Descendants().Single(element => element.LocalName == "ser");
        original.AddChild(new C.DataPoint(new C.Index { Val = 0 }), true);
        original.AddChild(new C.DataLabels(new C.ShowValue { Val = true }), true);
        if (kind != OfficeChartType.Doughnut) original.AddChild(new C.ErrorBars(new C.ErrorBarType { Val = C.ErrorBarValues.Both },
            new C.ErrorBarValueType { Val = C.ErrorValues.FixedValue }, new C.ErrorBarValue { Val = 0.5 }), true);
        chart.UpdateData(new PowerPointChartData(new[] { "A", "B" }, new[] {
            new PowerPointChartSeries("Original", new[] { 3d, 4d }), new PowerPointChartSeries("Added", new[] { 5d, 6d }) }));
        var series = part.ChartSpace.Descendants().Where(element => element.LocalName == "ser").ToArray();
        Assert.Single(series[0].Elements<C.DataPoint>());
        Assert.Single(series[0].Elements<C.DataLabels>());
        Assert.Empty(series[1].Elements<C.DataPoint>());
        Assert.Empty(series[1].Elements<C.DataLabels>());
        Assert.Empty(series[1].Elements<C.ErrorBars>());
        Assert.Empty(document.ValidateDocument());
    }
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Updates_RejectIncompleteErrorBarCachesAcrossEntryPoints(int path) {
        using var document = PowerPointDocument.Create();
        var data = new OfficeChartData(new[] { "1", "2" }, new[] { new OfficeSeries("Status", new[] { 1d, 2d }, new[] { 1d, 2d }) });
        var chart = document.AddSlide().AddChart(path == 2 ? OfficeChartType.Scatter : OfficeChartType.Line, data);
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        var source = (DocumentFormat.OpenXml.OpenXmlCompositeElement)part.ChartSpace!.Descendants().First(element => element.LocalName == "ser");
        source.AddChild(new C.ErrorBars(new C.ErrorBarType { Val = C.ErrorBarValues.Both }, new C.ErrorBarValueType { Val = C.ErrorValues.Custom },
            new C.Plus(new C.NumberReference(new C.Formula { Text = "Sheet1!$Z$2:$Z$3" }, new C.NumberingCache(new C.PointCount { Val = 2 },
                new C.NumericPoint { Index = 0, NumericValue = new C.NumericValue { Text = "0.5" } })))), true);
        string before = part.ChartSpace.OuterXml;
        using var output = new MemoryStream();
        using (var stream = part.GetPartsOfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>().Single().GetStream()) stream.CopyTo(output);
        byte[] workbook = output.ToArray();
        Assert.Throws<NotSupportedException>(() => {
            if (path == 0) chart.UpdateData(data);
            else if (path == 1) chart.UpdateData(new PowerPointChartData(data.Categories, new[] { new PowerPointChartSeries("Updated", new[] { 3d, 4d }) }));
            else chart.UpdateData(new PowerPointScatterChartData(new[] { new PowerPointScatterChartSeries("Updated", new[] { 1d, 2d }, new[] { 3d, 4d }) }));
        });
        Assert.Equal(before, part.ChartSpace.OuterXml);
        using var updated = new MemoryStream();
        using (var stream = part.GetPartsOfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>().Single().GetStream()) stream.CopyTo(updated);
        Assert.Equal(workbook, updated.ToArray());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Updates_RejectExternalWorkbookLinksAcrossEntryPoints(int path) {
        using var document = PowerPointDocument.Create();
        var data = new OfficeChartData(new[] { "1", "2" }, new[] { new OfficeSeries("Status", new[] { 1d, 2d }, new[] { 1d, 2d }) });
        var chart = document.AddSlide().AddChart(path == 2 ? OfficeChartType.Scatter : OfficeChartType.Line, data);
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        part.DeletePart(part.GetPartsOfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>().Single());
        var link = part.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package", new Uri("https://example.test/data.xlsx"));
        part.ChartSpace!.GetFirstChild<C.ExternalData>()!.Id = link.Id;
        string before = part.ChartSpace.OuterXml;
        Assert.Throws<NotSupportedException>(() => {
            if (path == 0) chart.UpdateData(data);
            else if (path == 1) chart.UpdateData(new PowerPointChartData(data.Categories, new[] { new PowerPointChartSeries("Updated", new[] { 3d, 4d }) }));
            else chart.UpdateData(new PowerPointScatterChartData(new[] { new PowerPointScatterChartSeries("Updated", new[] { 1d, 2d }, new[] { 3d, 4d }) }));
        });
        Assert.Equal(before, part.ChartSpace.OuterXml);
        Assert.Empty(part.GetPartsOfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Updates_MaterializeCachedLinkedTitlesAcrossSupportedEntryPoints(int path) {
        using PowerPointDocument document = PowerPointDocument.Create();
        OfficeChartType kind = path == 2 ? OfficeChartType.Scatter : path == 3 ? OfficeChartType.ColumnClustered : OfficeChartType.Line;
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Status", new[] { 1d, 2d }, new[] { 1d, 2d }) });
        var chart = document.AddSlide().AddChart(kind, data);
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        if (path == 3) {
            var plot = part.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
            var bar = plot.GetFirstChild<C.BarChart>()!;
            var advanced = new C.Bar3DChart(bar.ChildElements.Select(child => child.CloneNode(true)));
            advanced.RemoveAllChildren<C.Overlap>();
            plot.ReplaceChild(advanced, bar);
            Assert.Equal(PowerPointImportedChartSupport.EditableWithProjectedRendering, chart.InspectImportedContent().Support);
        }
        part.ChartSpace!.GetFirstChild<C.Chart>()!.AddChild(new C.Title(new C.ChartText(new C.StringReference(
            new C.Formula { Text = "OldSheet!$Z$1" }, new C.StringCache(new C.PointCount { Val = 1 },
                new C.StringPoint { Index = 0, NumericValue = new C.NumericValue { Text = "Retained title" } })))), true);
        if (path == 0 || path == 3) chart.UpdateData(data);
        else if (path == 1) chart.UpdateData(new PowerPointChartData(data.Categories,
            new[] { new PowerPointChartSeries("Status", new[] { 3d, 4d }) }));
        else chart.UpdateData(new PowerPointScatterChartData(new[] {
            new PowerPointScatterChartSeries("Status", new[] { 1d, 2d }, new[] { 3d, 4d }) }));
        var title = part.ChartSpace.GetFirstChild<C.Chart>()!.Title!;
        Assert.Equal("Retained title", title.InnerText);
        Assert.Empty(title.Descendants<C.StringReference>());
        Assert.Empty(document.ValidateDocument());
        using var bytes = new MemoryStream(document.ToBytes());
        using var reopened = PowerPointDocument.Load(bytes);
        Assert.Equal("Retained title", reopened.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.GetFirstChild<C.Chart>()!.Title!.InnerText);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyUpdates_RejectMixedNativeLayersWithoutChangingChart(bool scatter) {
        using PowerPointDocument document = PowerPointDocument.Create();
        var data = new OfficeChartData(new[] { "1", "2" }, new[] { new OfficeSeries("Status", new[] { 1d, 2d }, new[] { 1d, 2d }) });
        var slide = document.AddSlide();
        var chart = slide.AddChart(scatter ? OfficeChartType.Scatter : OfficeChartType.Line, data);
        slide.AddChart(OfficeChartType.Pie, data);
        var part = slide.SlidePart.ChartParts.Single(item => !item.ChartSpace!.Descendants<C.PieChart>().Any());
        var plot = part.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var source = slide.SlidePart.ChartParts.Single(item => item.ChartSpace!.Descendants<C.PieChart>().Any()).ChartSpace!.Descendants<C.PieChartSeries>().Single();
        plot.AddChild(new C.Pie3DChart(source.CloneNode(true)), true);
        Assert.Empty(document.ValidateDocument());
        string before = part.ChartSpace.OuterXml;
        byte[] workbookBefore = Workbook(part);
        Assert.Throws<NotSupportedException>(() => {
            if (scatter) chart.UpdateData(new PowerPointScatterChartData(new[] {
                new PowerPointScatterChartSeries("Updated", new[] { 1d, 2d }, new[] { 3d, 4d }) }));
            else chart.UpdateData(new PowerPointChartData(data.Categories, new[] { new PowerPointChartSeries("Updated", new[] { 3d, 4d }) }));
        });
        Assert.Equal(before, part.ChartSpace.OuterXml);
        Assert.Equal(workbookBefore, Workbook(part));
    }

    private static byte[] Workbook(DocumentFormat.OpenXml.Packaging.ChartPart part) {
        using var output = new MemoryStream();
        using (Stream source = part.GetPartsOfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>().Single().GetStream()) source.CopyTo(output);
        return output.ToArray();
    }
}
