using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartDataSourceIntegrityTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Update_RejectsIncompatibleWorkbookBeforeChangingChartOrPackage(int updatePath) {
        using PowerPointPresentation document = PowerPointPresentation.Create();
        OfficeChartKind kind = updatePath == 2 ? OfficeChartKind.Scatter : OfficeChartKind.Line;
        var data = new OfficeChartData(new[] { "1", "2" },
            new[] { new OfficeChartSeries("Old", new[] { 3d, 4d }, new[] { 1d, 2d }) });
        var chart = document.AddSlide().AddChart(kind, data);
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        part.DeletePart(part.GetPartsOfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>().Single());
        var legacy = part.AddEmbeddedPackagePart("application/vnd.ms-excel");
        byte[] payload = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        using (var content = new MemoryStream(payload)) legacy.FeedData(content);
        part.ChartSpace!.GetFirstChild<C.ExternalData>()!.Id = part.GetIdOfPart(legacy);
        string before = part.ChartSpace.OuterXml;
        Assert.Throws<System.NotSupportedException>(() => {
            if (updatePath == 0) chart.UpdateData(data);
            else if (updatePath == 1) chart.UpdateData(new PowerPointChartData(new[] { "A", "B" },
                new[] { new PowerPointChartSeries("Updated", new[] { 5d, 6d }) }));
            else chart.UpdateData(new PowerPointScatterChartData(new[] {
                new PowerPointScatterChartSeries("Updated", new[] { 1d, 2d }, new[] { 5d, 6d }) }));
        });
        Assert.Equal(before, part.ChartSpace.OuterXml);
        Assert.Equal("application/vnd.ms-excel", legacy.ContentType);
        using var preserved = new MemoryStream();
        using (Stream content = legacy.GetStream()) content.CopyTo(preserved);
        Assert.Equal(payload, preserved.ToArray());
    }

    [Fact]
    public void LegacyScatterUpdate_UsesSharedSchemaSafeSourceReplacement() {
        using PowerPointPresentation document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.Scatter, new OfficeChartData(new[] { "1", "2" },
            new[] { new OfficeChartSeries("Old", new[] { 3d, 4d }, new[] { 1d, 2d }) }));
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        C.ScatterChartSeries series = part.ChartSpace!.Descendants<C.ScatterChartSeries>().Single();
        series.AddChild(new C.SeriesText(new C.NumericValue { Text = "Literal title" }), true);
        series.RemoveAllChildren<C.XValues>();
        series.RemoveAllChildren<C.YValues>();
        series.AddChild(new C.Smooth { Val = false }, true);
        series.AddChild(new C.ScatterSerExtensionList(), true);
        Assert.Empty(document.ValidateDocument());
        chart.UpdateData(new PowerPointScatterChartData(new[] {
            new PowerPointScatterChartSeries("Updated", new[] { 1d, 2d }, new[] { 5d, 6d })
        }));
        Assert.Single(series.GetFirstChild<C.SeriesText>()!.ChildElements);
        Assert.Equal("extLst", series.LastChild!.LocalName);
        Assert.Empty(document.ValidateDocument());
        using var bytes = new MemoryStream(document.ToBytes());
        using PowerPointPresentation reopened = PowerPointPresentation.Load(bytes);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(OfficeChartKind.ColumnClustered)]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Area)]
    [InlineData(OfficeChartKind.Pie)]
    public void LegacyCategoryUpdate_ReplacesLiteralTitleAndNumericCategoryChoice(OfficeChartKind kind) {
        using PowerPointPresentation document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(kind, new OfficeChartData(new[] { "1", "2" },
            new[] { new OfficeChartSeries("Old", new[] { 3d, 4d }) }));
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        var series = part.ChartSpace!.Descendants<DocumentFormat.OpenXml.OpenXmlCompositeElement>().Single(element => element.LocalName == "ser");
        series.AddChild(new C.SeriesText(new C.NumericValue { Text = "Literal title" }), true);
        series.AddChild(new C.CategoryAxisData(new C.NumberLiteral(new C.FormatCode { Text = "General" },
            new C.PointCount { Val = 2 }, new C.NumericPoint { Index = 0, NumericValue = new C.NumericValue { Text = "1" } },
            new C.NumericPoint { Index = 1, NumericValue = new C.NumericValue { Text = "2" } })), true);
        Assert.Empty(document.ValidateDocument());
        chart.UpdateData(new PowerPointChartData(new[] { "A", "B" }, new[] { new PowerPointChartSeries("Updated", new[] { 5d, 6d }) }));
        Assert.Single(series.GetFirstChild<C.SeriesText>()!.ChildElements);
        Assert.Single(series.GetFirstChild<C.CategoryAxisData>()!.ChildElements);
        Assert.NotNull(series.GetFirstChild<C.CategoryAxisData>()!.GetFirstChild<C.StringReference>());
        Assert.Empty(document.ValidateDocument());
    }
}
