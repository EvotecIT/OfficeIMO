using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartDataSourceIntegrityTests {
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
