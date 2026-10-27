using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartAttachmentTests {
    [Theory]
    [InlineData("bar")]
    [InlineData("bar-array")]
    [InlineData("bar-list")]
    [InlineData("line")]
    [InlineData("line-array")]
    [InlineData("area")]
    [InlineData("area-array")]
    [InlineData("scatter")]
    [InlineData("radar")]
    [InlineData("bar3d")]
    [InlineData("line3d")]
    [InlineData("area3d")]
    public void AddedSeriesRetainsItsNativeIdentityAndSupportsRenameAndRemoval(string family) {
        using var document = WordDocument.Create();
        var chart = document.AddChart("Series");
        var axis = chart.AddCategories(new() { "One", "Two" });
        Assert.Same(chart, axis.Chart);
        var series = family switch {
            "bar" => chart.AddBar("Original", 2, OfficeColor.Blue),
            "bar-array" => chart.AddBar("Original", new[] { 2, 3 }, OfficeColor.Blue),
            "bar-list" => chart.AddBar("Original", new System.Collections.Generic.List<int> { 2, 3 }, OfficeColor.Blue),
            "line" => chart.AddLine("Original", new System.Collections.Generic.List<int> { 2, 3 }, OfficeColor.Blue),
            "line-array" => chart.AddChartLine<int>("Original", new[] { 2, 3 }, OfficeColor.Blue),
            "area" => chart.AddArea("Original", new System.Collections.Generic.List<int> { 2, 3 }, OfficeColor.Blue),
            "area-array" => chart.AddArea<int>("Original", new[] { 2, 3 }, OfficeColor.Blue),
            "scatter" => chart.AddScatter("Original", new() { 1, 2 }, new() { 2, 3 }, OfficeColor.Blue),
            "radar" => chart.AddRadar("Original", new System.Collections.Generic.List<int> { 2, 3 }, OfficeColor.Blue),
            "bar3d" => chart.AddBar3D("Original", new System.Collections.Generic.List<int> { 2, 3 }, OfficeColor.Blue),
            "line3d" => chart.AddLine3D("Original", new System.Collections.Generic.List<int> { 2, 3 }, OfficeColor.Blue),
            _ => chart.AddArea3D("Original", new System.Collections.Generic.List<int> { 2, 3 }, OfficeColor.Blue)
        };
        Assert.Same(chart, series.Chart);
        Assert.Equal("Original", series.Name);
        series.Name = "Renamed";
        Assert.Equal("Renamed", series.Name);
        Assert.Equal(2, series.Values[0]);
        using var loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.True(loaded.Charts.Single().TryGetSnapshot(out var snapshot));
        var detached = Assert.Single(snapshot.Data.Series);
        Assert.Null(detached.Chart);
        Assert.Equal("Renamed", detached.Name);
        Assert.Throws<InvalidOperationException>(() => detached.Name = "Cannot edit inspection data");
        series.Remove();
        Assert.Throws<InvalidOperationException>(() => series.Remove());
    }

    [Fact]
    public void CategorySettingsApplyToSubsequentSeriesAndLegendReplacementKeepsOneNativeLegend() {
        using var document = WordDocument.Create();
        var chart = document.AddChart("Categories");
        var axis = chart.AddChartAxisX(new() { "Before" });
        chart.AddBar("First", 1, OfficeColor.Blue);
        axis.Categories = new[] { "After" };
        chart.AddBar("Second", 2, OfficeColor.Red);
        var original = chart.AddLegend(OfficeChartLegendPosition.Right);
        var current = chart.AddLegend(OfficeChartLegendPosition.Bottom);
        current.Overlay = true;
        Assert.Same(chart, current.Chart);
        Assert.Throws<InvalidOperationException>(() => original.Position = OfficeChartLegendPosition.Top);
        using var package = WordprocessingDocument.Open(new MemoryStream(document.ToBytes()), false);
        var native = package.MainDocumentPart!.ChartParts.Single().ChartSpace;
        var legend = Assert.Single(native.Descendants<C.Legend>());
        Assert.Equal(C.LegendPositionValues.Bottom, legend.LegendPosition!.Val!.Value);
        Assert.True(legend.GetFirstChild<C.Overlay>()!.Val!.Value);
        Assert.Equal(new[] { "Before", "After" }, native.Descendants<C.CategoryAxisData>().Select(data => data.Descendants<C.StringPoint>().Single().NumericValue!.Text));
        var errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(package).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(error => error.Description)));
        current.Remove();
        Assert.Throws<InvalidOperationException>(() => current.Remove());
    }
}
