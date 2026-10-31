using C = DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfPaperChartsTests {
    [Theory]
    [InlineData(2, 1D)]
    [InlineData(4, 1D)]
    [InlineData(32, 5D)]
    [InlineData(40, 5D)]
    [InlineData(100, 10D)]
    [InlineData(2000, 200D)]
    public void SetCountAxis_PdfProjection_PreservesZeroAndReadableIntegerTicks(int count, double interval) {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart("Outcome counts");
        chart.AddCategories(new[] { "Passed", "Failed" }.ToList());
        chart.AddBar("Checks", new[] { count, 1 }, OfficeColor.ParseHex("276749"));
        chart.SetCountAxis().SetCountAxis().SetPaperStyle();
        Assert.True(WordPdfConverterExtensions.TryCreateNativeWordChartSnapshot(chart, out var snapshot, out string? warning), warning);
        Assert.NotNull(snapshot);
        Assert.Equal(0D, snapshot!.Layout.VerticalAxisMinimum ?? snapshot.Layout.HorizontalAxisMinimum);
        Assert.Equal(interval, snapshot.Layout.VerticalAxisMajorUnit ?? snapshot.Layout.HorizontalAxisMajorUnit);
        Assert.True((snapshot.Layout.VerticalAxisMaximum ?? snapshot.Layout.HorizontalAxisMaximum) > count);
        Assert.Equal(OfficeChartDataLabelPosition.OutsideEnd, snapshot.Layout.DataLabelPosition);
        Assert.Equal("202428", snapshot.Style.TitleColor.ToRgbHex());
        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
        Assert.Contains("Outcome counts", PdfDocument.Load(new System.IO.MemoryStream(bytes)).ExtractText());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DataTable_WithPaperFrameOrCountAxisRetainsNativeSeries(bool countAxis) {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart("Evidence counts");
        chart.AddCategories(new[] { "Passed", "Failed" }.ToList());
        chart.AddBar("Checks", new[] { 4, 1 }, OfficeColor.Blue);
        chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!.AddChild(new C.DataTable(), true);
        if (countAxis) chart.SetCountAxis();
        else chart.SetPaperStyle();
        Assert.True(WordPdfConverterExtensions.TryCreateNativeWordChartSnapshot(chart, out var snapshot, out string? warning), warning);
        Assert.Equal(2, Assert.Single(snapshot!.Data.Series).Values.Count);
        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
        Assert.Contains("Evidence counts", PdfDocument.Load(new System.IO.MemoryStream(bytes)).ExtractText());
    }

    [Theory]
    [InlineData("clustered", true)]
    [InlineData("stacked", false)]
    [InlineData("percentStacked", false)]
    [InlineData("line", false)]
    [InlineData("area", false)]
    [InlineData("doughnut", false)]
    [InlineData("pie", true)]
    public void PaperCountLabels_OnlyRequestOutsideEndForSupportedGroups(string kind, bool outsideEnd) {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart("Counts");
        chart.AddCategories(new[] { "Passed", "Failed" }.ToList());
        if (kind == "pie") chart.AddPie("Passed", 4).AddPie("Failed", 1);
        else if (kind == "doughnut") chart.AddDoughnut("Passed", 4).AddDoughnut("Failed", 1);
        else if (kind == "line") chart.AddLine("Checks", new[] { 4, 1 }.ToList(), OfficeColor.Blue);
        else if (kind == "area") chart.AddArea("Checks", new[] { 4, 1 }.ToList(), OfficeColor.Blue);
        else {
            chart.AddBar("Checks", new[] { 4, 1 }, OfficeColor.Blue);
            chart.ChartPart!.ChartSpace!.Descendants<C.BarGrouping>().Single().Val = kind switch {
                "stacked" => C.BarGroupingValues.Stacked,
                "percentStacked" => C.BarGroupingValues.PercentStacked,
                _ => C.BarGroupingValues.Clustered
            };
        }
        chart.SetPaperStyle().SetPaperStyle();
        Assert.NotEmpty(chart.ChartPart!.ChartSpace!.Descendants<C.DataLabels>());
        Assert.All(chart.ChartPart.ChartSpace.Descendants<C.DataLabels>(), labels => {
            if (outsideEnd) Assert.Equal(C.DataLabelPositionValues.OutsideEnd, labels.GetFirstChild<C.DataLabelPosition>()!.Val!.Value);
            else Assert.Null(labels.GetFirstChild<C.DataLabelPosition>());
            Assert.True(labels.GetFirstChild<C.ShowValue>()!.Val!.Value);
            Assert.False(labels.GetFirstChild<C.ShowPercent>()!.Val!.Value);
        });
    }

    [Fact]
    public void SetPaperStyle_PiePdfProjection_RemovesChartSpaceBorder() {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart("Outcomes");
        chart.AddPie("Passed", 2);
        chart.AddPie("Failed", 1);
        chart.SetPaperStyle(percentageLabels: true);
        Assert.True(WordPdfConverterExtensions.TryCreateNativeWordChartSnapshot(chart, out var snapshot, out string? warning), warning);
        Assert.False(snapshot!.Style.ShowBorder);
        Assert.False(snapshot.Style.ShowBackground);
    }
}
