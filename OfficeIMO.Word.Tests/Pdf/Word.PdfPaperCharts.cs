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
        byte[] bytes = document.ToPdfDocument(new WordToPdfOptions { IncludePageNumbers = false }).ToBytes();
        Assert.Contains("Outcome counts", PdfReadDocument.Open(bytes).ExtractText());
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
