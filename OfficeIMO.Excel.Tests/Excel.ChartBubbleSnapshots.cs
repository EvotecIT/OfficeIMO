using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartBubbleSnapshotTests {
    [Fact]
    public void BubbleSnapshot_ReopenedWorksheetRangesPreserveEachSeriesCoordinates() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Bubble data");
        sheet.CellValues(new[] {
            (2, 1, (object)2d), (2, 2, 11d), (2, 3, 4d),
            (3, 1, 8d), (3, 2, 17d), (3, 3, 16d),
            (2, 5, -3d), (2, 6, 23d), (2, 7, 9d)
        }, null);
        sheet.AddBubbleChartFromRanges(new[] {
            new ExcelChartSeriesRange("First", "A2:A3", "B2:B3", "C2:C3"),
            new ExcelChartSeriesRange("Second", "E2:E2", "F2:F2", "G2:G2")
        }, 8, 4, 480, 320, "Bubble snapshot");
        using var stream = new MemoryStream(document.ToBytes());
        using ExcelDocument reopened = ExcelDocument.Load(stream);
        ExcelChart chart = reopened.Sheets.Single().Charts.Single();

        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(ExcelChartType.Bubble, snapshot.ChartType);
        Assert.Equal("Bubble snapshot", snapshot.Title);
        Assert.Equal(2, snapshot.Data.Series.Count);
        Assert.Equal(new[] { 2d, 8d }, snapshot.Data.Series[0].XValues);
        Assert.Equal(new[] { 11d, 17d }, snapshot.Data.Series[0].Values);
        Assert.Equal(new[] { 4d, 16d }, snapshot.Data.Series[0].BubbleSizes);
        Assert.Equal(new[] { -3d }, snapshot.Data.Series[1].XValues);
        Assert.Equal(new[] { 23d }, snapshot.Data.Series[1].Values);
        Assert.Equal(new[] { 9d }, snapshot.Data.Series[1].BubbleSizes);
        Assert.False(snapshot.Data.Series[0].ConnectLine);
        Assert.True(snapshot.Data.Series[0].ShowMarkers);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BubbleSnapshot_LiteralAndCacheSourcesPreservePointIndices(bool cached) {
        using ExcelDocument document = ExcelDocument.Create();
        CreateLiteralChart(document, cached);
        using var stream = new MemoryStream(document.ToBytes());
        using ExcelDocument reopened = ExcelDocument.Load(stream);
        ExcelChart chart = reopened.Sheets.Single().Charts.Single();
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        ExcelChartSeries series = Assert.Single(snapshot.Data.Series);
        Assert.Equal("Cached bubbles", series.Name);
        Assert.Equal(new[] { 2d, 8d }, series.XValues);
        Assert.Equal(new[] { 11d, 17d }, series.Values);
        Assert.Equal(new[] { 4d, 16d }, series.BubbleSizes);
        Assert.False(series.ConnectLine);
        Assert.True(series.ShowMarkers);

        // Exercise the existing Excel render adapter without changing its scatter approximation.
        OfficeImageExportResult image = chart.ExportImage(OfficeImageExportFormat.Svg);
        Assert.True(image.EncodedLength > 0);
        Assert.Contains(image.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartKindApproximated);
    }

    [Fact]
    public void BubbleSnapshot_SeriesOrderKeepsCoordinatesSizesAndStylesTogether() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelChart chart = CreateLiteralChart(document, false);
        C.BubbleChart native = GetNativeBubble(document);
        C.BubbleChartSeries first = native.Elements<C.BubbleChartSeries>().Single();
        first.GetFirstChild<C.Order>()!.Val = 1U;
        first.InsertAfter(new C.ChartShapeProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })),
            first.GetFirstChild<C.SeriesText>()!);
        var second = (C.BubbleChartSeries)first.CloneNode(true);
        second.GetFirstChild<C.Index>()!.Val = 1U;
        second.GetFirstChild<C.Order>()!.Val = 0U;
        second.GetFirstChild<C.SeriesText>()!.GetFirstChild<C.NumericValue>()!.Text = "Draw first";
        second.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val = "0000FF";
        second.GetFirstChild<C.XValues>()!.Descendants<C.NumericPoint>().Single(p => p.Index!.Value == 0).NumericValue!.Text = "-7";
        second.GetFirstChild<C.BubbleSize>()!.Descendants<C.NumericPoint>().Single(p => p.Index!.Value == 0).NumericValue!.Text = "25";
        native.InsertBefore(second, first);

        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal("Draw first", snapshot.Data.Series[0].Name);
        Assert.Equal(new[] { -7d, 8d }, snapshot.Data.Series[0].XValues);
        Assert.Equal(new[] { 25d, 16d }, snapshot.Data.Series[0].BubbleSizes);
        Assert.Equal("0000FF", snapshot.Data.Series[0].SeriesColorArgb);
        Assert.Equal("FF0000", snapshot.Data.Series[1].SeriesColorArgb);
    }

    [Theory]
    [InlineData("missing-size")]
    [InlineData("mismatched-size")]
    [InlineData("duplicate-index")]
    [InlineData("sparse-index")]
    [InlineData("nonfinite-x")]
    [InlineData("negative-size")]
    [InlineData("excessive-count")]
    public void BubbleSnapshot_InvalidTriplesFailWithoutPartialSnapshot(string defect) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelChart chart = CreateLiteralChart(document, false);
        C.BubbleChartSeries series = GetNativeBubble(document).Elements<C.BubbleChartSeries>().Single();
        C.NumberLiteral sizes = series.GetFirstChild<C.BubbleSize>()!.GetFirstChild<C.NumberLiteral>()!;
        switch (defect) {
            case "missing-size": series.RemoveAllChildren<C.BubbleSize>(); break;
            case "mismatched-size":
                sizes.Elements<C.NumericPoint>().First().Remove();
                sizes.GetFirstChild<C.PointCount>()!.Val = 1U;
                sizes.Elements<C.NumericPoint>().Single().Index = 0U;
                break;
            case "duplicate-index": sizes.Elements<C.NumericPoint>().First().Index = 0U; break;
            case "sparse-index": sizes.GetFirstChild<C.PointCount>()!.Val = 3U; break;
            case "nonfinite-x": series.GetFirstChild<C.XValues>()!.Descendants<C.NumericValue>().First().Text = "NaN"; break;
            case "negative-size": sizes.Elements<C.NumericPoint>().First().NumericValue!.Text = "-1"; break;
            case "excessive-count": sizes.GetFirstChild<C.PointCount>()!.Val = 1_000_001U; break;
        }
        Assert.False(chart.TryGetSnapshot(out _));
    }

    private static ExcelChart CreateLiteralChart(ExcelDocument document, bool cached) {
        ExcelSheet sheet = document.AddWorksheet("Bubble data");
        ExcelChart chart = sheet.AddBubbleChartFromRanges(new[] {
            new ExcelChartSeriesRange("Unused", "A2:A3", "B2:B3", "C2:C3")
        }, 8, 4, 480, 320, "Bubble snapshot");
        C.BubbleChart native = GetNativeBubble(document);
        C.BubbleChartSeries series = native.Elements<C.BubbleChartSeries>().Single();
        series.RemoveAllChildren<C.SeriesText>();
        series.InsertAfter(cached
            ? new C.SeriesText(new C.StringReference(new C.Formula("'Missing sheet'!$A$1"),
                new C.StringCache(new C.PointCount { Val = 1U },
                    new C.StringPoint(new C.NumericValue("Cached bubbles")) { Index = 0U })))
            : new C.SeriesText(new C.NumericValue("Cached bubbles")), series.GetFirstChild<C.Order>()!);
        series.RemoveAllChildren<C.XValues>();
        series.RemoveAllChildren<C.YValues>();
        series.RemoveAllChildren<C.BubbleSize>();
        series.Append(new C.XValues(Numbers(cached, 2, 8)),
            new C.YValues(Numbers(cached, 11, 17)),
            new C.BubbleSize(Numbers(cached, 4, 16)));
        return chart;
    }

    private static OpenXmlCompositeElement Numbers(bool cached, params double[] values) {
        OpenXmlCompositeElement points = cached ? new C.NumberingCache() : new C.NumberLiteral();
        points.Append(new C.FormatCode("General"), new C.PointCount { Val = (uint)values.Length });
        for (int index = values.Length - 1; index >= 0; index--)
            points.Append(new C.NumericPoint(new C.NumericValue(values[index].ToString(CultureInfo.InvariantCulture))) { Index = (uint)index });
        return cached ? new C.NumberReference(new C.Formula("'Missing sheet'!$A$2:$A$3"), points) : points;
    }

    private static C.BubbleChart GetNativeBubble(ExcelDocument document) =>
        document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single()
            .ChartSpace.Descendants<C.BubbleChart>().Single();
}
