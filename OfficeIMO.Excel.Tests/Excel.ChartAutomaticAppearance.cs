using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Pins automatic Excel chart appearance to Microsoft Excel exports.</summary>
public sealed class ExcelChartAutomaticAppearanceTests {
    private static string FixtureFolder => Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "AutomaticColors");

    public static IEnumerable<object[]> ReferenceCharts() => ReadReference().Select(chart => new object[] { chart.Number });

    [Theory]
    [MemberData(nameof(ReferenceCharts))]
    public void AutomaticColorsMatchTheExcelExport(int number) {
        using ExcelDocument document = ExcelDocument.Load(Path.Combine(FixtureFolder, "automatic-colors.xlsx"));
        ExcelChart chart = document.Sheets.Single(sheet => sheet.Name == "Chart" + number).Charts.Single();

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        var fills = new List<string>();
        if (snapshot.ChartType is ExcelChartType.Pie or ExcelChartType.Doughnut) {
            fills.AddRange(snapshot.Style!.Palette.Take(snapshot.Data.Categories.Count).Select(color => color.ToRgbHex()));
        } else {
            bool lines = snapshot.ChartType is ExcelChartType.Line or ExcelChartType.Scatter or ExcelChartType.Stock;
            fills.AddRange(snapshot.Data.Series.Where(series => !lines || series.ShowMarkers)
                .Select(series => series.SeriesColorArgb!.ToUpperInvariant()));
        }
        Assert.Equal(ReadReference().Single(chart => chart.Number == number).Fills, fills.Distinct().ToArray());
    }

    [Fact]
    public void StockChartExportsAsALineProjectionWithHighLowLines() {
        using ExcelDocument document = ExcelDocument.Load(Path.Combine(FixtureFolder, "automatic-colors.xlsx"));
        int stock = ReadReference().Single(chart => chart.Title == "Stock high-low-close").Number;
        Assert.True(document.Sheets.Single(sheet => sheet.Name == "Chart" + stock).Charts.Single()
            .TryGetSnapshot(out ExcelChartInfo snapshot));
        Assert.True(snapshot.Layout!.ShowHighLowLines);

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence();

        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "worksheet-chart");
        Assert.Contains(result.Warnings, warning => warning.Code == "chart-approximation" &&
            warning.Message.Contains("high-low lines", StringComparison.Ordinal));
    }

    [Fact]
    public void ExcelProducedPieLegendLabelsStayVisibleBesideTheNextChart() {
        // The pie's legend sits close to the doughnut chart anchored beside it. A chart frame
        // wider than its anchor used to paint over the legend labels.
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "Excel", "exploded-slices.xlsx");
        using ExcelDocument document = ExcelDocument.Load(path);
        byte[] pdf = document.ConvertToPdfWithEvidence().ToBytes();

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));
        using var reader = UglyToad.PdfPig.PdfDocument.Open(pdf);
        var page = reader.GetPage(1);
        var labels = page.GetWords().Where(word => word.Text is "Complete" or "Pending").ToArray();
        Assert.True(labels.Length >= 4, "Both chart legends should name both slices.");
        double scaleX = raster.Width / page.Width, scaleY = raster.Height / page.Height;
        foreach (var label in labels) {
            int left = (int)Math.Floor(label.BoundingBox.Left * scaleX), right = (int)Math.Ceiling(label.BoundingBox.Right * scaleX);
            int top = (int)Math.Floor((page.Height - label.BoundingBox.Top) * scaleY);
            int bottom = (int)Math.Ceiling((page.Height - label.BoundingBox.Bottom) * scaleY);
            int ink = 0;
            for (int y = Math.Max(0, top); y < Math.Min(raster.Height, bottom); y++)
                for (int x = Math.Max(0, left); x < Math.Min(raster.Width, right); x++) {
                    OfficeColor pixel = raster.GetPixel(x, y);
                    if (pixel.R + pixel.G + pixel.B < 3 * 160) ink++;
                }
            Assert.True(ink > 0, $"The label '{label.Text}' at {label.BoundingBox.Left:0.#} has no visible ink.");
        }
    }

    private static IReadOnlyList<(int Number, string Title, string[] Fills)> ReadReference() {
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureFolder, "office-reference.json")));
        return document.RootElement.GetProperty("charts").EnumerateArray().Select(chart => (
            chart.GetProperty("number").GetInt32(),
            chart.GetProperty("title").GetString()!,
            chart.GetProperty("officeFills").EnumerateArray().Select(fill => fill.GetString()!).ToArray())).ToArray();
    }
}
