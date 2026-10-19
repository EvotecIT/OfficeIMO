using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Pins automatic Word chart appearance to Microsoft Word exports.</summary>
public sealed class WordChartAutomaticAppearanceTests {
    private static string FixtureFolder => Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "AutomaticColors");

    public static IEnumerable<object[]> ReferenceCharts() => ReadReference().Select(chart => new object[] { chart.Number });

    [Theory]
    [MemberData(nameof(ReferenceCharts))]
    public void SharedSnapshotColorsMatchTheWordExport(int number) {
        using WordDocument document = WordDocument.Load(Path.Combine(FixtureFolder, "automatic-colors.docx"));
        WordChart chart = document.Charts[number - 1];

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));

        Assert.Equal(ReadReference().Single(item => item.Number == number).Fills, DrawnFills(snapshot));
    }

    [Theory]
    [MemberData(nameof(ReferenceCharts))]
    public void PdfChartColorsMatchTheWordExport(int number) {
        using WordDocument document = WordDocument.Load(Path.Combine(FixtureFolder, "automatic-colors.docx"));
        WordChart chart = document.Charts[number - 1];

        Assert.True(WordPdfConverterExtensions.TryCreateNativeWordChartSnapshot(chart, out OfficeChartSnapshot? snapshot, out string? warning), warning);

        Assert.Equal(ReadReference().Single(item => item.Number == number).Fills, DrawnFills(snapshot!));
    }

    [Fact]
    public void EveryAutomaticChartConvertsToPdf() {
        using WordDocument document = WordDocument.Load(Path.Combine(FixtureFolder, "automatic-colors.docx"));
        OfficeIMO.Pdf.PdfDocumentConversionResult result = document.ToPdfDocumentResult();

        Assert.DoesNotContain(document.Charts, chart => !chart.TryGetOfficeSnapshot(out _));
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
    }

    // The fills a chart draws in series order: filled series and visible markers per series,
    // and slices per point for radial charts.
    private static IReadOnlyList<string> DrawnFills(OfficeChartSnapshot snapshot) {
        var fills = new List<string>();
        void Add(OfficeColor color) {
            string hex = color.ToRgbHex();
            if (!fills.Contains(hex)) fills.Add(hex);
        }
        if (snapshot.ChartKind is OfficeChartKind.Pie or OfficeChartKind.Doughnut) {
            OfficeChartSeries ring = snapshot.Data.Series[0];
            for (int index = 0; index < snapshot.Data.Categories.Count; index++)
                Add(ring.PointColors?.ElementAtOrDefault(index) ?? snapshot.Style.GetSeriesColor(index));
            return fills;
        }
        bool lines = snapshot.ChartKind is OfficeChartKind.Line or OfficeChartKind.Scatter;
        for (int index = 0; index < snapshot.Data.Series.Count; index++) {
            OfficeChartSeries series = snapshot.Data.Series[index];
            if (lines && (!series.ShowMarkers || series.MarkerShape == OfficeChartMarkerShape.None)) continue;
            Add(series.Color ?? snapshot.Style.GetSeriesColor(index));
        }
        return fills;
    }

    private static IReadOnlyList<(int Number, string Title, string[] Fills)> ReadReference() {
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureFolder, "office-reference.json")));
        return document.RootElement.GetProperty("charts").EnumerateArray().Select(chart => (
            chart.GetProperty("number").GetInt32(),
            chart.GetProperty("title").GetString()!,
            chart.GetProperty("officeFills").EnumerateArray().Select(fill => fill.GetString()!).ToArray())).ToArray();
    }
}
