using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordChartDoughnutTests {
    [Fact]
    public void Doughnut_SurvivesSaveReopenAppendAndPointStyling() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        try {
            using (WordDocument document = WordDocument.Create(path)) {
                WordChart chart = document.AddChart("Status", false, 360, 180);
                chart.AddDoughnut("Pass", 8).AddDoughnut("Unknown", 2);
                chart.SetDataPointStyle(0, 0, new(fillColor: OfficeColor.Parse("#168A56")));
                chart.SetDataPointStyle(0, 1, new(noFill: true, outlineColor: OfficeColor.Black, outlineWidth: 2));
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(path)) {
                WordChart chart = Assert.Single(document.Charts);
                chart.AddDoughnut("Fail", 1);
                Assert.True(chart.TryGetSnapshot(out WordChartSnapshot snapshot));
                Assert.Equal(WordChartSnapshotKind.Doughnut, snapshot.ChartKind);
                Assert.Equal(new[] { "Pass", "Unknown", "Fail" }, snapshot.Data.Categories);
                Assert.Equal(new double[] { 8, 2, 1 }, snapshot.Data.Series[0].Values);
                Assert.True(snapshot.Data.Series[0].PointStyles![1]!.NoFill);
                Assert.Empty(document.ValidateDocument());
                var options = new WordToPdfOptions { IncludePageNumbers = false };
                byte[] pdf = document.ToPdfBytes(options);
                var page = Assert.Single(OfficeIMO.Pdf.PdfDocument.Load(pdf).Render.Pages(options:
                    new OfficeIMO.Pdf.PdfPageRenderOptions { Dpi = 72, MaxPages = 1, ContinueOnError = false }));
                Assert.True(OfficePngReader.TryDecode(page.Bytes!, out OfficeRasterImage? raster));
                int greenPixels = 0;
                for (int y = 0; y < raster!.Height; y++)
                    for (int x = 0; x < raster.Width; x++)
                        if (raster.GetPixel(x, y).Equals(OfficeColor.Parse("#168A56"))) greenPixels++;
                Assert.True(greenPixels > 100, "Expected the native doughnut slice in rendered PDF.");
                Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
                document.Save();
            }
            using WordDocument final = WordDocument.Load(path);
            Assert.True(final.Charts.Single().TryGetSnapshot(out WordChartSnapshot persisted));
            Assert.Equal(new double[] { 8, 2, 1 }, persisted.Data.Series[0].Values);
            Assert.Empty(final.ValidateDocument());
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Doughnut_RejectsInvalidValuesBeforeChangingTheChart() {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart();
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.AddDoughnut("Invalid", double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.AddDoughnut("Invalid", -1));
        chart.AddPie("Pass", 1);
        Assert.Throws<NotSupportedException>(() => chart.AddDoughnut("Fail", 1));
        Assert.True(chart.TryGetSnapshot(out WordChartSnapshot snapshot));
        Assert.Equal(new[] { "Pass" }, snapshot.Data.Categories);
    }
}
