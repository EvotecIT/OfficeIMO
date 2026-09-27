using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class WordChartPointStylesTests {
    [Fact]
    public void PointStyles_PreserveNativePieAppearanceAndPdfAfterReopen() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        try {
            using (WordDocument authored = WordDocument.Create(path)) {
                WordChart authoredChart = authored.AddChart("Status", false, 360, 180);
                authoredChart.AddPie("Pass", 3).AddPie("Unknown", 2).AddPie("Fail", 1);
                authoredChart.SetDataPointStyle(0, 0, new(showOutline: true));
                authoredChart.SetDataPointStyle(0, 1, new(noFill: true, outlineColor: OfficeColor.Black, outlineWidth: 2,
                    outlineJoin: OfficeStrokeLineJoin.Round));
                authoredChart.SetDataPointStyle(0, 2, new(OfficeColor.White, hatch: OfficeChartHatchPattern.WideForwardDiagonal,
                    hatchColor: OfficeColor.Parse("#7300A3"), outlineColor: OfficeColor.Black));
                authored.Save();
            }
            using WordDocument reopened = WordDocument.Load(path);
            WordChart chart = Assert.Single(reopened.Charts);
            Assert.True(chart.TryGetSnapshot(out WordChartSnapshot snapshot));
            Assert.True(snapshot.Data.Series[0].PointStyles![0]!.ShowOutline);
            Assert.True(snapshot.Data.Series[0].PointStyles![1]!.NoFill);
            Assert.Equal(OfficeStrokeLineJoin.Round, snapshot.Data.Series[0].PointStyles![1]!.OutlineJoin);
            Assert.Equal(OfficeChartHatchPattern.WideForwardDiagonal, snapshot.Data.Series[0].PointStyles![2]!.Hatch);
            Assert.Empty(reopened.ValidateDocument());
            var options = new WordToPdfOptions { IncludePageNumbers = false };
            byte[] pdf = reopened.ToPdfBytes(options);
            Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
            var page = Assert.Single(PdfCore.PdfDocument.Load(pdf).Render.Pages(options:
                new PdfCore.PdfPageRenderOptions { Dpi = 72, MaxPages = 1, ContinueOnError = false }));
            Assert.True(OfficePngReader.TryDecode(page.Bytes!, out OfficeRasterImage? raster));
            int pixels = 0;
            for (int y = 0; y < raster!.Height; y++)
                for (int x = 0; x < raster.Width; x++)
                    if (IsPurpleStroke(raster.GetPixel(x, y))) pixels++;
            Assert.True(pixels > 25, "Expected hatch strokes in the PDF; actual pixels " + pixels);
            chart.SetDataPointStyle(0, 0, null).SetDataPointStyle(0, 1, null).SetDataPointStyle(0, 2, null);
            Assert.True(chart.TryGetSnapshot(out snapshot));
            Assert.Null(snapshot.Data.Series[0].PointStyles);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }
    // Thin hatch strokes blend with their background during antialiasing.
    private static bool IsPurpleStroke(OfficeColor pixel) =>
        pixel.R < 200 && pixel.G < 150 && pixel.B > pixel.G + 40 && pixel.R > pixel.G + 20;
}
