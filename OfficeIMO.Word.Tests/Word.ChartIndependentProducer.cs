using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartIndependentProducerTests {
    [Fact]
    public void LibreOfficeStatusPieRetainsPointAppearanceInNativeAndStaticPaths() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "LibreOffice", "status-pie.docx");
        using WordDocument document = WordDocument.Load(path);
        WordChart chart = Assert.Single(document.Charts);
        string nativeXml = chart.ChartPart!.ChartSpace!.OuterXml;

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartKind.Pie, snapshot.ChartKind);
        Assert.Equal(new[] { "Pass", "Fail", "Unknown" }, snapshot.Data.Categories);
        var styles = Assert.Single(snapshot.Data.Series).PointStyles!;
        Assert.Equal(OfficeColor.Parse("#228844"), styles[0]!.FillColor);
        Assert.Equal(OfficeChartHatchPattern.WideForwardDiagonal, styles[1]!.Hatch);
        Assert.Equal(OfficeColor.Parse("#D97706"), styles[1]!.HatchColor);
        Assert.True(styles[2]!.NoFill);
        Assert.Equal(OfficeColor.Parse("#445566"), styles[2]!.OutlineColor);
        Assert.Equal(OfficeStrokeLineJoin.Round, styles[2]!.OutlineJoin);
        Assert.NotNull(snapshot.Style.LegendBackgroundColor);
        Assert.NotNull(snapshot.Style.LegendBorderColor);

        var rendered = document.ExportImage(OfficeImageExportFormat.Png,
            new WordImageExportOptions { Policy = new OfficeImageExportPolicy { RequireNoOmissions = true, RequireNoFailures = true } });
        Assert.NotEmpty(rendered.Bytes);
        string referencePath = Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "LibreOffice", "status-pie-docx-reference.png");
        Assert.True(OfficePngReader.TryDecode(File.ReadAllBytes(referencePath), out OfficeRasterImage? reference));
        Assert.True(OfficePngReader.TryDecode(rendered.Bytes, out OfficeRasterImage? actual));
        // Page placement differs between LibreOffice and OfficeIMO, so compare
        // the distinctive colour and hatch evidence rather than raw page pixels.
        Assert.True(CountColour(reference!, "#228844") > 1000);
        Assert.True(CountColour(actual!, "#228844") > 1000);
        Assert.True(CountWarmHatch(reference!) > 50);
        Assert.True(CountWarmHatch(actual!) > 50);
        Assert.Equal(7, GetWarmHatchSpacing(reference!));
        Assert.InRange(GetWarmHatchSpacing(actual!), 6, 8);
        Assert.True(CountColour(reference!, "#445566") > 50);
        Assert.True(CountColour(actual!, "#445566") > 50);
        Assert.Equal(nativeXml, chart.ChartPart.ChartSpace.OuterXml);

        using var package = document.ToStream();
        using WordDocument reopened = WordDocument.Load(package);
        Assert.True(Assert.Single(reopened.Charts).TryGetOfficeSnapshot(out var second));
        Assert.Equal(OfficeChartHatchPattern.WideForwardDiagonal, second.Data.Series[0].PointStyles![1]!.Hatch);
    }

    [Fact]
    public void LibreOfficeAreaPreservesNativePointRecordsAndRendersItsOpaqueSeriesFill() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "LibreOffice", "area-point.docx");
        using WordDocument document = WordDocument.Load(path);
        WordChart chart = Assert.Single(document.Charts);
        string nativeXml = chart.ChartPart!.ChartSpace!.OuterXml;

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartKind.Area, snapshot.ChartKind);
        Assert.Equal(new[] { "A", "B", "C", "D" }, snapshot.Data.Categories);
        OfficeChartSeries series = Assert.Single(snapshot.Data.Series);
        Assert.Equal(OfficeColor.Parse("#004586"), series.Color);
        Assert.False(series.ConnectLine);
        Assert.Equal(4, series.PointStyles!.Count);
        Assert.Equal(OfficeColor.Parse("#228844"), series.PointStyles[0]!.FillColor);

        OfficeChartRenderingResult drawing = OfficeChartDrawingRenderer.RenderWithQuality(snapshot);
        Assert.Contains(drawing.QualityReport.Issues, issue => issue.Kind == OfficeDrawingQualityIssueKind.UnsupportedAppearance);
        var exported = document.ExportImage(OfficeImageExportFormat.Png,
            new WordImageExportOptions { Policy = new OfficeImageExportPolicy { RequireNoOmissions = true, RequireNoFailures = true } });
        Assert.Contains(exported.Diagnostics, item => item.Code == "ChartPointStylesUnsupported");
        Assert.True(OfficePngReader.TryDecode(exported.Bytes, out OfficeRasterImage? raster));
        OfficeColor nativeFill = OfficeColor.Parse("#004586");
        int bluePixels = 0;
        for (int y = 0; y < raster!.Height; y++)
            for (int x = 0; x < raster.Width; x++)
                if (raster.GetPixel(x, y) == nativeFill) bluePixels++;
        Assert.True(bluePixels > 1000, "Expected the opaque native series fill; actual pixels " + bluePixels);
        Assert.Equal(nativeXml, chart.ChartPart.ChartSpace.OuterXml);

        chart.SetData(snapshot.ChartKind, snapshot.Data);
        Assert.NotNull(chart.ChartPart.ChartSpace.Descendants<C.AreaChartSeries>().Single()
            .GetFirstChild<C.ChartShapeProperties>()?.GetFirstChild<A.Outline>()?.GetFirstChild<A.NoFill>());

        using var package = document.ToStream();
        using WordDocument reopened = WordDocument.Load(package);
        Assert.True(Assert.Single(reopened.Charts).TryGetOfficeSnapshot(out var second));
        Assert.Equal(4, second.Data.Series[0].PointStyles!.Count);
    }

    private static int CountColour(OfficeRasterImage image, string hex) {
        OfficeColor target = OfficeColor.Parse(hex);
        int count = 0;
        for (int y = 0; y < image.Height; y++)
            for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (Math.Abs(pixel.R - target.R) <= 30 && Math.Abs(pixel.G - target.G) <= 30 &&
                    Math.Abs(pixel.B - target.B) <= 30) count++;
            }
        return count;
    }

    private static int CountWarmHatch(OfficeRasterImage image) {
        int count = 0;
        for (int y = 0; y < image.Height; y++)
            for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (pixel.R > 190 && pixel.R > pixel.G + 20 && pixel.G > 80 &&
                    pixel.G > pixel.B + 20) count++;
            }
        return count;
    }

    private static int GetWarmHatchSpacing(OfficeRasterImage image) {
        var gapCounts = new int[21];
        for (int y = 175; y < Math.Min(270, image.Height); y += 5) {
            int lastStart = -1;
            bool inStroke = false;
            for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                bool warm = pixel.R > 190 && pixel.R > pixel.G + 20 && pixel.G > 80 &&
                    pixel.G > pixel.B + 20;
                if (warm && !inStroke) {
                    int gap = x - lastStart;
                    if (lastStart >= 0 && gap >= 4 && gap < gapCounts.Length) gapCounts[gap]++;
                    lastStart = x;
                }
                inStroke = warm;
            }
        }
        int mostFrequentGap = 0;
        for (int gap = 4; gap < gapCounts.Length; gap++)
            if (gapCounts[gap] > gapCounts[mostFrequentGap]) mostFrequentGap = gap;
        return mostFrequentGap;
    }
}
