using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartPointStylesTests {
    public static IEnumerable<object[]> StyledKinds() {
        foreach (OfficeChartKind kind in new[] { OfficeChartKind.Pie, OfficeChartKind.Doughnut,
            OfficeChartKind.ColumnClustered, OfficeChartKind.Scatter, OfficeChartKind.Bubble })
            foreach (OfficeChartHatchPattern hatch in Enum.GetValues(typeof(OfficeChartHatchPattern)))
                yield return new object[] { kind, hatch };
    }

    [Theory]
    [MemberData(nameof(StyledKinds))]
    public void PointStyles_SurviveNativeAuthoringReopenUpdateAndReset(OfficeChartKind kind, OfficeChartHatchPattern hatch) {
        var outline = new OfficeChartPointStyle(noFill: true, outlineColor: OfficeColor.Black, outlineWidth: 2);
        var pattern = new OfficeChartPointStyle(OfficeColor.White, hatch: hatch,
            hatchColor: OfficeColor.Parse("#7300A3"), outlineColor: OfficeColor.Black, outlineWidth: 1);
        using PowerPointPresentation authored = PowerPointPresentation.Create();
        authored.AddSlide().AddChartPoints(kind, Data(kind, new[] { null, outline, pattern }), 20, 20, 600, 320);
        using var bytes = new MemoryStream(authored.ToBytes());
        using PowerPointPresentation reopened = PowerPointPresentation.Load(bytes);
        PowerPointChart chart = reopened.Slides.Single().Charts.Single();
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        AssertStyles(snapshot, hatch);
        Assert.Empty(reopened.ValidateDocument());
        using (PowerPointPresentation htmlReopened = OfficeIMO.Html.HtmlConversionDocument.Parse(reopened.ToHtml()).ToPowerPointPresentation()) {
            Assert.True(htmlReopened.Slides.Single().Charts.Single().TryGetOfficeSnapshot(out OfficeChartSnapshot htmlSnapshot));
            AssertStyles(htmlSnapshot, hatch);
        }
        chart.UpdateData(Data(kind, new[] { null, outline, pattern }));
        Assert.True(chart.TryGetOfficeSnapshot(out snapshot));
        AssertStyles(snapshot, hatch);
        chart.SetDataPointStyle(0, 1, null).SetDataPointStyle(0, 2, null);
        Assert.True(chart.TryGetOfficeSnapshot(out snapshot));
        Assert.Null(snapshot.Data.Series[0].PointStyles);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    [InlineData(OfficeChartKind.ColumnClustered)]
    public void PointStyles_RenderHatchesInPngAndHtml(OfficeChartKind kind) {
        var hatch = new OfficeChartPointStyle(OfficeColor.White, hatch: OfficeChartHatchPattern.DiagonalCross,
            hatchColor: OfficeColor.Parse("#7300A3"), outlineColor: OfficeColor.Black, outlineWidth: 2);
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.SlideSize.SetSizePoints(640, 360);
        presentation.AddSlide().AddChartPoints(kind, Data(kind, new[] { null, null, hatch }), 20, 20, 600, 320);
        Assert.True(OfficePngReader.TryDecode(presentation.Slides.Single().ToPng(), out OfficeRasterImage? raster));
        int hatchPixels = 0;
        for (int y = 0; y < raster!.Height; y++)
            for (int x = 0; x < raster.Width; x++)
                if (raster.GetPixel(x, y).Equals(hatch.HatchColor!.Value)) hatchPixels++;
        Assert.True(hatchPixels > 50, "Hatch strokes should be visible; actual pixels " + hatchPixels);
        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions { ExportProfile = PowerPointHtmlExportProfile.VisualReview });
        Assert.Contains("clipPath", html, StringComparison.Ordinal);
        Assert.Contains("#7300A3", html, StringComparison.OrdinalIgnoreCase);
    }

    private static OfficeChartData Data(OfficeChartKind kind, OfficeChartPointStyle?[] styles) {
        OfficeChartSeries series = kind == OfficeChartKind.Bubble
            ? OfficeChartSeries.CreateBubble("Status", new double[] { 1, 2, 3 }, new double[] { 3, 2, 1 }, new double[] { 3, 2, 1 })
            : new OfficeChartSeries("Status", new double[] { 3, 2, 1 },
                kind == OfficeChartKind.Scatter ? new double[] { 1, 2, 3 } : null);
        return new OfficeChartData(new[] { "Pass", "Unknown", "Fail" }, new[] { series.WithPointStyles(styles) });
    }

    private static void AssertStyles(OfficeChartSnapshot snapshot, OfficeChartHatchPattern hatch) {
        IReadOnlyList<OfficeChartPointStyle?> styles = snapshot.Data.Series[0].PointStyles!;
        Assert.Null(styles[0]);
        Assert.True(styles[1]!.NoFill);
        Assert.Equal(OfficeColor.Black, styles[1]!.OutlineColor);
        Assert.Equal(2, styles[1]!.OutlineWidth);
        Assert.Equal(hatch, styles[2]!.Hatch);
        Assert.Equal(OfficeColor.Parse("#7300A3"), styles[2]!.HatchColor);
    }
}
