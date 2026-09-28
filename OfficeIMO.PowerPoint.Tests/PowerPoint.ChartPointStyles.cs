using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartPointStylesTests {
    [Fact]
    public void PointStyles_ClearCompetingDirectAndMarkerOverrides() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChart(OfficeChartKind.Line,
            Data(OfficeChartKind.Line, new OfficeChartPointStyle?[] { null, null, null }));
        LineChartSeries series = presentation.Slides.Single().SlidePart.ChartParts.Single()
            .ChartSpace!.Descendants<LineChartSeries>().Single();
        var point = new DataPoint(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0U },
            new ChartShapeProperties(new SolidFill(new RgbColorModelHex { Val = "FF0000" })),
            new Marker(new ChartShapeProperties(new SolidFill(new RgbColorModelHex { Val = "0000FF" }))));
        series.AddChild(point, true);

        chart.SetDataPointStyle(0, 0, new OfficeChartPointStyle(OfficeColor.FromRgb(0, 128, 0)));
        Assert.Null(point.ChartShapeProperties?.GetFirstChild<SolidFill>());
        Assert.Equal("008000", point.Marker!.ChartShapeProperties!.GetFirstChild<SolidFill>()!
            .GetFirstChild<RgbColorModelHex>()!.Val!.Value);
        chart.SetDataPointStyle(0, 0, null);
        Assert.Null(point.ChartShapeProperties?.GetFirstChild<SolidFill>());
        Assert.Null(point.Marker?.ChartShapeProperties?.GetFirstChild<SolidFill>());
        Assert.Empty(presentation.ValidateDocument());
    }

    [Fact]
    public void PointStyles_WidthOnlyColumnOutlineMatchesDirectAndReopenedRendering() {
        var data = Data(OfficeChartKind.ColumnClustered, new OfficeChartPointStyle?[] {
            new(outlineWidth: 2), null, null });
        using PowerPointPresentation authored = PowerPointPresentation.Create();
        authored.AddSlide().AddChartPoints(OfficeChartKind.ColumnClustered, data, 20, 20, 600, 320);
        using var bytes = new MemoryStream(authored.ToBytes());
        using PowerPointPresentation reopened = PowerPointPresentation.Load(bytes);
        Assert.True(reopened.Slides.Single().Charts.Single().TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        foreach (OfficeChartData source in new[] { data, snapshot.Data }) {
            OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Results", null,
                OfficeChartKind.ColumnClustered, source, 600, 320, layout: new OfficeChartLayout(showLegend: false)));
            Assert.Contains(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                shape.Shape.StrokeColor == OfficeColor.Black && shape.Shape.StrokeWidth == 2 && shape.Shape.Height > 50);
        }
        Assert.Empty(reopened.ValidateDocument());
    }

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
    [InlineData(OfficeChartKind.Bubble)]
    public void PointStyles_PreserveExplicitOutlineVisibilityWithoutColourOrWidth(OfficeChartKind kind) {
        using PowerPointPresentation authored = PowerPointPresentation.Create();
        authored.AddSlide().AddChartPoints(kind, Data(kind, new OfficeChartPointStyle?[] {
            new(showOutline: true), new(showOutline: false), null }), 20, 20, 600, 320);
        using var bytes = new MemoryStream(authored.ToBytes());
        using PowerPointPresentation reopened = PowerPointPresentation.Load(bytes);
        Assert.True(reopened.Slides.Single().Charts.Single().TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.True(snapshot.Data.Series[0].PointStyles![0]!.ShowOutline);
        Assert.False(snapshot.Data.Series[0].PointStyles![1]!.ShowOutline);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    [InlineData(OfficeChartKind.ColumnClustered)]
    public void PointStyles_RenderHatchesInPngHtmlAndPdf(OfficeChartKind kind) {
        var hatch = new OfficeChartPointStyle(hatch: OfficeChartHatchPattern.DiagonalCross,
            hatchColor: OfficeColor.Parse("#7300A3"), outlineColor: OfficeColor.Black, outlineWidth: 2);
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.SlideSize.SetSizePoints(640, 360);
        presentation.AddSlide().AddChartPoints(kind, Data(kind, new[] { null, null, hatch }), 20, 20, 600, 320);
        Assert.True(OfficePngReader.TryDecode(presentation.Slides.Single().ToPng(), out OfficeRasterImage? raster));
        int hatchPixels = 0;
        for (int y = 0; y < raster!.Height; y++)
            for (int x = 0; x < raster.Width; x++)
                if (IsPurpleStroke(raster.GetPixel(x, y))) hatchPixels++;
        Assert.True(hatchPixels > 50, "Hatch strokes should be visible; actual pixels " + hatchPixels);
        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions { ExportProfile = PowerPointHtmlExportProfile.VisualReview });
        Assert.Contains("clipPath", html, StringComparison.Ordinal);
        Assert.Contains("#7300A3", html, StringComparison.OrdinalIgnoreCase);
        byte[] pdf = presentation.ToPdfBytes();
        var page = Assert.Single(OfficeIMO.Pdf.PdfDocument.Load(pdf).Render.Pages(options:
            new OfficeIMO.Pdf.PdfPageRenderOptions { Dpi = 72, MaxPages = 1, ContinueOnError = false }));
        Assert.True(OfficePngReader.TryDecode(page.Bytes!, out raster));
        hatchPixels = 0;
        for (int y = 0; y < raster!.Height; y++)
            for (int x = 0; x < raster.Width; x++)
                if (IsPurpleStroke(raster.GetPixel(x, y))) hatchPixels++;
        Assert.True(hatchPixels > 50, "Hatch strokes should be visible in PDF; actual pixels " + hatchPixels);
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
    // Thin hatch strokes blend with their background during antialiasing.
    private static bool IsPurpleStroke(OfficeColor pixel) =>
        pixel.R < 200 && pixel.G < 150 && pixel.B > pixel.G + 40 && pixel.R > pixel.G + 20;
}
