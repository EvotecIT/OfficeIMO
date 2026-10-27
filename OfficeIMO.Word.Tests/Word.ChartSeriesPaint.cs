using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordChartSeriesPaintTests {
    [Fact]
    public void NativeSeriesFillAndOutlineRemainIndependent() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "SeriesPaint", "series-paint.docx");
        for (int number = 0; number < 16; number++) {
            using var document = WordDocument.Load(path);
            var chart = document.Charts[number];
            Assert.True(chart.TryGetOfficeSnapshot(out var snapshot), "Native chart " + (number + 1));
            Assert.True(WordPdfConverterExtensions.TryCreateNativeWordChartSnapshot(chart, out var pdfSnapshot, out var reason), reason);
            CheckPaint(pdfSnapshot!, number);
            CheckPaint(snapshot!, number);
        }
    }

    [Fact]
    public void SharedNativeWriterKeepsSeriesAndPointPaintThroughDataUpdates() {
        var series = new OfficeSeries("Paint", new[] { 2d, 4d }).WithSeriesStyle(
            new OfficeChartPointStyle(noFill: true, outlineColor: OfficeColor.Red, outlineWidth: 3, showOutline: true))
            .WithPointStyles(new OfficeChartPointStyle?[] { new(fillColor: OfficeColor.Green, showOutline: false), null });
        using var document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { series });
        var chart = document.AddChart(OfficeChartType.ColumnClustered,data);
        chart.SetData(OfficeChartType.ColumnClustered, data);
        Assert.Empty(document.ValidateDocument());
        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using var reopened = WordDocument.Load(stream);
        Assert.True(reopened.Charts[0].TryGetOfficeSnapshot(out var snapshot));
        var saved = Assert.Single(snapshot.Data.Series);
        Assert.True(saved.SeriesStyle!.NoFill);
        Assert.Equal(OfficeColor.Red, saved.SeriesStyle.OutlineColor);
        Assert.Equal(3, saved.SeriesStyle.OutlineWidth);
        Assert.Equal(OfficeColor.Green, saved.PointStyles![0]!.FillColor);
        Assert.False(saved.PointStyles[0]!.ShowOutline);
    }

    [Theory]
    [InlineData(OfficeChartType.ColumnClustered)]
    [InlineData(OfficeChartType.Area)]
    [InlineData(OfficeChartType.Bubble)]
    public void NativeRoundTripKeepsSeriesHatchAndOutlineJoin(OfficeChartType kind) {
        var series = kind == OfficeChartType.Bubble
            ? OfficeSeries.CreateBubble("Paint", new[] { 1d, 2d }, new[] { 2d, 4d }, new[] { 3d, 5d })
            : new OfficeSeries("Paint", new[] { 2d, 4d });
        series = series.WithSeriesStyle(new OfficeChartPointStyle(fillColor: OfficeColor.White,
            hatch: OfficeChartHatchPattern.Cross, hatchColor: OfficeColor.Green,
            outlineColor: OfficeColor.Red, outlineWidth: 3, outlineJoin: OfficeStrokeLineJoin.Round));
        using var document = WordDocument.Create();
        document.AddChart(kind,new OfficeChartData(new[] { "A", "B" }, new[] { series }));
        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using var reopened = WordDocument.Load(stream);
        Assert.True(reopened.Charts[0].TryGetOfficeSnapshot(out var snapshot));
        var paint = Assert.Single(snapshot.Data.Series).SeriesStyle!;
        Assert.Equal(OfficeChartHatchPattern.Cross, paint.Hatch);
        Assert.Equal(OfficeColor.Green, paint.HatchColor);
        Assert.Equal(OfficeStrokeLineJoin.Round, paint.OutlineJoin);
        Assert.Empty(reopened.ValidateDocument());
    }

    private static void CheckPaint(OfficeChartSnapshot snapshot, int number) {
        bool fill = number % 4 < 2, outline = number % 2 == 0, line = number >= 12;
        var paint = Assert.Single(snapshot.Data.Series).SeriesStyle;
        Assert.NotNull(paint);
        Assert.Equal(!fill, paint.NoFill);
        Assert.Equal(outline, paint.ShowOutline);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
        Assert.Equal(fill && !line, drawing.Shapes.Any(shape => shape.Shape.FillColor == OfficeColor.Parse("#156082")));
        Assert.Equal(outline, drawing.Shapes.Any(shape => shape.Shape.StrokeColor == OfficeColor.Parse("#E97132") && shape.Shape.StrokeWidth == 3));
    }

    [Theory]
    [InlineData(OfficeChartType.ColumnClustered)]
    [InlineData(OfficeChartType.Area)]
    [InlineData(OfficeChartType.Bubble)]
    [InlineData(OfficeChartType.Line)]
    public void NativeColourSetterReplacesNoFillWithoutOverwritingIndependentPaint(OfficeChartType kind) {
        var series = kind == OfficeChartType.Bubble
            ? OfficeSeries.CreateBubble("Paint", new[] { 1d, 2d }, new[] { 2d, 4d }, new[] { 3d, 5d })
            : new OfficeSeries("Paint", new[] { 2d, 4d }, null, null, null, showMarkers: false);
        series = series.WithSeriesStyle(new OfficeChartPointStyle(noFill: true,
            outlineColor: OfficeColor.Red, outlineWidth: 3, showOutline: true));
        using var document = WordDocument.Create();
        var chart = document.AddChart(kind,new OfficeChartData(new[] { "A", "B" }, new[] { series }));
        chart.SetSeriesColor(0, OfficeColor.Green);
        Assert.Empty(document.ValidateDocument());
        using var stream = new MemoryStream(); document.Save(stream); stream.Position = 0;
        using var reopened = WordDocument.Load(stream);
        Assert.True(reopened.Charts[0].TryGetOfficeSnapshot(out var snapshot));
        var paint = Assert.Single(snapshot.Data.Series).SeriesStyle!;
        Assert.Equal(kind == OfficeChartType.Line, paint.NoFill);
        Assert.Equal(kind == OfficeChartType.Line ? OfficeColor.Green : OfficeColor.Red, paint.OutlineColor);
        Assert.Equal(3, paint.OutlineWidth);
        if (kind != OfficeChartType.Line) Assert.Equal(OfficeColor.Green, paint.FillColor);
    }

    [Fact]
    public void NativeColourSetterReplacesGradientOutlineAndKeepsItsDashAndWidth() {
        var series = new OfficeSeries("Paint", new[] { 2d, 4d }, null, OfficeColor.Blue, null,
            showMarkers: false, strokeWidth: 3, strokeDashStyle: OfficeStrokeDashStyle.Dash);
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.Line,new OfficeChartData(new[] { "A", "B" }, new[] { series }));
        var outline = chart.ChartPart!.ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Outline>().First();
        outline.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.SolidFill>();
        outline.AddChild(new DocumentFormat.OpenXml.Drawing.GradientFill(
            new DocumentFormat.OpenXml.Drawing.GradientStopList(
                new DocumentFormat.OpenXml.Drawing.GradientStop(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "0000FF" }) { Position = 0 },
                new DocumentFormat.OpenXml.Drawing.GradientStop(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FF0000" }) { Position = 100000 })), true);
        Assert.Empty(document.ValidateDocument());
        chart.SetSeriesColor(0, OfficeColor.Green);
        Assert.Empty(document.ValidateDocument());
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        var saved = Assert.Single(snapshot.Data.Series);
        Assert.Equal(OfficeColor.Green, saved.SeriesStyle!.OutlineColor);
        Assert.Equal(3, saved.SeriesStyle.OutlineWidth);
        Assert.Equal(OfficeStrokeDashStyle.Dash, saved.StrokeDashStyle);
    }

    [Theory]
    [InlineData(true, true)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(false, false)]
    public void BothWordSnapshotRoutesProjectTheSameRibbonPaint(bool fill, bool outline) {
        using var source = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(
            Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "SeriesPaint", "advanced-ribbon.pptx"), false);
        var sourcePlot = source.PresentationPart!.SlideParts.SelectMany(slide => slide.ChartParts)
            .First(part => part.ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Line3DChart>().Any())
            .ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().Single();
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.Line,new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeSeries("Paint", new[] { 2d, 4d }) }));
        var part = chart.ChartPart!;
        var nativeChart = part.ChartSpace!.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>()!;
        nativeChart.ReplaceChild(sourcePlot.CloneNode(true), nativeChart.PlotArea!);
        nativeChart.Legend?.Remove();
        part.ChartSpace.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.Charts.ExternalData>();
        foreach (var embedded in part.GetPartsOfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>().ToArray()) part.DeletePart(embedded);
        foreach (var native in nativeChart.Descendants<DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries>()) {
            native.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.Charts.Marker>();
            native.AddChild(new DocumentFormat.OpenXml.Drawing.Charts.Marker(new DocumentFormat.OpenXml.Drawing.Charts.Symbol {
                Val = DocumentFormat.OpenXml.Drawing.Charts.MarkerStyleValues.Circle }), true);
            var properties = native.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>()!;
            properties.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.SolidFill>(); properties.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.NoFill>();
            properties.AddChild(fill ? (DocumentFormat.OpenXml.OpenXmlElement)new DocumentFormat.OpenXml.Drawing.SolidFill(
                new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "0000FF" }) : new DocumentFormat.OpenXml.Drawing.NoFill(), true);
            var line = properties.GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>()!;
            line.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.SolidFill>(); line.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.NoFill>();
            line.Width = 38100;
            line.AddChild(outline ? (DocumentFormat.OpenXml.OpenXmlElement)new DocumentFormat.OpenXml.Drawing.SolidFill(
                new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FF0000" }) : new DocumentFormat.OpenXml.Drawing.NoFill(), true);
        }
        Assert.True(chart.TryGetOfficeSnapshot(out var shared));
        Assert.True(WordPdfConverterExtensions.TryCreateNativeWordChartSnapshot(chart, out var pdf, out var reason), reason);
        foreach (var snapshot in new[] { shared, pdf! }) {
            Assert.All(snapshot.Data.Series, series => {
                Assert.False(series.ShowMarkers);
                Assert.Equal(fill || outline, series.SeriesStyle!.ShowOutline);
            });
            Assert.Equal(fill || outline, OfficeChartDrawingRenderer.Render(snapshot).Shapes.Any(shape =>
                shape.Shape.StrokeWidth == 3 && shape.Shape.StrokeColor == (fill ? OfficeColor.Blue : OfficeColor.Red) &&
                shape.Shape.Kind is OfficeShapeKind.Line or OfficeShapeKind.Path));
        }
    }
}
