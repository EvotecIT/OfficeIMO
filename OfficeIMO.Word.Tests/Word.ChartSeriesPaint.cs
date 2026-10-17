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
        var series = new OfficeChartSeries("Paint", new[] { 2d, 4d }).WithSeriesStyle(
            new OfficeChartPointStyle(noFill: true, outlineColor: OfficeColor.Red, outlineWidth: 3, showOutline: true))
            .WithPointStyles(new OfficeChartPointStyle?[] { new(fillColor: OfficeColor.Green, showOutline: false), null });
        using var document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { series });
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, data);
        chart.SetData(OfficeChartKind.ColumnClustered, data);
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
    [InlineData(OfficeChartKind.ColumnClustered)]
    [InlineData(OfficeChartKind.Area)]
    [InlineData(OfficeChartKind.Bubble)]
    public void NativeRoundTripKeepsSeriesHatchAndOutlineJoin(OfficeChartKind kind) {
        var series = kind == OfficeChartKind.Bubble
            ? OfficeChartSeries.CreateBubble("Paint", new[] { 1d, 2d }, new[] { 2d, 4d }, new[] { 3d, 5d })
            : new OfficeChartSeries("Paint", new[] { 2d, 4d });
        series = series.WithSeriesStyle(new OfficeChartPointStyle(fillColor: OfficeColor.White,
            hatch: OfficeChartHatchPattern.Cross, hatchColor: OfficeColor.Green,
            outlineColor: OfficeColor.Red, outlineWidth: 3, outlineJoin: OfficeStrokeLineJoin.Round));
        using var document = WordDocument.Create();
        document.AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] { series }));
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
}
