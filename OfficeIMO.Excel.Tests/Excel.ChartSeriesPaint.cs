using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;

namespace OfficeIMO.Tests;
public sealed class ExcelChartSeriesPaintTests
{
    [Fact]
    public void NativeSeriesFillAndOutlineRemainIndependent() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "SeriesPaint", "series-paint.xlsx");
        for (int number = 0; number < 16; number++) {
            using var document = ExcelDocument.Load(path);
            var chart = document.Sheets[number].Charts.Single();
            Assert.True(chart.TryGetSnapshot(out var native));
            Assert.True(OfficeIMO.Excel.ExcelRangeImageRenderer.TryCreateOfficeChartSnapshot(native,
                640, 360, null, "Chart" + (number + 1), out var snapshot));
            Assert.NotNull(snapshot);
            CheckPaint(snapshot!, number);
        }
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
    [InlineData(true, true)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(false, false)]
    public void RibbonProjectionKeepsItsEnabledPaintIndependent(bool fill, bool outline)
    {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Paint");
        sheet.CellValue(1, 1, "Category");
        sheet.CellValue(1, 2, "Series");
        sheet.CellValue(2, 1, "A");
        sheet.CellValue(2, 2, 2);
        sheet.CellValue(3, 1, "B");
        sheet.CellValue(3, 2, 4);
        var chart = sheet.AddChartFromRange("A1:B3", row: 1, column: 4, width: OfficeLength.Pixels(640), height: OfficeLength.Pixels(360), title: null, includeCachedData: true, type: OfficeChartType.Line3D, hasHeaders: true);
        chart.SetSeriesFillColor(0, OfficeColor.Parse("0000FF")).SetSeriesLineColor(0, OfficeColor.Parse("FF0000"), OfficeLength.Points(3));
        if (!fill)
            chart.SetSeriesNoFill(0, noLine: !outline);
        else if (!outline)
        {
            var line = sheet.WorksheetPart.DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Outline>().First();
            line.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.SolidFill>();
            line.AddChild(new DocumentFormat.OpenXml.Drawing.NoFill(), true);
        }

        Assert.True(chart.TryGetSnapshot(out var native));
        Assert.True(ExcelRangeImageRenderer.TryCreateOfficeChartSnapshot(native, 640, 360, null, "Paint", out var snapshot));
        var series = Assert.Single(snapshot!.Data.Series);
        Assert.False(series.ShowMarkers);
        Assert.Equal(fill || outline, series.SeriesStyle!.ShowOutline);
        var drawing = OfficeChartDrawingRenderer.Render(snapshot);
        Assert.Equal(fill || outline, drawing.Shapes.Any(shape => shape.Shape.StrokeWidth == 3 && shape.Shape.StrokeColor == (fill ? OfficeColor.Blue : OfficeColor.Red) && shape.Shape.Kind is OfficeShapeKind.Line or OfficeShapeKind.Path));
    }
}