using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelChartSeriesPaintTests {
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
}
