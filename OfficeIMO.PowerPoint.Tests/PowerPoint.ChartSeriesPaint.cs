using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartSeriesPaintTests {
    [Fact]
    public void NativeSeriesFillAndOutlineRemainIndependent() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "Charts", "SeriesPaint", "series-paint.pptx");
        for (int number = 0; number < 16; number++) {
            using var document = PowerPointPresentation.Load(path);
            var chart = document.Slides[number].Charts.Single();
            Assert.True(chart.TryGetSnapshot(out var native, out var reason), reason);
            var snapshot = PowerPointChartSnapshotMapper.ToOfficeSnapshot(native, 640, 360);
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
