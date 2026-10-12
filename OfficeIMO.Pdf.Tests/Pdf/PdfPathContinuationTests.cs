using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;
public sealed class PdfPathContinuationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void QuadraticAfterCloseUsesTheClosedOriginInBothNativeWriters(bool local) {
        var shape = OfficeShape.Path(20, 20, OfficePathCommand.MoveTo(2, 3), OfficePathCommand.LineTo(10, 3),
            OfficePathCommand.LineTo(2, 10), OfficePathCommand.Close(), OfficePathCommand.QuadraticBezierTo(12, 3, 12, 12));
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 1;
        if (local) shape.Transform = OfficeTransform.Identity;
        byte[] bytes = PdfDocument.Create(new PdfOptions { PageWidth = 100, PageHeight = 100,
            MarginLeft = 0, MarginRight = 0, MarginTop = 0, MarginBottom = 0, CompressContentStreams = false })
            .Canvas(canvas => canvas.Drawing(new OfficeDrawing(100, 100).AddShape(shape, 0, 0), 0, 0, 100, 100)).ToBytes();
        Match curve = Regex.Match(Encoding.ASCII.GetString(bytes), @"(-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) c\b");
        Assert.True(curve.Success);
        Match origin = Regex.Match(Encoding.ASCII.GetString(bytes), @"(-?[\d.]+) (-?[\d.]+) m");
        Assert.True(origin.Success);
        Assert.Equal(double.Parse(origin.Groups[2].Value, CultureInfo.InvariantCulture),
            double.Parse(curve.Groups[2].Value, CultureInfo.InvariantCulture), 3);
    }

    [Fact]
    public void CanvasAnnotationsSurviveTheImplicitContourAfterClose() {
        var clip = OfficeClipPath.Path(100, 100, new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(10, 0),
            OfficePathCommand.LineTo(0, 10), OfficePathCommand.Close(), OfficePathCommand.LineTo(100, 0),
            OfficePathCommand.LineTo(100, 100), OfficePathCommand.LineTo(0, 100) });
        byte[] bytes = PdfDocument.Create(new PdfOptions { PageWidth = 200, PageHeight = 200 })
            .Canvas(canvas => canvas.Clip(0, 0, clip, inner => inner.Text(new[] { PdfTextRun.Link("kept", "https://example.test/kept") }, 40, 40, 40, 15))).ToBytes();
        Assert.Single(PdfInspector.Inspect(bytes).LinkAnnotations);
    }
}

