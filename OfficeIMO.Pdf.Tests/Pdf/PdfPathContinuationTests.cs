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
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(100), PageHeight = global::OfficeIMO.OfficeLength.Points(100),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(0), MarginRight = global::OfficeIMO.OfficeLength.Points(0), MarginTop = global::OfficeIMO.OfficeLength.Points(0), MarginBottom = global::OfficeIMO.OfficeLength.Points(0), CompressContentStreams = false })
            .Canvas(canvas => canvas.Drawing(new OfficeDrawing(100, 100).AddShape(shape, 0, 0), global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(100), global::OfficeIMO.OfficeLength.Points(100))).ToBytes();
        Match curve = Regex.Match(Encoding.ASCII.GetString(bytes), @"(-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) c\b");
        Assert.True(curve.Success);
        Match origin = Regex.Match(Encoding.ASCII.GetString(bytes), @"(-?[\d.]+) (-?[\d.]+) m");
        Assert.True(origin.Success);
        Assert.Equal(double.Parse(origin.Groups[2].Value, CultureInfo.InvariantCulture),
            double.Parse(curve.Groups[2].Value, CultureInfo.InvariantCulture), 3);
    }

    [Fact]
    public void CanvasAnnotationsSurviveTheImplicitContourAfterClose() {
        var clip = OfficeClipPath.Path(new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(10, 0),
            OfficePathCommand.LineTo(0, 10), OfficePathCommand.Close(), OfficePathCommand.LineTo(100, 0),
            OfficePathCommand.LineTo(100, 100), OfficePathCommand.LineTo(0, 100) });
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(200), PageHeight = global::OfficeIMO.OfficeLength.Points(200) })
            .Canvas(canvas => canvas.Clip(global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(0), clip, inner => inner.Text(new[] { PdfTextRun.Link("kept", "https://example.test/kept") }, global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(15)))).ToBytes();
        Assert.Single(PdfInspector.Inspect(bytes).LinkAnnotations);
    }
}
