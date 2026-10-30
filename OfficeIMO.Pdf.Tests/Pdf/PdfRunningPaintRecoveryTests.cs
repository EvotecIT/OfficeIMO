using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfRunningPaintRecoveryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DuplicateEndpointStopsUseTheNativePaintersPadding(bool end) {
        var shape = OfficeShape.Path(100, 1, OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(10, 0),
            OfficePathCommand.LineTo(10, 1), OfficePathCommand.LineTo(0, 1), OfficePathCommand.Close());
        shape.FillGradient = end
            ? new OfficeLinearGradient(1, 0, .5, 0, new[] { new OfficeGradientStop(0, OfficeColor.Red),
                new OfficeGradientStop(1, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Transparent) })
            : new OfficeLinearGradient(.5, 0, 1, 0, new[] { new OfficeGradientStop(0, OfficeColor.Transparent),
                new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Red) });
        shape.Transform = OfficeTransform.Translate(0, -30);
        var report = new PdfConversionReport();
        Assert.NotEmpty(Running(shape, report));
        Assert.Contains(report.Warnings, w => w.Code == "header-footer-page-bounds-clipped");
    }

    [Fact]
    public void RadialCoordinateScalingPreservesPaintAtThePhysicalEdge() {
        var shape = OfficeShape.Path(100, 100, OfficePathCommand.MoveTo(9, 49), OfficePathCommand.LineTo(9.5, 49),
            OfficePathCommand.LineTo(9.5, 50), OfficePathCommand.LineTo(9, 50), OfficePathCommand.Close());
        shape.FillRadialGradient = new OfficeRadialGradient(-.499, 0, 0, -.499, 0, 1, new[] {
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(.5, OfficeColor.Red),
            new OfficeGradientStop(.5, OfficeColor.Transparent), new OfficeGradientStop(1, OfficeColor.Transparent)
        }).TransformCoordinates(OfficeTransform.Scale(100, 100));
        shape.Transform = OfficeTransform.Translate(80.75, 0);
        var report = new PdfConversionReport();
        Assert.NotEmpty(Running(shape, report));
        Assert.Contains(report.Warnings, w => w.Code == "header-footer-page-bounds-clipped");
    }

    [Fact]
    public void TranslatedLinkedLineIncludesTheShaftBeyondItsStartMarker() {
        var shape = OfficeShape.Line(0, 0, 60, 0);
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 2;
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 30, 20);
        shape.Transform = OfficeTransform.Translate(100, 0);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(200), PageHeight = global::OfficeIMO.OfficeLength.Points(100),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10), MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20) })
            .Compose(c => c.Page(p => p.Content(b => b.Shape(shape, linkUri: "https://example.com/shape")))).ToBytes();
        Match match = Regex.Match(Encoding.ASCII.GetString(bytes), @"/Rect\s*\[([^\]]+)\]");
        Assert.True(match.Success);
        double[] rect = match.Groups[1].Value.Split(new[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
        Assert.True(rect[2] >= 170, "The translated shaft must remain clickable beyond the start marker.");
    }

    [Fact]
    public void TransparentPartOutsideThePageDoesNotHideVisibleGradientInkInsideIt() {
        var shape = OfficeShape.Path(100, 1, OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(60, 0),
            OfficePathCommand.LineTo(60, 1), OfficePathCommand.LineTo(0, 1), OfficePathCommand.Close());
        shape.FillGradient = new OfficeLinearGradient(0, 0, 1, 0, new[] {
            new OfficeGradientStop(0, OfficeColor.Transparent), new OfficeGradientStop(.5, OfficeColor.Transparent),
            new OfficeGradientStop(1, OfficeColor.Red) });
        shape.Transform = OfficeTransform.Translate(-30, 0);
        var report = new PdfConversionReport();
        Assert.NotEmpty(Running(shape, report));
        Assert.DoesNotContain(report.Warnings, w => w.Code == "header-footer-page-bounds-clipped");
    }

    [Fact]
    public void ScaledBezierPaintReportsInkBeyondThePage() {
        var shape = OfficeShape.Path(3, 1, OfficePathCommand.MoveTo(1, 0),
            OfficePathCommand.QuadraticBezierTo(2, .01, 0, .02), OfficePathCommand.Close());
        shape.FillColor = OfficeColor.Red;
        // The exact x extremum is 4/3 at t=1/3. The old eight-piece local
        // approximation misses the physical edge after a50x horizontal scale.
        shape.Transform = new OfficeTransform(50, 0, 0, 1, 100.1 - 10 - 50D * 4D / 3D, 0);
        var report = new PdfConversionReport();
        Assert.NotEmpty(Running(shape, report));
        Assert.Contains(report.Warnings, w => w.Code == "header-footer-page-bounds-clipped");
    }

    [Fact]
    public void DiagnosticDashSubdivisionLimitDoesNotPreventNativePdfWriting() {
        var shape = OfficeShape.Path(10000, 1, OfficePathCommand.MoveTo(0, .5), OfficePathCommand.LineTo(10000, .5));
        shape.StrokeWidth = .01; shape.StrokeColor = OfficeColor.Red;
        shape.SetStrokeDashArray(new[] { .01, .01 });
        var report = new PdfConversionReport();
        byte[] bytes = Running(shape, report);
        Assert.NotEmpty(bytes);
        Assert.Contains(report.Warnings, w => w.Code == "header-footer-paint-bounds-unmeasured");
        Assert.DoesNotContain(report.Warnings, w => w.Code == "header-footer-page-bounds-clipped");
    }

    [Fact]
    public void LinkedLineAnnotationIncludesItsMarkerOverhang() {
        var shape = OfficeShape.Line(0, 0, 60, 0);
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 2;
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 30, 20);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false,
            PageWidth = global::OfficeIMO.OfficeLength.Points(200), PageHeight = global::OfficeIMO.OfficeLength.Points(200), MarginLeft = global::OfficeIMO.OfficeLength.Points(30), MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(50), MarginBottom = global::OfficeIMO.OfficeLength.Points(30) }).Compose(c => c.Page(p => p.Content(b =>
                b.Shape(shape, linkUri: "https://example.com/shape")))).ToBytes();
        Match match = Regex.Match(Encoding.ASCII.GetString(bytes), @"/Rect\s*\[([^\]]+)\]");
        Assert.True(match.Success);
        double[] rect = match.Groups[1].Value.Split(new[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
        Assert.True(rect[3] - rect[1] >= 30, "The clickable rectangle must include the full marker height.");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FullyTransparentGradientRegionDoesNotReportPhysicalClipping(bool radial) {
        var shape = OfficeShape.Path(100, 1, OfficePathCommand.MoveTo(0, .5), OfficePathCommand.LineTo(10, .5),
            OfficePathCommand.LineTo(10, .51), OfficePathCommand.LineTo(0, .51), OfficePathCommand.Close());
        var stops = new[] {
            new OfficeGradientStop(0, OfficeColor.Transparent), new OfficeGradientStop(.5, OfficeColor.Transparent),
            new OfficeGradientStop(1, OfficeColor.Red) };
        if (radial) shape.FillRadialGradient = new OfficeRadialGradient(0, .5, 0, 0, .5, 1, stops);
        else shape.FillGradient = new OfficeLinearGradient(0, 0, 1, 0, stops);
        shape.Transform = OfficeTransform.Translate(0, -30);
        var report = new PdfConversionReport();
        Assert.NotEmpty(Running(shape, report));
        Assert.DoesNotContain(report.Warnings, w => w.Code == "header-footer-page-bounds-clipped");
        Assert.False(report.HasLoss);
    }

    private static byte[] Running(OfficeShape shape, PdfConversionReport report) =>
        PdfDocument.Create(new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(100), PageHeight = global::OfficeIMO.OfficeLength.Points(100),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10), MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20), HeaderOffsetY = global::OfficeIMO.OfficeLength.Points(0) }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")).Compose(c => c.Page(p => {
                p.Header(b => b.Shape(shape)); p.Content(b => b.Spacer(10));
            })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
}
