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
        Assert.DoesNotContain(report.Warnings, w => w.Code == "HeaderFooterPageBoundsClipped");
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
        Assert.Contains(report.Warnings, w => w.Code == "HeaderFooterPageBoundsClipped");
    }

    [Fact]
    public void DiagnosticDashSubdivisionLimitDoesNotPreventNativePdfWriting() {
        var shape = OfficeShape.Path(10000, 1, OfficePathCommand.MoveTo(0, .5), OfficePathCommand.LineTo(10000, .5));
        shape.StrokeWidth = .01; shape.StrokeColor = OfficeColor.Red;
        shape.SetStrokeDashArray(new[] { .01, .01 });
        var report = new PdfConversionReport();
        byte[] bytes = Running(shape, report);
        Assert.NotEmpty(bytes);
        Assert.Contains(report.Warnings, w => w.Code == "HeaderFooterPaintBoundsUnmeasured");
        Assert.DoesNotContain(report.Warnings, w => w.Code == "HeaderFooterPageBoundsClipped");
    }

    [Fact]
    public void LinkedLineAnnotationIncludesItsMarkerOverhang() {
        var shape = OfficeShape.Line(0, 0, 60, 0);
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 2;
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 30, 20);
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false,
            PageWidth = 200, PageHeight = 200, MarginLeft = 30, MarginRight = 30,
            MarginTop = 50, MarginBottom = 30 }).Compose(c => c.Page(p => p.Content(b =>
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
        Assert.DoesNotContain(report.Warnings, w => w.Code == "HeaderFooterPageBoundsClipped");
        Assert.False(report.HasLoss);
    }

    private static byte[] Running(OfficeShape shape, PdfConversionReport report) =>
        PdfDocument.Create(new PdfOptions { PageWidth = 100, PageHeight = 100,
            MarginLeft = 10, MarginRight = 10, MarginTop = 20, MarginBottom = 20, HeaderOffsetY = 0 }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")).Compose(c => c.Page(p => {
                p.Header(b => b.Shape(shape)); p.Content(b => b.Spacer(10));
            })).ToBytes();
}
