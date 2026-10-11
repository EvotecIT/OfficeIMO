using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfStrokeContractTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void ColorAlphaAndStrokeOpacityComposeForFilledAndOpenMarkersAtEveryPlacement(int placement) {
        var shape = OfficeShape.Line(0, 0, 60, 0);
        shape.StrokeColor = OfficeColor.FromRgba(204, 32, 64, 128);
        shape.StrokeOpacity = .5; shape.StrokeWidth = 2;
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 8, 10);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 8, 10);
        byte[] bytes = PaintAtPlacement(new PdfOptions { PageWidth = 100, PageHeight = 100,
            MarginLeft = 10, MarginRight = 10, MarginTop = 20, MarginBottom = 20 }, shape, placement);
        var paint = Shapes(PdfReadDocument.Open(bytes).Pages[0].ToDrawing()).Select(item => item.Shape)
            .Where(item => item.StrokeColor == OfficeColor.FromRgb(204, 32, 64) || item.FillColor == OfficeColor.FromRgb(204, 32, 64)).ToArray();
        Assert.Equal(3, paint.Length);
        Assert.All(paint, item => Assert.Equal(128D / 255D * .5,
            (item.StrokeColor.HasValue ? item.StrokeOpacity : item.FillOpacity).GetValueOrDefault(1), 3));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void ColorAlphaAndFillOpacityComposeAtEveryPlacement(int placement) {
        var shape = OfficeShape.Rectangle(40, 10);
        shape.FillColor = OfficeColor.FromRgba(204, 32, 64, 128); shape.FillOpacity = .25;
        byte[] bytes = PaintAtPlacement(new PdfOptions { PageWidth = 100, PageHeight = 100,
            MarginLeft = 10, MarginRight = 10, MarginTop = 20, MarginBottom = 20 }, shape, placement);
        var painted = Assert.Single(Shapes(PdfReadDocument.Open(bytes).Pages[0].ToDrawing()),
            item => item.Shape.FillColor == OfficeColor.FromRgb(204, 32, 64));
        Assert.Equal(128D / 255D * .25, painted.Shape.FillOpacity.GetValueOrDefault(1), 3);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void RunningShadowBoundsUseTheirEmittedMiterLimit(bool footer, bool transformed) {
        var report = new PdfConversionReport();
        var shape = OfficeShape.Path(40, 20, OfficePathCommand.MoveTo(14, 20), OfficePathCommand.LineTo(20, 0), OfficePathCommand.LineTo(26, 20));
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 2;
        shape.StrokeLineJoin = OfficeStrokeLineJoin.Miter; shape.StrokeMiterLimit = 2;
        shape.Shadow = new OfficeShadow(OfficeColor.Green, 1, 0, 0, 0);
        if (transformed) shape.Transform = OfficeTransform.Identity;
        byte[] bytes = Running(new PdfOptions { CompressContentStreams = false, PageWidth = 100, PageHeight = 100,
            MarginLeft = 3, MarginRight = 10, MarginTop = 40, MarginBottom = 40,
            HeaderOffsetY = 20, FooterOffsetY = 20 }.ReportDiagnosticsTo(report, "OfficeIMO.Tests"), shape, footer);
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(Encoding.ASCII.GetString(bytes), @"(?m)^2 M\r?$").Count);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "HeaderFooterPageBoundsClipped");
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    public void ExplicitMiterLimitSurvivesPdfPaintingOfSolidShapeShadows(int placement, bool transformedDash) {
        var shape = OfficeShape.Path(40, 40, OfficePathCommand.MoveTo(14, 40), OfficePathCommand.LineTo(20, 0), OfficePathCommand.LineTo(26, 40));
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 2;
        shape.StrokeLineJoin = OfficeStrokeLineJoin.Miter;
        shape.Shadow = new OfficeShadow(OfficeColor.Green, 1, 10, 0, 0);
        if (transformedDash) {
            shape.Transform = OfficeTransform.Identity;
            shape.SetStrokeDashArray(new[] { 4D, 2D });
        }
        var options = new PdfOptions { CompressContentStreams = false, PageWidth = 100, PageHeight = 100,
            MarginLeft = 10, MarginRight = 10, MarginTop = 20, MarginBottom = 20, HeaderOffsetY = 0, FooterOffsetY = 0 };
        byte[] bytes = PaintAtPlacement(options, shape, placement);
        var shadows = Shapes(PdfReadDocument.Open(bytes).Pages[0].ToDrawing()).Where(item => item.Shape.StrokeColor == OfficeColor.Green).ToArray();
        Assert.NotEmpty(shadows);
        Assert.All(shadows, item => Assert.Empty(item.Shape.StrokeDashArray));
        // The two independent stroke scopes must both declare their limit.
        // The PDF reader explicitly reports simplified M-operator rendering.
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(Encoding.ASCII.GetString(bytes), @"(?m)^4 M\r?$").Count);
    }

    [Theory]
    [InlineData(false, OfficeStrokeDashStyle.Dot)]
    [InlineData(true, OfficeStrokeDashStyle.Dot)]
    [InlineData(false, OfficeStrokeDashStyle.DashDot)]
    [InlineData(true, OfficeStrokeDashStyle.DashDot)]
    public void RunningPresetRoundCapsReportTheirPaintOutsideThePage(bool footer, OfficeStrokeDashStyle style) {
        var report = new PdfConversionReport();
        var shape = OfficeShape.Path(40, 10, OfficePathCommand.MoveTo(0, 5), OfficePathCommand.LineTo(40, 5));
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 2; shape.StrokeDashStyle = style;
        var options = new PdfOptions { PageWidth = 100, PageHeight = 100, MarginLeft = .5, MarginRight = 10,
            MarginTop = 20, MarginBottom = 20, HeaderOffsetY = 0, FooterOffsetY = 0 }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");
        byte[] bytes = Running(options, shape, footer);
        Assert.Contains(Shapes(PdfReadDocument.Open(bytes).Pages[0].ToDrawing()), item => item.Shape.StrokeLineCap == OfficeStrokeLineCap.Round);
        Assert.Contains(report.Warnings, warning => warning.Code == "HeaderFooterPageBoundsClipped");
        shape.StrokeLineCap = OfficeStrokeLineCap.Butt;
        var explicitReport = new PdfConversionReport();
        Running(new PdfOptions { PageWidth = 100, PageHeight = 100, MarginLeft = .5, MarginRight = 10,
            MarginTop = 20, MarginBottom = 20, HeaderOffsetY = 0, FooterOffsetY = 0 }.ReportDiagnosticsTo(explicitReport, "OfficeIMO.Tests"), shape, footer);
        Assert.DoesNotContain(explicitReport.Warnings, warning => warning.Code == "HeaderFooterPageBoundsClipped");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RunningDashGapsUseThePdfPresetPatternForPaintBounds(bool footer) {
        var report = new PdfConversionReport();
        var shape = OfficeShape.Path(12, 10, OfficePathCommand.MoveTo(0, 5), OfficePathCommand.LineTo(9, 5));
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 2; shape.StrokeDashStyle = OfficeStrokeDashStyle.Dash;
        byte[] bytes = Running(new PdfOptions { PageWidth = 9.5, PageHeight = 100, MarginLeft = 2, MarginRight = 0,
            MarginTop = 20, MarginBottom = 20, HeaderOffsetY = 0, FooterOffsetY = 0 }.ReportDiagnosticsTo(report, "OfficeIMO.Tests"), shape, footer);
        Assert.NotEmpty(bytes);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "HeaderFooterPageBoundsClipped");
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void ExplicitMiterLimitSurvivesPdfPaintingOfOpenMarkers(int placement) {
        var shape = OfficeShape.Line(0, 0, 60, 0);
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 2;
        shape.StrokeLineJoin = OfficeStrokeLineJoin.Miter;
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 8, 25);
        var options = new PdfOptions { CompressContentStreams = false, PageWidth = 100, PageHeight = 100,
            MarginLeft = 10, MarginRight = 10, MarginTop = 20, MarginBottom = 20, HeaderOffsetY = 0, FooterOffsetY = 0 };
        byte[] bytes = PaintAtPlacement(options, shape, placement);
        Assert.Contains("4 M", Encoding.ASCII.GetString(bytes));
        var strokes = Shapes(PdfReadDocument.Open(bytes).Pages[0].ToDrawing()).Where(item => item.Shape.StrokeColor == OfficeColor.Red).ToArray();
        Assert.True(strokes.Length >= 2);
        Assert.All(strokes, item => Assert.Equal(4, item.Shape.StrokeMiterLimit));
    }

    private static byte[] Running(PdfOptions options, OfficeShape shape, bool footer) => PdfDocument.Create(options).Compose(compose => compose.Page(page => {
        if (footer) page.Footer(content => content.Shape(shape));
        else page.Header(content => content.Shape(shape));
        page.Content(content => content.Spacer(10));
    })).ToBytes();

    private static byte[] PaintAtPlacement(PdfOptions options, OfficeShape shape, int placement) => placement == 0
        ? PdfDocument.Create(options).Compose(compose => compose.Page(page => page.Canvas(canvas =>
            canvas.Drawing(new OfficeDrawing(100, 100).AddShape(shape, 10, 50), 0, 0, 100, 100)))).ToBytes()
        : Running(options, shape, placement == 2);

    private static IEnumerable<OfficeDrawingShape> Shapes(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement item in drawing.Elements) {
            if (item is OfficeDrawingShape shape) yield return shape;
            if (item is OfficeDrawingGroup group) foreach (var child in Shapes(group.Drawing)) yield return child;
        }
    }
}
