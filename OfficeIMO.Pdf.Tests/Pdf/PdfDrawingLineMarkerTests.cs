using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfDrawingLineMarkerTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DrawingMarkersReachPdfAsFilledAndOpenGeometryWithStrokeOpacity(bool path) {
        OfficeColor ink = OfficeColor.FromRgb(204, 32, 64);
        OfficeShape shape = path ? OfficeShape.Path(80, 1, new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(40, 0), OfficePathCommand.LineTo(80, 0) })
            : OfficeShape.Line(0, 0, 80, 0);
        shape.StrokeColor = ink; shape.StrokeWidth = 2; shape.StrokeOpacity = 0.5;
        if (path) shape.FillColor = OfficeColor.Blue;
        shape.FillOpacity = 0.1;
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 10, 14);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 18, 20);
        shape.Transform = OfficeTransform.RotateDegrees(30, 40, 0);
        var drawing = new OfficeDrawing(140, 100).AddShape(shape, 20, 50);
        PdfDocument pdf = PdfDocument.Create(new PdfCreateOptions { BackgroundColor = null }).Compose(builder => builder.Page(page => page.Size(140, 100).Margin(0)
            .Canvas(canvas => canvas.Drawing(drawing, 0, 0, 140, 100))));
        OfficeDrawing read = PdfReadDocument.Open(pdf.ToBytes()).Pages[0].ToDrawing();
        OfficeShape[] painted = Shapes(read).Select(item => item.Shape).ToArray();
        OfficeShape marker = Assert.Single(painted, item => item.FillColor == ink);
        Assert.Equal(0.5, marker.FillOpacity.GetValueOrDefault(1), 5);
        Assert.Equal(2, painted.Count(item => item.StrokeColor == ink));
        Assert.All(painted.Where(item => item.StrokeColor == ink), item => Assert.Equal(0.5, item.StrokeOpacity.GetValueOrDefault(1), 5));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GradientMarkersRetainStrokeBrushOpacityAndClipWithoutRasterFallback(bool radial) {
        var shape = OfficeShape.Path(80, 40, OfficePathCommand.MoveTo(10, 20), OfficePathCommand.LineTo(70, 20));
        shape.StrokeWidth = 4;
        shape.StrokeOpacity = .5;
        if (radial) shape.StrokeRadialGradient = new OfficeRadialGradient(0, .5, 0, 0, .5, 1,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        else shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 20, 20);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 20, 20);
        shape.ClipPath = OfficeClipPath.Rectangle(60, 40);
        var drawing = new OfficeDrawing(140, 100).AddShape(shape, 20, 20);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { BackgroundColor = null }).Compose(builder => builder.Page(page => page.Size(140, 100).Margin(0)
            .Canvas(canvas => canvas.Drawing(drawing, 0, 0, 140, 100)))).ToBytes();
        Assert.DoesNotContain("/Subtype /Image", System.Text.Encoding.ASCII.GetString(bytes));
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes));
        OfficeColor filledEnd = image.GetPixel(43, 36), openEnd = image.GetPixel(75, 33);
        Assert.InRange(filledEnd.A, (byte)120, (byte)135);
        Assert.True(filledEnd.R > filledEnd.B);
        Assert.InRange(openEnd.A, (byte)120, (byte)135);
        Assert.True(openEnd.B > openEnd.R);
        Assert.Equal(0, image.GetPixel(76, 36).A);
        Assert.Equal(0, image.GetPixel(84, 37).A);
        Assert.Null(shape.FillGradient);
        Assert.Null(shape.FillRadialGradient);
        Assert.Equal(.5, shape.StrokeOpacity);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void HeaderAndFooterPaintBothMarkerKindsWithTheSourceStrokeBrush(bool footer, bool gradient) {
        var shape = OfficeShape.Path(80, 40, OfficePathCommand.MoveTo(10, 20), OfficePathCommand.LineTo(70, 20));
        shape.StrokeWidth = 4;
        shape.StrokeOpacity = .5;
        if (gradient) shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        else shape.StrokeColor = OfficeColor.Red;
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 20, 20);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 20, 20);
        shape.Transform = OfficeTransform.RotateDegrees(30, 40, 20);
        shape.ClipPath = OfficeClipPath.Rectangle(80, 40);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { BackgroundColor = null,
            PageWidth = global::OfficeIMO.OfficeLength.Points(200), PageHeight = global::OfficeIMO.OfficeLength.Points(400), MarginTop = global::OfficeIMO.OfficeLength.Points(100), MarginBottom = global::OfficeIMO.OfficeLength.Points(100),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(20), MarginRight = global::OfficeIMO.OfficeLength.Points(20)
        }).Compose(compose => compose.Page(page => {
            if (footer) page.Footer(content => content.Shape(shape));
            else page.Header(content => content.Shape(shape));
            page.Content(content => content.Spacer(20));
        })).ToBytes();
        Assert.DoesNotContain("/Subtype /Image", System.Text.Encoding.ASCII.GetString(bytes));
        OfficeShape[] paint = Shapes(PdfReadDocument.Open(bytes).Pages[0].ToDrawing()).Select(item => item.Shape)
            .Where(item => item.StrokeColor == OfficeColor.Red || item.FillColor == OfficeColor.Red || item.FillGradient != null).ToArray();
        Assert.Equal(3, paint.Length);
        Assert.All(paint, item => Assert.Equal(.5, (item.StrokeColor.HasValue ? item.StrokeOpacity : item.FillOpacity).GetValueOrDefault(1), 5));
        if (gradient) Assert.All(paint, item => Assert.NotNull(item.FillGradient));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void RunningMarkerClippingReportsPaintBeyondThePageButHonorsTheShapeClip(bool footer, bool clipped) {
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions { BackgroundColor = null,
            PageWidth = global::OfficeIMO.OfficeLength.Points(100), PageHeight = global::OfficeIMO.OfficeLength.Points(100), MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10),
            MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20), HeaderOffsetY = global::OfficeIMO.OfficeLength.Points(0), FooterOffsetY = global::OfficeIMO.OfficeLength.Points(0)
        }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");
        var shape = OfficeShape.Path(70, 1, OfficePathCommand.MoveTo(0, .5), OfficePathCommand.LineTo(70, .5));
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 2;
        shape.StrokeLineCap = OfficeStrokeLineCap.Butt;
        shape.StrokeLineJoin = OfficeStrokeLineJoin.Round;
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 60, 30);
        if (clipped) shape.ClipPath = OfficeClipPath.Rectangle(70, 1);
        byte[] bytes = PdfDocument.Create(options).Compose(compose => compose.Page(page => {
            if (footer) page.Footer(content => content.Shape(shape));
            else page.Header(content => content.Shape(shape));
            page.Content(content => content.Spacer(10));
        })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        var warnings = report.Warnings.Where(warning => warning.Code == "header-footer-page-bounds-clipped").ToArray();
        if (clipped) Assert.Empty(warnings);
        else {
            var warning = Assert.Single(warnings);
            Assert.Equal(footer ? "Footer" : "Header", warning.Source);
            Assert.True(footer ? warning.LayoutDiagnostic!.Y < 0 :
                warning.LayoutDiagnostic!.Y + warning.LayoutDiagnostic.Height > options.PageHeight.ToPoints());
            Assert.True(report.HasLoss);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RunningGradientArrowReportsPhysicalPageClipping(bool radial) {
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions { BackgroundColor = null, PageWidth = global::OfficeIMO.OfficeLength.Points(100), PageHeight = global::OfficeIMO.OfficeLength.Points(100), MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10),
            MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20), HeaderOffsetY = global::OfficeIMO.OfficeLength.Points(0) }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");
        var shape = OfficeShape.Path(70, 1, OfficePathCommand.MoveTo(0, .5), OfficePathCommand.LineTo(70, .5));
        shape.StrokeWidth = 2;
        if (radial) shape.StrokeRadialGradient = new OfficeRadialGradient(0, .5, 0, 0, .5, 1,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        else shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 60, 30);
        byte[] bytes = PdfDocument.Create(options).Compose(b => b.Page(page => {
            page.Header(c => c.Shape(shape)); page.Content(c => c.Spacer(10));
        })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        Assert.Contains(report.Warnings, w => w.Code == "header-footer-page-bounds-clipped" && w.Source == "Header");
        Assert.True(report.HasLoss);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void RunningPaintBoundsUseTheEffectiveRadialGradient(bool radialVisible, bool marker) {
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions { BackgroundColor = null, PageWidth = global::OfficeIMO.OfficeLength.Points(100), PageHeight = global::OfficeIMO.OfficeLength.Points(100), MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10),
            MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20), HeaderOffsetY = global::OfficeIMO.OfficeLength.Points(0) }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");
        var linear = OfficeLinearGradient.Horizontal(radialVisible ? OfficeColor.Transparent : OfficeColor.Red,
            radialVisible ? OfficeColor.Transparent : OfficeColor.Blue);
        var radial = new OfficeRadialGradient(0, .5, 0, 0, .5, 1,
            new OfficeGradientStop(0, radialVisible ? OfficeColor.Red : OfficeColor.Transparent),
            new OfficeGradientStop(1, radialVisible ? OfficeColor.Blue : OfficeColor.Transparent));
        var shape = marker
            ? OfficeShape.Path(70, 1, OfficePathCommand.MoveTo(0, .5), OfficePathCommand.LineTo(70, .5))
            : OfficeShape.Rectangle(70, 1);
        if (marker) {
            shape.StrokeWidth = 2; shape.StrokeGradient = linear; shape.StrokeRadialGradient = radial;
            shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 60, 30);
        } else {
            shape.FillGradient = linear; shape.FillRadialGradient = radial;
            shape.Transform = OfficeTransform.Translate(0, -30);
        }
        byte[] bytes = PdfDocument.Create(options).Compose(b => b.Page(page => {
            page.Header(c => c.Shape(shape)); page.Content(c => c.Spacer(10));
        })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        Assert.Equal(radialVisible, report.Warnings.Any(w => w.Code == "header-footer-page-bounds-clipped"));
        Assert.Equal(radialVisible, report.HasLoss);
    }

    [Fact]
    public void RunningGradientArrowReportsDefaultRoundCapClipping() {
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions { BackgroundColor = null, PageWidth = global::OfficeIMO.OfficeLength.Points(100), PageHeight = global::OfficeIMO.OfficeLength.Points(100), MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10),
            MarginTop = global::OfficeIMO.OfficeLength.Points(31), MarginBottom = global::OfficeIMO.OfficeLength.Points(20), HeaderOffsetY = global::OfficeIMO.OfficeLength.Points(0) }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");
        var shape = OfficeShape.Path(70, 1, OfficePathCommand.MoveTo(0, .5), OfficePathCommand.LineTo(70, .5));
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 4; shape.StrokeLineJoin = OfficeStrokeLineJoin.Round;
        shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 60, 30);
        byte[] bytes = PdfDocument.Create(options).Compose(b => b.Page(page => {
            page.Header(c => c.Shape(shape)); page.Content(c => c.Spacer(10));
        })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        Assert.Contains(report.Warnings, w => w.Code == "header-footer-page-bounds-clipped" && w.Source == "Header");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FreeformMarkerClipRetainsOnlyPaintInsideThePage(bool gradient) {
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions { BackgroundColor = null, PageWidth = global::OfficeIMO.OfficeLength.Points(100), PageHeight = global::OfficeIMO.OfficeLength.Points(100), MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10),
            MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20), FooterOffsetY = global::OfficeIMO.OfficeLength.Points(0) }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");
        var shape = OfficeShape.Path(70, 30, OfficePathCommand.MoveTo(0, .5), OfficePathCommand.LineTo(70, .5));
        shape.StrokeWidth = 2;
        if (gradient) shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        else shape.StrokeColor = OfficeColor.Red;
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 60, 30);
        shape.Transform = OfficeTransform.Translate(0, 30);
        shape.ClipPath = OfficeClipPath.Path(OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(70, 0),
            OfficePathCommand.LineTo(70, 30), OfficePathCommand.Close());
        byte[] bytes = PdfDocument.Create(options).Compose(b => b.Page(page => {
            page.Footer(c => c.Shape(shape)); page.Content(c => c.Spacer(10));
        })).ToBytes();
        Assert.NotEmpty(bytes);
        Assert.DoesNotContain(report.Warnings, w => w.Code == "header-footer-page-bounds-clipped");
        Assert.False(report.HasLoss);
    }

    [Fact]
    public void RunningGradientShaftReportsPhysicalPageClipping() {
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions { BackgroundColor = null, PageWidth = global::OfficeIMO.OfficeLength.Points(100), PageHeight = global::OfficeIMO.OfficeLength.Points(100), MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10),
            MarginTop = global::OfficeIMO.OfficeLength.Points(1), MarginBottom = global::OfficeIMO.OfficeLength.Points(20), HeaderOffsetY = global::OfficeIMO.OfficeLength.Points(0) }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");
        var shape = OfficeShape.Path(70, 1, OfficePathCommand.MoveTo(0, .5), OfficePathCommand.LineTo(70, .5));
        shape.StrokeWidth = 4; shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        byte[] bytes = PdfDocument.Create(options).Compose(b => b.Page(page => {
            page.Header(c => c.Shape(shape)); page.Content(c => c.Spacer(10));
        })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        Assert.Contains(report.Warnings, w => w.Code == "header-footer-page-bounds-clipped");
    }

    [Fact]
    public void RunningPaintInspectionLimitReportsUnassessedClipping() {
        var commands = new List<OfficePathCommand> { OfficePathCommand.MoveTo(0, 0) };
        for (int i = 1; i <= 4100; i++) commands.Add(OfficePathCommand.LineTo((i & 1) == 0 ? 0 : 70, i * 30D / 4100D));
        commands.Add(OfficePathCommand.Close());
        var shape = OfficeShape.Path(70, 30, OfficePathCommand.MoveTo(0, .5), OfficePathCommand.LineTo(70, .5));
        shape.StrokeWidth = 2; shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        shape.ClipPath = OfficeClipPath.Path(commands);
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions { BackgroundColor = null, PageWidth = global::OfficeIMO.OfficeLength.Points(100), PageHeight = global::OfficeIMO.OfficeLength.Points(100), MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10),
            MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20), HeaderOffsetY = global::OfficeIMO.OfficeLength.Points(0) }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");
        byte[] bytes = PdfDocument.Create(options).Compose(b => b.Page(page => {
            page.Header(c => c.Shape(shape)); page.Content(c => c.Spacer(10));
        })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        var warning = Assert.Single(report.Warnings, w => w.Code == "header-footer-paint-bounds-unmeasured");
        Assert.Equal(OfficeConversionLossKind.Unassessed, warning.LossKind);
        Assert.Null(warning.LayoutDiagnostic);
        Assert.DoesNotContain(report.Warnings, w => w.Code == "header-footer-page-bounds-clipped");
        Assert.True(report.HasLoss);
    }

    [Fact]
    public void MarkerTangentsSelectTheFirstAndLastEligibleOpenContours() {
        var shape = OfficeShape.Path(100, 60,
            OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(1, 1), OfficePathCommand.Close(),
            OfficePathCommand.MoveTo(20, 20), OfficePathCommand.CubicBezierTo(20, 20, 40, 20, 40, 40),
            OfficePathCommand.MoveTo(70, 10), OfficePathCommand.LineTo(70, 10),
            OfficePathCommand.MoveTo(80, 40), OfficePathCommand.LineTo(81, 41), OfficePathCommand.Close(),
            OfficePathCommand.MoveTo(90, 50));
        shape.StrokeWidth = 1; shape.StrokeColor = OfficeColor.Red;
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 10, 10);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 10, 10);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { BackgroundColor = null }).Compose(b => b.Page(p => p.Size(100, 60).Margin(0)
            .Canvas(c => c.Drawing(new OfficeDrawing(100, 60).AddShape(shape, 0, 0), 0, 0, 100, 60)))).ToBytes();
        Assert.Equal(2, Shapes(PdfReadDocument.Open(bytes).Pages[0].ToDrawing()).Count(item => item.Shape.FillColor == OfficeColor.Red));
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes));
        Assert.True(image.GetPixel(29, 17).R > 200 && image.GetPixel(29, 17).G < 50);
        Assert.True(image.GetPixel(43, 31).R > 200 && image.GetPixel(43, 31).G < 50);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DrawingAfterCloseRetainsItsImplicitOpenContourMarkers(bool svg) {
        var shape = OfficeShape.Path(40, 40, OfficePathCommand.MoveTo(10, 10), OfficePathCommand.LineTo(20, 10),
            OfficePathCommand.Close(), OfficePathCommand.LineTo(30, 10), OfficePathCommand.LineTo(30, 30));
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 1;
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 10, 10);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 10, 10);
        if (svg) {
            var xml = System.Xml.Linq.XElement.Parse(OfficeDrawingSvgExporter.ToSvg(new OfficeDrawing(40, 40).AddShape(shape, 0, 0)));
            Assert.Equal(2, xml.Descendants().Count(element => element.Name.LocalName == "polygon"));
        } else {
            byte[] bytes = PdfDocument.Create(new PdfCreateOptions { BackgroundColor = null }).Shape(shape).ToBytes();
            Assert.Equal(2, Shapes(PdfReadDocument.Open(bytes).Pages[0].ToDrawing()).Count(item => item.Shape.FillColor == OfficeColor.Red));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TransformScaledGradientCurvesRetainBoundedVectorOutput(bool radial) {
        var commands = new List<OfficePathCommand> { OfficePathCommand.MoveTo(0, 1) };
        for (int i = 0; i < 500; i++) commands.Add(OfficePathCommand.CubicBezierTo(i, 2, i + 1, 0, i + 1, 1));
        var shape = OfficeShape.Path(500, 2, commands);
        shape.Transform = OfficeTransform.Scale(1E14, 1E14); shape.StrokeWidth = 1;
        if (radial) shape.StrokeRadialGradient = new OfficeRadialGradient(0, 0, 0, .5, .5, 1,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        else shape.StrokeGradient = new OfficeLinearGradient(0, 0, 1, 1,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { BackgroundColor = null, PageWidth = global::OfficeIMO.OfficeLength.Points(1000), PageHeight = global::OfficeIMO.OfficeLength.Points(1000),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10) }).Shape(shape).ToBytes();
        Assert.DoesNotContain("/Subtype /Image", System.Text.Encoding.ASCII.GetString(bytes));
        Assert.Single(PdfReadDocument.Open(bytes).Pages);
    }

    [Fact]
    public void NativeCurvedMarkersDoNotRequirePhysicalResolutionFlattening() {
        var commands = new List<OfficePathCommand> { OfficePathCommand.MoveTo(0, 1) };
        for (int i = 0; i < 500; i++) commands.Add(OfficePathCommand.CubicBezierTo(i, 2, i + 1, 0, i + 1, 1));
        var shape = OfficeShape.Path(500, 2, commands);
        shape.Transform = OfficeTransform.Scale(1E14, 1E14);
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 1;
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 1, 1);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 1, 1);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { BackgroundColor = null, PageWidth = global::OfficeIMO.OfficeLength.Points(1000), PageHeight = global::OfficeIMO.OfficeLength.Points(1000),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10) }).Shape(shape).ToBytes();
        string text = OfficeIMO.Tests.Pdf.PdfTestContentText.ReadDecoded(bytes);
        Assert.Contains(" c", text);
        Assert.DoesNotContain("/Subtype /Image", text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnmeasuredRadialMarkerPaintStillExpandsTheLinkGeometry(bool shaft) {
        var shape = OfficeShape.Line(0, 0, 80, 0);
        shape.StrokeWidth = shaft ? 20 : 2;
        shape.StrokeLineCap = OfficeStrokeLineCap.Round;
        shape.StrokeRadialGradient = new OfficeRadialGradient(0, 0, 0, .5, .5, 1,
            new OfficeGradientStop(0, OfficeColor.Transparent),
            new OfficeGradientStop(1, new OfficeColor(0, 0, 255, 128)));
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Diamond, shaft ? 2 : 40, shaft ? 2 : 40);
        var options = new PdfCreateOptions { BackgroundColor = null, PageWidth = global::OfficeIMO.OfficeLength.Points(200), PageHeight = global::OfficeIMO.OfficeLength.Points(200), MarginLeft = global::OfficeIMO.OfficeLength.Points(30), MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30), MarginBottom = global::OfficeIMO.OfficeLength.Points(30) };
        byte[] bytes = PdfDocument.Create(options).Shape(shape, linkUri: "https://example.test/marker").ToBytes();
        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(bytes).LinkAnnotations);
        Assert.True(link.Y2 - link.Y1 >= (shaft ? 19.99 : 39.99));
        if (shaft) { Assert.True(link.X1 <= 20.01); Assert.True(link.X2 >= 119.99); }
    }

    private static IEnumerable<OfficeDrawingShape> Shapes(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingShape shape) yield return shape;
            if (element is OfficeDrawingGroup group) foreach (OfficeDrawingShape child in Shapes(group.InnerDrawing)) yield return child;
            if (element is OfficeDrawingEffectGroup effect) foreach (OfficeDrawingShape child in Shapes(effect.InnerDrawing)) yield return child;
        }
    }
}
