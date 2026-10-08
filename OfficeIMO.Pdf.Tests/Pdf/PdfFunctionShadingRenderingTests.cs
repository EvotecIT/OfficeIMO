using System;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFunctionShadingRenderingTests {
    [Theory]
    [InlineData(1, OfficeGradientSpreadMode.Repeat)]
    [InlineData(2, OfficeGradientSpreadMode.Reflect)]
    public void ReopenedPeriodicFieldPreservesResolutionAndAlpha(int scale, OfficeGradientSpreadMode spread) {
        var shape = Shape(spread);
        var expected = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(60, 40).AddShape(shape, 0, 0),
            scale: scale, background: OfficeColor.White);
        var page = PdfReadDocument.Open(Document(shape)).Pages[0];
        var exported = page.ExportImage(OfficeImageExportFormat.Png, new PdfImageExportOptions { Scale = scale });
        Assert.True(OfficeRasterImageDecoder.TryDecode(exported.Bytes, out var actual));
        Assert.Equal(expected.Width, actual!.Width); Assert.Equal(expected.Height, actual.Height);
        foreach (var point in new[] { (7, 8), (16, 22), (29, 18), (46, 32), (55, 7) }) {
            var a = actual.GetPixel(point.Item1 * scale, point.Item2 * scale);
            var b = expected.GetPixel(point.Item1 * scale, point.Item2 * scale);
            Assert.InRange(Math.Abs(a.R - b.R), 0, 3);
            Assert.InRange(Math.Abs(a.G - b.G), 0, 3);
            Assert.InRange(Math.Abs(a.B - b.B), 0, 3);
        }
    }

    [Theory]
    [InlineData(128)]
    [InlineData(1024)]
    public void PeriodicFieldsReopenAcrossSupportedStopCount(int stopCount) {
        var shape = Shape(OfficeGradientSpreadMode.Reflect);
        var stops = new OfficeGradientStop[stopCount];
        for (int i = 0; i < stopCount; i++) stops[i] = new OfficeGradientStop(i / (double)(stopCount - 1),
            OfficeColor.FromRgba((byte)(i % 256), (byte)((i * 7) % 256), (byte)((i * 17) % 256), (byte)(64 + i % 192)));
        shape.FillRadialGradient = new OfficeRadialGradient(.25, .5, 0, .25, .5, .08, stops)
            .WithSpreadMode(OfficeGradientSpreadMode.Reflect);
        var page = PdfReadDocument.Open(Document(shape)).Pages[0];
        var exported = page.ExportImage(OfficeImageExportFormat.Png);
        Assert.DoesNotContain(exported.Diagnostics, d => d.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.True(OfficeRasterImageDecoder.TryDecode(exported.Bytes, out var actual));
        var expected = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(60, 40).AddShape(shape, 0, 0),
            background: OfficeColor.White);
        var a = actual!.GetPixel(29, 18); var b = expected.GetPixel(29, 18);
        Assert.InRange(Math.Abs(a.R - b.R), 0, 3);
        Assert.InRange(Math.Abs(a.G - b.G), 0, 3);
        Assert.InRange(Math.Abs(a.B - b.B), 0, 3);
    }

    [Fact]
    public void DefaultLimitsRenderLetterPageWithPeriodicColorAndAlphaAt96Dpi() {
        var shape = Shape(OfficeGradientSpreadMode.Reflect, 612, 792);
        var page = PdfReadDocument.Open(Document(shape)).Pages[0];
        var exported = page.ExportImage(OfficeImageExportFormat.Png, new PdfImageExportOptions { Scale = 4D / 3D });
        Assert.True(OfficeRasterImageDecoder.TryDecode(exported.Bytes, out var actual));
        Assert.Equal(816, actual!.Width); Assert.Equal(1056, actual.Height);
        var expected = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(612, 792).AddShape(shape, 0, 0),
            scale: 4D / 3D, background: OfficeColor.White);
        var a = actual.GetPixel(400, 500); var b = expected.GetPixel(400, 500);
        Assert.InRange(Math.Abs(a.R - b.R), 0, 3);
        Assert.InRange(Math.Abs(a.G - b.G), 0, 3);
        Assert.InRange(Math.Abs(a.B - b.B), 0, 3);
    }

    [Fact]
    public void ExportDownscalesFunctionColorAndMaskBeforeChargingIntermediatePixels() {
        var shape = Shape(OfficeGradientSpreadMode.Reflect);
        var expected = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(60, 40).AddShape(shape, 0, 0),
            scale: 1, background: OfficeColor.White);
        var page = PdfReadDocument.Open(Document(shape), new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxFunctionShadingPixels = 4800 } }).Pages[0];
        var exported = page.ExportImage(OfficeImageExportFormat.Png, new PdfImageExportOptions {
            Scale = 4, MaximumOutputWidth = 60,
            RasterOverflowBehavior = OfficeRasterOverflowBehavior.ReduceScale
        });
        Assert.True(OfficeRasterImageDecoder.TryDecode(exported.Bytes, out var actual));
        Assert.Equal(60, actual!.Width); Assert.Equal(40, actual.Height);
        var a = actual.GetPixel(29, 18); var b = expected.GetPixel(29, 18);
        Assert.InRange(Math.Abs(a.R - b.R), 0, 3);
        Assert.InRange(Math.Abs(a.G - b.G), 0, 3);
        Assert.InRange(Math.Abs(a.B - b.B), 0, 3);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void FunctionFieldStopsAtConfiguredPixelOrEvaluationLimit(bool pixels) {
        var limits = pixels ? new PdfReadLimits { MaxFunctionShadingPixels = 2 }
            : new PdfReadLimits { MaxFunctionShadingEvaluationWork = 2 };
        var page = PdfReadDocument.Open(Document(Shape(OfficeGradientSpreadMode.Repeat)),
            new PdfLoadOptions { Limits = limits }).Pages[0];
        var error = Assert.Throws<PdfReadLimitException>(() => page.ExportImage(OfficeImageExportFormat.Png));
        Assert.Equal(pixels ? PdfReadLimitKind.FunctionShadingPixels : PdfReadLimitKind.FunctionShadingEvaluationWork, error.Kind);
        Assert.Equal(2, error.Limit);
    }

    private static OfficeShape Shape(OfficeGradientSpreadMode spread, double width = 60, double height = 40) {
        var shape = OfficeShape.Rectangle(width, height); shape.StrokeWidth = 0;
        shape.FillRadialGradient = new OfficeRadialGradient(.25, .5, 0, .25, .5, .08,
            new OfficeGradientStop(0, OfficeColor.FromRgba(255, 0, 0, 64)),
            new OfficeGradientStop(1, OfficeColor.FromRgba(0, 0, 255, 192))).WithSpreadMode(spread);
        return shape;
    }

    private static byte[] Document(OfficeShape shape) =>
        PdfDocument.Create(new PdfOptions { PageWidth = shape.Width, PageHeight = shape.Height,
            MarginTop = 0, MarginBottom = 0, MarginLeft = 0, MarginRight = 0 })
            .Compose(c => c.Page(p => p.Content(content => content.Shape(shape)))).ToBytes();
}
