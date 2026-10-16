using System;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfGraphicsGradientAlphaTests {
    [Theory]
    [InlineData(false, OfficeGradientColorInterpolation.Srgb)]
    [InlineData(true, OfficeGradientColorInterpolation.Srgb)]
    [InlineData(false, OfficeGradientColorInterpolation.LinearRgb)]
    [InlineData(true, OfficeGradientColorInterpolation.LinearRgb)]
    public void CanvasGradientTransparentEndpointRetainsVisibleHue(bool radial, OfficeGradientColorInterpolation interpolation) {
        var shape = OfficeShape.Rectangle(80D, 80D);
        shape.StrokeWidth = 0D;
        if (radial) shape.FillRadialGradient = OfficeRadialGradient.Centered(OfficeColor.Red, OfficeColor.FromRgba(0, 0, 255, 0)).WithColorInterpolation(interpolation);
        else shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.FromRgba(0, 0, 255, 0)).WithColorInterpolation(interpolation);
        var drawing = new OfficeDrawing(80D, 80D).AddShape(shape, 0D, 0D);
        byte[] pdf = PdfDocument.Create(Options()).Drawing(drawing).ToBytes();
        OfficeRasterImage actual = Render(pdf);
        OfficeColor middle = actual.GetPixel(radial ? 60 : 40, 40);
        Assert.InRange(middle.R, 250, 255);
        Assert.InRange(middle.G, 115, 140);
        Assert.InRange(middle.B, 115, 140);
        // Independent shared-raster route must agree about the producer intent.
        OfficeColor canvas = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White).GetPixel(radial ? 60 : 40, 40);
        Assert.InRange(Math.Abs(canvas.R - middle.R), 0, 3);
        Assert.InRange(Math.Abs(canvas.B - middle.B), 0, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PartiallyTransparentEndpointsUsePremultipliedColor(bool radial) {
        var first = OfficeColor.FromRgba(255, 0, 0, 64);
        var second = OfficeColor.FromRgba(0, 0, 255, 192);
        var shape = OfficeShape.Rectangle(80D, 80D); shape.StrokeWidth = 0D;
        if (radial) shape.FillRadialGradient = OfficeRadialGradient.Centered(first, second);
        else shape.FillGradient = OfficeLinearGradient.Horizontal(first, second);
        byte[] pdf = PdfDocument.Create(Options()).Shape(shape).ToBytes();
        OfficeColor middle = Render(pdf).GetPixel(radial ? 60 : 40, 40);
        // At t≈0.5: alpha≈0.5, unassociated red≈0.25 and blue≈0.75.
        Assert.InRange(middle.R, 153, 166);
        Assert.InRange(middle.G, 120, 133);
        Assert.InRange(middle.B, 217, 230);
    }

    [Fact]
    public void SvgGradientRetainsStraightRgbAndDoesNotReuseCanvasShading() {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='80' height='80'><defs><linearGradient id='g'><stop stop-color='red'/><stop offset='1' stop-color='blue' stop-opacity='0'/></linearGradient></defs><rect width='80' height='80' fill='url(#g)'/></svg>");
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out OfficeDrawing? imported, out int unsupported));
        Assert.Equal(0, unsupported);
        var canvas = OfficeShape.Rectangle(80D, 80D); canvas.StrokeWidth = 0D;
        canvas.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.FromRgba(0, 0, 255, 0));
        byte[] pdf = PdfDocument.Create(Options()).Drawing(imported!).Shape(canvas).ToBytes();
        OfficeRasterImage raster = Render(pdf);
        OfficeColor svgMiddle = raster.GetPixel(40, 40), canvasMiddle = raster.GetPixel(40, 120);
        Assert.InRange(svgMiddle.R, 180, 200);
        Assert.InRange(svgMiddle.B, 180, 200);
        Assert.InRange(canvasMiddle.R, 250, 255);
        Assert.InRange(canvasMiddle.B, 120, 140);
    }

    [Theory]
    [InlineData(0.5D)]
    [InlineData(0.99D)]
    public void TransparentInteriorStopKeepsBothAdjacentVisibleHues(double transparentOffset) {
        var shape = OfficeShape.Rectangle(80D, 80D); shape.StrokeWidth = 0D;
        shape.FillGradient = new OfficeLinearGradient(0D, 0.5D, 1D, 0.5D, new[] {
            new OfficeGradientStop(0D, OfficeColor.Red),
            new OfficeGradientStop(transparentOffset, OfficeColor.Transparent),
            new OfficeGradientStop(1D, OfficeColor.Blue)
        });
        byte[] pdf = PdfDocument.Create(Options()).Shape(shape).ToBytes();
        AssertStrictStitchingBounds(pdf);
        OfficeRasterImage raster = Render(pdf);
        int leftX = (int)(40D * transparentOffset);
        int rightX = Math.Min(79, (int)(80D * (transparentOffset + (1D - transparentOffset) / 2D)));
        OfficeColor left = raster.GetPixel(leftX, 40), right = raster.GetPixel(rightX, 40);
        double leftAlpha = 1D - ((leftX + 0.5D) / 80D) / transparentOffset;
        double rightAlpha = (((rightX + 0.5D) / 80D) - transparentOffset) / (1D - transparentOffset);
        Assert.InRange(left.R, 250, 255);
        Assert.InRange(Math.Abs(left.B - 255D * (1D - leftAlpha)), 0D, 3D);
        Assert.InRange(right.B, 250, 255);
        Assert.InRange(Math.Abs(right.R - 255D * (1D - rightAlpha)), 0D, 3D);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ShortGradientIntervalsRetainStrictBoundsForAxialRadialAndPrint(bool radial, bool convertToPrint) {
        PdfOptions options = Options();
        if (convertToPrint) options.ConfigurePdfXGroundwork(PdfComplianceProfile.PdfX4,
            IccMabTestProfiles.CreateCmykLab8Bidirectional(), "Managed test CMYK", PdfTrappingStatus.False);
        var document = PdfDocument.Create(options);
        var cases = new[] {
            new[] { new OfficeGradientStop(0D, OfficeColor.Red), new OfficeGradientStop(0.99D, OfficeColor.Transparent), new OfficeGradientStop(1D, OfficeColor.Blue) },
            new[] { new OfficeGradientStop(0D, OfficeColor.Red), new OfficeGradientStop(0.9999999999D, OfficeColor.Transparent), new OfficeGradientStop(1D, OfficeColor.Blue) },
            new[] { new OfficeGradientStop(0D, OfficeColor.Red), new OfficeGradientStop(0.99D, OfficeColor.Green), new OfficeGradientStop(0.9900000001D, OfficeColor.Blue), new OfficeGradientStop(1D, OfficeColor.White) }
        };
        foreach (OfficeGradientStop[] stops in cases) {
            var shape = OfficeShape.Rectangle(80D, 40D); shape.StrokeWidth = 0D;
            if (radial) shape.FillRadialGradient = new OfficeRadialGradient(0.5D, 0.5D, 0D, 0.5D, 0.5D, 0.5D, stops);
            else shape.FillGradient = new OfficeLinearGradient(0D, 0.5D, 1D, 0.5D, stops);
            document.Shape(shape);
        }
        AssertStrictStitchingBounds(document.ToBytes());
    }

    private static void AssertStrictStitchingBounds(byte[] pdf) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        var functions = objects.Values.Select(value => value.Value).OfType<PdfDictionary>()
            .Where(dictionary => dictionary.Get<PdfNumber>("FunctionType")?.Value == 3D).ToArray();
        Assert.NotEmpty(functions);
        foreach (PdfDictionary function in functions) {
            double previous = 0D;
            foreach (PdfObject value in Assert.IsType<PdfArray>(function.Items["Bounds"]).Items) {
                double offset = Assert.IsType<PdfNumber>(value).Value;
                Assert.True(offset > previous && offset < 1D, $"Expected a strictly increasing interior bound after {previous:R}, got {offset:R}.");
                previous = offset;
            }
        }
    }

    private static OfficeRasterImage Render(byte[] pdf) => OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
    private static PdfOptions Options() => new PdfOptions { PageWidth = 80D, PageHeight = 160D,
        MarginLeft = 0D, MarginRight = 0D, MarginTop = 0D, MarginBottom = 0D };
}
