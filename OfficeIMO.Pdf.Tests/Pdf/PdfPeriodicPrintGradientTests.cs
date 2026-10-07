using System;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPeriodicPrintGradientTests {
    [Theory]
    [InlineData(OfficeGradientSpreadMode.Repeat, OfficeGradientColorInterpolation.Srgb)]
    [InlineData(OfficeGradientSpreadMode.Reflect, OfficeGradientColorInterpolation.Srgb)]
    [InlineData(OfficeGradientSpreadMode.Repeat, OfficeGradientColorInterpolation.LinearRgb)]
    [InlineData(OfficeGradientSpreadMode.Reflect, OfficeGradientColorInterpolation.LinearRgb)]
    public void PeriodicCmykFieldMatchesTheConfiguredIccTransform(
        OfficeGradientSpreadMode spread, OfficeGradientColorInterpolation interpolation) {
        var stops = new[] { new OfficeGradientStop(0, OfficeColor.FromRgb(230, 40, 80)),
            new OfficeGradientStop(.4, OfficeColor.FromRgb(50, 210, 90)),
            new OfficeGradientStop(1, OfficeColor.FromRgb(30, 60, 220)) };
        var transform = PdfPrintColorTransform.Create(Options())!;
        var colors = PdfRadialSpreadColors.Create(stops, interpolation, null, false, transform);
        var field = new OfficeRadialGradient(0, 0, 0, 0, 0, 1, stops).WithSpreadMode(spread);
        string source = PdfRadialSpreadFunction.Build(0, 0, 0, 0, 0, 1, stops, spread,
            interpolation, null, false, false, colors: colors);
        Assert.True(PdfCalculatorProgram.TryParse(Encoding.ASCII.GetBytes(source), out var program));
        for (int x = -12; x <= 12; x++) {
            double px = x / 4D, py = .23;
            double ratio = field.SampleRatio(px, py);
            OfficeColor rgb = ratio <= .4
                ? OfficeGradientColors.Interpolate(stops[0].Color, stops[1].Color, ratio / .4, interpolation)
                : OfficeGradientColors.Interpolate(stops[1].Color, stops[2].Color, (ratio - .4) / .6, interpolation);
            var expected = new double[4]; transform.Convert(rgb, expected);
            var actual = program.Evaluate(new[] { px, py }, 4);
            Assert.NotNull(actual);
            for (int channel = 0; channel < 4; channel++)
                Assert.InRange(Math.Abs(actual![channel] - expected[channel]), 0, .01);
        }
    }

    [Fact]
    public void LargeAdaptiveColorFieldPreservesValuesAcrossLookupRanges() {
        var stops = Enumerable.Range(0, 300).Select(i => new OfficeGradientStop(i / 299D,
            OfficeColor.FromRgb((byte)(i * 255 / 299), 80, 170))).ToArray();
        var transform = PdfPrintColorTransform.Create(Options())!;
        var colors = PdfRadialSpreadColors.Create(stops, OfficeGradientColorInterpolation.Srgb, null, false, transform);
        Assert.InRange(colors.Samples.Count, 1025, 4096);
        for (int channel = 0; channel < 4; channel++) {
            string source = PdfRadialSpreadFunction.Build(0, 0, 0, 0, 0, 1, stops,
                OfficeGradientSpreadMode.Reflect, OfficeGradientColorInterpolation.Srgb,
                null, false, false, channel, colors);
            Assert.True(PdfCalculatorProgram.TryParse(Encoding.ASCII.GetBytes(source), out var program));
            foreach (double ratio in new[] { 0D, .25, .5, .75, .85, .9, 1D }) {
                int first = Math.Min((int)(ratio * 299), 298);
                var rgb = OfficeGradientColors.Interpolate(stops[first].Color, stops[first + 1].Color,
                    ratio * 299 - first, OfficeGradientColorInterpolation.Srgb);
                var expected = new double[4]; transform.Convert(rgb, expected);
                var actual = program.Evaluate(new[] { ratio, 0D }, 1);
                Assert.NotNull(actual);
                Assert.InRange(Math.Abs(actual![0] - expected[channel]), 0, .01);
            }
        }
    }

    [Fact]
    public void MaximumAuthoredStopsExportCmykWithIndependentScalarAlpha() {
        var stops = Enumerable.Range(0, 1024).Select(i => new OfficeGradientStop(i / 1023D,
            OfficeColor.FromRgba(230, 40, 80, (byte)(64 + i % 128)))).ToArray();
        var options = Options();
        var colors = PdfRadialSpreadColors.Create(stops, OfficeGradientColorInterpolation.Srgb,
            null, false, PdfPrintColorTransform.Create(options));
        Assert.Equal(4093, colors.Samples.Count);
        var shape = OfficeShape.Rectangle(60, 40); shape.StrokeWidth = 0;
        shape.FillRadialGradient = new OfficeRadialGradient(.25, .5, 0, .25, .5, .08, stops)
            .WithSpreadMode(OfficeGradientSpreadMode.Reflect);
        var pdf = PdfDocument.Create(options).Compose(c => c.Page(p => p.Content(content => content.Shape(shape)))).ToBytes();
        var document = PdfReadDocument.Open(pdf);
        var shadings = document.Objects.Values.Select(item => item.Value).OfType<PdfDictionary>()
            .Where(d => d.Items.TryGetValue("ShadingType", out var type) && type is PdfNumber { Value: 1 }).ToArray();
        var color = Assert.Single(shadings, d => d.Items["ColorSpace"] is PdfName { Name: "DeviceCMYK" });
        Assert.Equal(4, Assert.IsType<PdfArray>(color.Items["Function"]).Items.Count);
        Assert.Single(shadings, d => d.Items["ColorSpace"] is PdfName { Name: "DeviceGray" });
        var image = document.Pages[0].ExportImage(OfficeImageExportFormat.Png);
        Assert.DoesNotContain(image.Diagnostics, d => d.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.True(OfficeRasterImageDecoder.TryDecode(image.Bytes, out var raster));
        Assert.NotEqual(OfficeColor.White, raster!.GetPixel(29, 18));
    }

    private static PdfOptions Options() => new PdfOptions { PageWidth = 60, PageHeight = 40,
        MarginTop = 0, MarginBottom = 0, MarginLeft = 0, MarginRight = 0 }.ConfigurePdfXGroundwork(
            PdfComplianceProfile.PdfX4, IccMabTestProfiles.CreateCmykLab8Bidirectional(),
            "Managed test CMYK", PdfTrappingStatus.False);
}
