using System;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgGradientAlphaInterpolationTests {
    [Theory]
    [InlineData(false, false, 1D)]
    [InlineData(true, false, 1D)]
    [InlineData(false, true, .5D)]
    [InlineData(true, true, .5D)]
    public void ImportedPaintKeepsSeparateColorAndAlphaAcrossContoursTransformsAndOpacity(bool radial, bool transformed, double opacity) {
        string gradient = radial ? "radialGradient" : "linearGradient";
        string attributes = radial ? "cx='.5' cy='.5' r='.5'" : "x1='0' y1='.5' x2='1' y2='.5'";
        var transform = transformed ? new OfficeTransform(-1, 0, .5, 1, 120, 0) : OfficeTransform.Identity;
        string shapeTransform = transformed ? " transform='matrix(-1 0 .5 1 120 0)'" : "";
        foreach (string geometry in new[] { "<rect width='100' height='100'", "<polygon points='0,0 100,0 100,100 0,100'", "<path d='M0,0H100V100H0Z M40,40V60H60V40Z'" }) {
            string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='180' height='120'><defs><" + gradient + " id='g' " + attributes + "><stop offset='0' stop-color='red' stop-opacity='0.2509803921568627'/><stop offset='1' stop-color='blue' stop-opacity='0.5019607843137255'/></" + gradient + "></defs>" + geometry + shapeTransform + " fill='url(#g)' fill-opacity='" + opacity.ToString(System.Globalization.CultureInfo.InvariantCulture) + "'/></svg>";
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
            Assert.Equal(0, unsupported);
            var raster = OfficeDrawingRasterRenderer.Render(drawing!);
            var inverse = transform.Invert();
            foreach (var sample in new[] { (25, 35), (75, 65) }) {
                var target = transform.TransformPoint(new OfficePoint(sample.Item1, sample.Item2));
                int x = (int)target.X, y = (int)target.Y;
                var local = inverse.TransformPoint(new OfficePoint(x + .5D, y + .5D));
                double ratio = radial ? Math.Min(1D, Math.Sqrt(Math.Pow(local.X / 100D - .5D, 2) + Math.Pow(local.Y / 100D - .5D, 2)) / .5D) : local.X / 100D;
                var pixel = raster.GetPixel(x, y);
                Assert.InRange(Math.Abs(pixel.R - 255D * (1D - ratio)), 0D, 1D);
                Assert.Equal(0, pixel.G);
                Assert.InRange(Math.Abs(pixel.B - 255D * ratio), 0D, 1D);
                Assert.InRange(Math.Abs(pixel.A - (64D + 64D * ratio) * opacity), 0D, 1D);
            }
            if (geometry.StartsWith("<path", StringComparison.Ordinal)) {
                var hole = transform.TransformPoint(new OfficePoint(50, 50));
                Assert.Equal(0, raster.GetPixel((int)hole.X, (int)hole.Y).A);
            }
        }
    }
}
