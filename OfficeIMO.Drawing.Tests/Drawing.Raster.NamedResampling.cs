using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterNamedResamplingTests {
    [Fact]
    public void NamedKernelsPreserveConstantColorAndAlphaThroughDownsampling() {
        var color = OfficeColor.FromRgba(120, 80, 40, 79);
        var source = new OfficeRasterImage(9, 7, color);
        foreach (OfficeRasterResamplingMode mode in Enum.GetValues(typeof(OfficeRasterResamplingMode))) {
            OfficeRasterImage result = OfficeRasterResampler.Resize(source, 3, 2, mode);
            Assert.Equal(3, result.Width); Assert.Equal(2, result.Height);
            for (int y = 0; y < 2; y++) for (int x = 0; x < 3; x++) Assert.Equal(color, result.GetPixel(x, y));
        }
        Assert.Equal(color, source.GetPixel(0, 0));
    }

    [Fact]
    public void BicubicMatchesIndependentCatmullRomStepInterpolation() {
        var source = new OfficeRasterImage(4, 1, OfficeColor.Black);
        source.SetPixel(2, 0, OfficeColor.White); source.SetPixel(3, 0, OfficeColor.White);
        OfficeRasterImage cubic = OfficeRasterResampler.Resize(source, 8, 1, OfficeRasterResamplingMode.Bicubic);
        // At source coordinate 1.25, the Catmull-Rom step is .203125 * 255 = 51.796875.
        Assert.Equal(OfficeColor.FromRgb(52, 52, 52), cubic.GetPixel(3, 0));
        Assert.Equal(OfficeColor.FromRgb(64, 64, 64), OfficeRasterResampler.Resize(source, 8, 1, OfficeRasterResamplingMode.Triangle).GetPixel(3, 0));
        Assert.Equal(cubic.GetPixels(), OfficeRasterResampler.Resize(source, 8, 1, OfficeRasterResamplingMode.CatmullRom).GetPixels());
    }

    [Theory]
    [InlineData(OfficeRasterResamplingMode.Box, 255)]
    [InlineData(OfficeRasterResamplingMode.Triangle, 191)]
    [InlineData(OfficeRasterResamplingMode.Hermite, 215)]
    [InlineData(OfficeRasterResamplingMode.Bicubic, 221)]
    [InlineData(OfficeRasterResamplingMode.CatmullRom, 221)]
    [InlineData(OfficeRasterResamplingMode.MitchellNetravali, 199)]
    [InlineData(OfficeRasterResamplingMode.Spline, 156)]
    public void NamedPolynomialKernelsHaveDistinctIndependentQuarterPixelValues(OfficeRasterResamplingMode mode, byte expected) {
        var source = new OfficeRasterImage(4, 1, OfficeColor.Black);
        source.SetPixel(1, 0, OfficeColor.White);
        // A unit impulse evaluated one quarter-pixel from its center, multiplied by 255.
        OfficeRasterImage result = OfficeRasterResampler.Resize(source, 8, 1, mode);
        Assert.Equal(OfficeColor.FromRgb(expected, expected, expected), result.GetPixel(3, 0));
    }
}
