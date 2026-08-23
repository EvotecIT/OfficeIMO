using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterResamplingQualityTests {
    [Fact]
    public void AreaDownsamplingComputesIndependentPixelCoverageAverage() {
        var source = new OfficeRasterImage(4, 4);
        for (int y = 0; y < source.Height; y++) {
            for (int x = 0; x < source.Width; x++) {
                source.SetPixel(x, y, OfficeColor.FromRgba(
                    (byte)(x * 40),
                    (byte)(y * 50),
                    (byte)((x + y) * 20),
                    255));
            }
        }

        OfficeColor result = OfficeRasterResampler.Resize(
            source,
            1,
            1,
            OfficeRasterResamplingMode.Area).GetPixel(0, 0);

        Assert.Equal(OfficeColor.FromRgba(60, 75, 60, 255), result);
    }

    [Fact]
    public void AreaDownsamplingUsesPremultipliedAlpha() {
        var source = new OfficeRasterImage(4, 1);
        source.SetPixel(0, 0, OfficeColor.FromRgba(0, 0, 255, 0));
        source.SetPixel(1, 0, OfficeColor.FromRgba(0, 0, 255, 0));
        source.SetPixel(2, 0, OfficeColor.Red);
        source.SetPixel(3, 0, OfficeColor.Red);

        OfficeColor result = OfficeRasterResampler.Resize(
            source,
            1,
            1,
            OfficeRasterResamplingMode.Area).GetPixel(0, 0);

        Assert.InRange(result.R, 254, 255);
        Assert.InRange(result.G, 0, 1);
        Assert.InRange(result.B, 0, 1);
        Assert.InRange(result.A, 127, 128);
    }

    [Fact]
    public void AreaDownsamplingAntialiasesHighFrequencyLineArt() {
        var source = new OfficeRasterImage(6, 6);
        for (int y = 0; y < source.Height; y++) {
            for (int x = 0; x < source.Width; x++) {
                source.SetPixel(x, y, ((x + y) & 1) == 0 ? OfficeColor.Black : OfficeColor.White);
            }
        }

        OfficeColor result = OfficeRasterResampler.Resize(
            source,
            1,
            1,
            OfficeRasterResamplingMode.Area).GetPixel(0, 0);

        Assert.InRange(result.R, 127, 128);
        Assert.Equal(result.R, result.G);
        Assert.Equal(result.R, result.B);
        Assert.Equal(255, result.A);
    }

    [Theory]
    [InlineData(2, 2)]
    [InlineData(13, 9)]
    public void Lanczos3PreservesConstantPremultipliedColor(int width, int height) {
        OfficeColor expected = OfficeColor.FromRgba(20, 80, 200, 137);
        var source = new OfficeRasterImage(5, 4, expected);

        OfficeRasterImage result = OfficeRasterResampler.Resize(
            source,
            width,
            height,
            OfficeRasterResamplingMode.Lanczos3);

        for (int y = 0; y < result.Height; y++) {
            for (int x = 0; x < result.Width; x++) {
                AssertColorNear(expected, result.GetPixel(x, y), 1);
            }
        }
    }

    [Fact]
    public void Lanczos3DoesNotLeakHiddenTransparentColor() {
        var source = new OfficeRasterImage(5, 1, OfficeColor.FromRgba(0, 0, 255, 0));
        source.SetPixel(2, 0, OfficeColor.Red);

        OfficeRasterImage result = OfficeRasterResampler.Resize(
            source,
            13,
            1,
            OfficeRasterResamplingMode.Lanczos3);

        for (int x = 0; x < result.Width; x++) {
            OfficeColor pixel = result.GetPixel(x, 0);
            if (pixel.A == 0) continue;
            Assert.InRange(pixel.R, 254, 255);
            Assert.InRange(pixel.G, 0, 1);
            Assert.InRange(pixel.B, 0, 1);
        }
    }

    [Fact]
    public void AreaUsesLinearSamplingOnlyOnEnlargedAxis() {
        var source = new OfficeRasterImage(2, 2);
        source.SetPixel(0, 0, OfficeColor.Black);
        source.SetPixel(1, 0, OfficeColor.White);
        source.SetPixel(0, 1, OfficeColor.Black);
        source.SetPixel(1, 1, OfficeColor.White);

        OfficeRasterImage result = OfficeRasterResampler.Resize(
            source,
            4,
            1,
            OfficeRasterResamplingMode.Area);

        Assert.Equal(OfficeColor.Black, result.GetPixel(0, 0));
        Assert.Equal(OfficeColor.White, result.GetPixel(3, 0));
        Assert.InRange(result.GetPixel(1, 0).R, 63, 64);
        Assert.InRange(result.GetPixel(2, 0).R, 191, 192);
    }

    [Fact]
    public void ResizeRejectsUnknownResamplingMode() {
        var source = new OfficeRasterImage(2, 2, OfficeColor.Red);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            OfficeRasterResampler.Resize(source, 1, 1, (OfficeRasterResamplingMode)999));
    }

    private static void AssertColorNear(OfficeColor expected, OfficeColor actual, int tolerance) {
        Assert.InRange(actual.R, Math.Max(0, expected.R - tolerance), Math.Min(255, expected.R + tolerance));
        Assert.InRange(actual.G, Math.Max(0, expected.G - tolerance), Math.Min(255, expected.G + tolerance));
        Assert.InRange(actual.B, Math.Max(0, expected.B - tolerance), Math.Min(255, expected.B + tolerance));
        Assert.InRange(actual.A, Math.Max(0, expected.A - tolerance), Math.Min(255, expected.A + tolerance));
    }
}
