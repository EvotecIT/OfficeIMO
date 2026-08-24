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

    [Fact]
    public void LinearLightAreaDownsamplingAvoidsEncodedSrgbDarkening() {
        var source = new OfficeRasterImage(2, 1);
        source.SetPixel(0, 0, OfficeColor.Black);
        source.SetPixel(1, 0, OfficeColor.White);

        OfficeColor encoded = OfficeRasterResampler.Resize(
            source, 1, 1, OfficeRasterResamplingMode.Area,
            OfficeRasterResamplingColorSpace.EncodedSrgb).GetPixel(0, 0);
        OfficeColor linear = OfficeRasterResampler.Resize(
            source, 1, 1, OfficeRasterResamplingMode.Area,
            OfficeRasterResamplingColorSpace.LinearLight).GetPixel(0, 0);

        Assert.InRange(encoded.R, 127, 128);
        Assert.InRange(linear.R, 187, 188);
        Assert.Equal(linear.R, linear.G);
        Assert.Equal(linear.R, linear.B);
        Assert.Equal(255, linear.A);
    }

    [Fact]
    public void LinearLightLanczosStillUsesPremultipliedAlpha() {
        var source = new OfficeRasterImage(5, 1, OfficeColor.FromRgba(0, 0, 255, 0));
        source.SetPixel(2, 0, OfficeColor.FromRgba(128, 64, 32, 180));

        OfficeRasterImage result = OfficeRasterResampler.Resize(
            source, 13, 1, OfficeRasterResamplingMode.Lanczos3,
            OfficeRasterResamplingColorSpace.LinearLight);

        for (int x = 0; x < result.Width; x++) {
            OfficeColor pixel = result.GetPixel(x, 0);
            if (pixel.A == 0) continue;
            Assert.InRange(pixel.R, 127, 129);
            Assert.InRange(pixel.G, 63, 65);
            Assert.InRange(pixel.B, 31, 33);
        }
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

    [Theory]
    [InlineData(OfficeRasterResamplingMode.Area)]
    [InlineData(OfficeRasterResamplingMode.Lanczos3)]
    public void HighQualityResizeRejectsOversizedAxisMetadataBeforeAllocation(
        OfficeRasterResamplingMode mode) {
        var source = new OfficeRasterImage(1, 1, OfficeColor.Red);

        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            OfficeRasterResampler.Resize(source, 50_000_000, 1, mode));

        Assert.Contains("scratch space", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(OfficeRasterResamplingMode.Area)]
    [InlineData(OfficeRasterResamplingMode.Lanczos3)]
    public void HighQualityResizeAccountsForTheCompleteWorkingSetBeforeAllocation(
        OfficeRasterResamplingMode mode) {
        bool accepted = OfficeRasterResampler.TryGetHighQualityWorkingSetBytes(
            sourceWidth: 4000,
            sourceHeight: 10000,
            width: 800,
            height: 62500,
            mode,
            out long workingSetBytes);

        Assert.False(accepted);
        Assert.True(workingSetBytes > OfficeRasterGuards.MaximumDecodedBytes);
    }

    [Fact]
    public void HighQualityScratchAllocationMatchesTheBudgetedLength() {
        float[] scratch = OfficeRasterResampler.AllocateExactHighQualityScratch(1001);

        Assert.Equal(1001, scratch.Length);
    }

    [Theory]
    [InlineData(OfficeRasterResamplingMode.Area)]
    [InlineData(OfficeRasterResamplingMode.Lanczos3)]
    public void HighQualityWorkingSetAcceptsAnExactScratchNearTheLimit(
        OfficeRasterResamplingMode mode) {
        bool accepted = OfficeRasterResampler.TryGetHighQualityWorkingSetBytes(
            sourceWidth: 1000,
            sourceHeight: 1000,
            width: 6000,
            height: 6000,
            mode,
            out long workingSetBytes);

        Assert.True(accepted);
        Assert.InRange(
            workingSetBytes,
            230L * 1024L * 1024L,
            OfficeRasterGuards.MaximumDecodedBytes);
    }

    [Theory]
    [InlineData(OfficeRasterResamplingMode.Area)]
    [InlineData(OfficeRasterResamplingMode.Lanczos3)]
    public void HighQualityWorkingSetIncludesCallerRetainedBuffers(
        OfficeRasterResamplingMode mode) {
        Assert.True(OfficeRasterResampler.TryGetHighQualityWorkingSetBytes(
            sourceWidth: 1000,
            sourceHeight: 1000,
            width: 6000,
            height: 6000,
            mode,
            out long standaloneBytes));
        Assert.False(OfficeRasterResampler.TryGetHighQualityWorkingSetBytes(
            sourceWidth: 1000,
            sourceHeight: 1000,
            width: 6000,
            height: 6000,
            mode,
            retainedManagedBytes: 32L * 1024L * 1024L,
            out long optimizerBytes));
        Assert.Equal(standaloneBytes + 32L * 1024L * 1024L, optimizerBytes);
    }

    [Theory]
    [InlineData(OfficeRasterResamplingMode.NearestNeighbor)]
    [InlineData(OfficeRasterResamplingMode.Bilinear)]
    public void SimpleResizeRejectsCallerRetainedBuffersBeyondTheManagedLimit(
        OfficeRasterResamplingMode mode) {
        var source = new OfficeRasterImage(1, 1, OfficeColor.Red);

        Assert.Throws<ArgumentException>(() => OfficeRasterResampler.Resize(
            source,
            4096,
            4096,
            mode,
            OfficeRasterResamplingColorSpace.EncodedSrgb,
            retainedManagedBytes: 200L * 1024L * 1024L));
    }

    private static void AssertColorNear(OfficeColor expected, OfficeColor actual, int tolerance) {
        Assert.InRange(actual.R, Math.Max(0, expected.R - tolerance), Math.Min(255, expected.R + tolerance));
        Assert.InRange(actual.G, Math.Max(0, expected.G - tolerance), Math.Min(255, expected.G + tolerance));
        Assert.InRange(actual.B, Math.Max(0, expected.B - tolerance), Math.Min(255, expected.B + tolerance));
        Assert.InRange(actual.A, Math.Max(0, expected.A - tolerance), Math.Min(255, expected.A + tolerance));
    }
}
