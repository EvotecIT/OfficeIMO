using System;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterFiltersTests {
    [Fact]
    public void PublicRgbaConstructionAndCloneOwnIndependentBuffers() {
        byte[] samples = { 10, 20, 30, 40 };
        OfficeRasterImage original = OfficeRasterImage.FromRgba32(1, 1, samples);
        OfficeRasterImage clone = original.Clone();
        samples[0] = 255;
        clone.SetPixel(0, 0, OfficeColor.White);
        Assert.Equal(OfficeColor.FromRgba(10, 20, 30, 40), original.GetPixel(0, 0));
        Assert.Equal(OfficeColor.White, clone.GetPixel(0, 0));
        Assert.Throws<ArgumentException>(() => OfficeRasterImage.FromRgba32(1, 1, new byte[3]));
        Assert.Throws<ArgumentException>(() => OfficeRasterImage.FromRgba32(1, 1, new byte[5]));
        Assert.Throws<ArgumentNullException>(() => OfficeRasterImage.FromRgba32(1, 1, null!));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterImage.FromRgba32(0, 1, Array.Empty<byte>()));
    }

    [Theory]
    [InlineData(OfficeRasterGrayscaleMode.Bt709, 54)]
    [InlineData(OfficeRasterGrayscaleMode.Bt601, 76)]
    public void GrayscaleSelectsDocumentedWeightsAndRetainsAlpha(OfficeRasterGrayscaleMode mode, byte expected) {
        var source = new OfficeRasterImage(1, 1, OfficeColor.FromRgba(255, 0, 0, 37));
        OfficeRasterImage result = OfficeRasterFilters.Grayscale(source, mode);
        Assert.Equal(OfficeColor.FromRgba(expected, expected, expected, 37), result.GetPixel(0, 0));
        Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 37), source.GetPixel(0, 0));
    }

    [Fact]
    public void ColorOperationsRespectUnitsAndPreserveSource() {
        var source = new OfficeRasterImage(1, 1, OfficeColor.FromRgba(200, 80, 20, 100));
        Assert.Equal(OfficeColor.FromRgba(100, 40, 10, 100), OfficeRasterFilters.Brightness(source, .5D).GetPixel(0, 0));
        Assert.Equal(OfficeColor.FromRgba(200, 80, 20, 25), OfficeRasterFilters.Opacity(source, .25D).GetPixel(0, 0));
        Assert.Equal(OfficeColor.FromRgba(55, 175, 235, 100), OfficeRasterFilters.Invert(source).GetPixel(0, 0));
        Assert.Equal(source.GetPixels(), OfficeRasterFilters.Hue(source, 360D).GetPixels());
        Assert.Equal(source.GetPixels(), OfficeRasterFilters.Lightness(source).GetPixels());
        Assert.Equal(source.GetPixels(), OfficeRasterFilters.Sepia(source, 0D).GetPixels());
        Assert.Equal(OfficeColor.FromRgba(200, 80, 20, 100), source.GetPixel(0, 0));
        var red = new OfficeRasterImage(1, 1, OfficeColor.Red);
        Assert.Equal(OfficeColor.Lime, OfficeRasterFilters.Hue(red, 120D).GetPixel(0, 0));
    }

    [Fact]
    public void ColorMatrixOwnsCoefficientsAndCanTransformAlphaExplicitly() {
        double[] coefficients = { 0, 0, 1, 0, 0, 0, 1, 0, 0, .1D, 1, 0, 0, 0, 0, 0, 0, 0, .5D, 0 };
        var matrix = new OfficeRasterColorMatrix(coefficients);
        coefficients[0] = 99;
        var source = new OfficeRasterImage(1, 1, OfficeColor.FromRgba(100, 50, 20, 80));
        Assert.Equal(OfficeColor.FromRgba(20, 76, 100, 40), OfficeRasterFilters.ColorMatrix(source, matrix).GetPixel(0, 0));
        Assert.Equal(source.GetPixels(), OfficeRasterFilters.ColorMatrix(source, OfficeRasterColorMatrix.Identity).GetPixels());
        Assert.Throws<ArgumentException>(() => new OfficeRasterColorMatrix(new double[19]));
    }

    [Fact]
    public void BlurDoesNotBleedInvisibleBlueIntoVisibleRed() {
        OfficeRasterImage source = OfficeRasterImage.FromRgba32(3, 1, new byte[] { 0, 0, 255, 0, 255, 0, 0, 255, 0, 0, 255, 0 });
        foreach (OfficeRasterImage blurred in new[] {
            OfficeRasterFilters.GaussianBlur(source, 1D), OfficeRasterFilters.BoxBlur(source, 1), OfficeRasterFilters.BokehBlur(source, 1, 1D)
        }) {
            OfficeColor center = blurred.GetPixel(1, 0);
            Assert.Equal(255, center.R); Assert.Equal(0, center.B);
            Assert.InRange(center.A, 1, 254);
        }
        Assert.Equal(OfficeColor.FromRgba(0, 0, 255, 0), source.GetPixel(0, 0));
        // Gamma adjustment is performed at filtering precision, so a constant dark
        // source does not lose its low bits in an intermediate eight-bit image.
        var dark = new OfficeRasterImage(3, 3, OfficeColor.FromRgba(10, 20, 30, 83));
        Assert.Equal(dark.GetPixels(), OfficeRasterFilters.BokehBlur(dark).GetPixels());
    }

    [Fact]
    public void PixelateHandlesPartialBlocksAndPremultipliedAlpha() {
        OfficeRasterImage source = OfficeRasterImage.FromRgba32(3, 1, new byte[] { 255, 0, 0, 255, 0, 0, 255, 0, 0, 255, 0, 255 });
        OfficeRasterImage result = OfficeRasterFilters.Pixelate(source, 2);
        Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 128), result.GetPixel(0, 0));
        Assert.Equal(result.GetPixel(0, 0), result.GetPixel(1, 0));
        Assert.Equal(OfficeColor.Lime, result.GetPixel(2, 0));
    }

    [Fact]
    public void ThresholdDitherAndHistogramHaveObservableTonalContracts() {
        var uniform = new OfficeRasterImage(4, 4, OfficeColor.FromRgba(128, 128, 128, 37));
        OfficeRasterImage dithered = OfficeRasterFilters.Dither(uniform);
        int white = 0;
        for (int y = 0; y < 4; y++) for (int x = 0; x < 4; x++) {
            OfficeColor pixel = dithered.GetPixel(x, y);
            if (pixel.R == 255) white++;
            Assert.Equal(pixel.R, pixel.G); Assert.Equal(pixel.R, pixel.B); Assert.Equal(37, pixel.A);
        }
        Assert.Equal(8, white);
        Assert.Equal(OfficeColor.FromRgba(255, 255, 255, 37), OfficeRasterFilters.Threshold(uniform).GetPixel(0, 0));
        Assert.Equal(uniform.GetPixels(), OfficeRasterFilters.HistogramEqualization(uniform).GetPixels());
        var ramp = OfficeRasterImage.FromRgba32(2, 1, new byte[] { 80, 80, 80, 255, 160, 160, 160, 255 });
        OfficeRasterImage equalized = OfficeRasterFilters.HistogramEqualization(ramp);
        Assert.Equal(OfficeColor.Black, equalized.GetPixel(0, 0)); Assert.Equal(OfficeColor.White, equalized.GetPixel(1, 0));
    }

    [Fact]
    public void AdaptiveThresholdAndOilPaintingIgnoreInvisibleNeighborhoodColors() {
        var source = new OfficeRasterImage(3, 3, OfficeColor.FromRgba(0, 0, 255, 0));
        source.SetPixel(1, 1, OfficeColor.FromRgba(100, 0, 0, 77));
        OfficeRasterImage oil = OfficeRasterFilters.OilPaint(source, 10, 3);
        Assert.Equal(source.GetPixel(1, 1), oil.GetPixel(1, 1));
        OfficeRasterImage threshold = OfficeRasterFilters.AdaptiveThreshold(source, 1, .15D);
        Assert.Equal(OfficeColor.FromRgba(255, 255, 255, 77), threshold.GetPixel(1, 1));
    }

    [Fact]
    public void VignettePreservesCenterAndAlphaWhileDarkeningEdges() {
        var source = new OfficeRasterImage(7, 7, OfficeColor.FromRgba(200, 200, 200, 83));
        OfficeRasterImage result = OfficeRasterFilters.Vignette(source);
        Assert.Equal(source.GetPixel(3, 3), result.GetPixel(3, 3));
        Assert.True(result.GetPixel(0, 0).R < result.GetPixel(3, 3).R);
        Assert.Equal(83, result.GetPixel(0, 0).A);
    }

    [Fact]
    public void FilteringRejectsNonFiniteArgumentsAndObservesCancellation() {
        var source = new OfficeRasterImage(2, 2, OfficeColor.White);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.Brightness(source, double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.GaussianBlur(source, double.PositiveInfinity));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.Opacity(source, 1.1D));
        Assert.Throws<ArgumentException>(() => OfficeRasterFilters.Convolve(source, new double[9], 3, 3));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeRasterFilters.Invert(source, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => OfficeRasterFilters.BoxBlur(source, cancellationToken: cancellation.Token));
    }

    [Fact]
    public void NonNormalizedEdgeKernelRetainsVisibleAlphaAndFilterWorkIsBounded() {
        OfficeRasterImage ramp = OfficeRasterImage.FromRgba32(3, 1, new byte[] { 0, 0, 0, 255, 100, 100, 100, 255, 200, 200, 200, 255 });
        OfficeRasterImage edges = OfficeRasterFilters.Convolve(ramp, new double[] { -1, 0, 1 }, 3, 1, normalize: false);
        Assert.Equal(OfficeColor.FromRgb(200, 200, 200), edges.GetPixel(1, 0));
        // Reject before allocating filter output or performing billions of neighborhood visits.
        var bounded = new OfficeRasterImage(1001, 1001);
        Assert.Throws<ArgumentException>(() => OfficeRasterFilters.OilPaint(bounded, 10, 65));
    }
}
