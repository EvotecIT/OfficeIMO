using OfficeIMO.Drawing;
using System;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RasterFilterBudgetContracts {
    [Fact]
    public void SequencePlanningReservesPixelScaleScratchAlongsideRetainedFrames() {
        const int width = 7000, height = 1000, frameCount = 3;
        const long maximumWorkingBytes = 256L * 1024L * 1024L;
        long retainedSourceAndResults = (long)width * height * frameCount * 8L;
        Assert.True(retainedSourceAndResults < maximumWorkingBytes);

        long[] estimates = {
            OfficeRasterFilters.EstimateGaussianBlurAdditionalWorkingBytes(width, height, .01D),
            OfficeRasterFilters.EstimateBoxBlurAdditionalWorkingBytes(width, height, 1),
            OfficeRasterFilters.EstimateGaussianSharpenAdditionalWorkingBytes(width, height, .01D),
            OfficeRasterFilters.EstimateAdaptiveThresholdAdditionalWorkingBytes(width, height)
        };
        foreach (long estimate in estimates) {
            Assert.True(retainedSourceAndResults + estimate > maximumWorkingBytes);
        }
        long smoothingScratchAndKernel = (long)width * height * 4L * sizeof(float) + 3L * sizeof(double) + 65536L;
        for (int index = 0; index < 3; index++) {
            Assert.True(estimates[index] >= smoothingScratchAndKernel);
        }
        Assert.True(estimates[3] >= ((long)width + 1L) * (height + 1L) * 16L + 65536L);
    }

    [Fact]
    public void PlanningAndExecutionRejectTheSameSingleImageMemoryBudget() {
        var source = new OfficeRasterImage(4000, 3000, OfficeColor.Blue);
        AssertWorkingSetRejection(
            () => OfficeRasterFilters.EstimateGaussianBlurAdditionalWorkingBytes(source.Width, source.Height, .01D),
            () => OfficeRasterFilters.GaussianBlur(source, .01D));
        AssertWorkingSetRejection(
            () => OfficeRasterFilters.EstimateBoxBlurAdditionalWorkingBytes(source.Width, source.Height, 1),
            () => OfficeRasterFilters.BoxBlur(source, 1));
        AssertWorkingSetRejection(
            () => OfficeRasterFilters.EstimateGaussianSharpenAdditionalWorkingBytes(source.Width, source.Height, .01D),
            () => OfficeRasterFilters.GaussianSharpen(source, .01D));
        AssertWorkingSetRejection(
            () => OfficeRasterFilters.EstimateAdaptiveThresholdAdditionalWorkingBytes(source.Width, source.Height),
            () => OfficeRasterFilters.AdaptiveThreshold(source));
        Assert.Equal(OfficeColor.Blue, source.GetPixel(source.Width - 1, source.Height - 1));
    }

    [Fact]
    public void SmoothingPlanningAndExecutionEnforceTheBoundedWorkCount() {
        var source = new OfficeRasterImage(2000, 2000);
        AssertOperationRejection(
            () => OfficeRasterFilters.EstimateGaussianBlurAdditionalWorkingBytes(source.Width, source.Height, 42D),
            () => OfficeRasterFilters.GaussianBlur(source, 42D));
        AssertOperationRejection(
            () => OfficeRasterFilters.EstimateBoxBlurAdditionalWorkingBytes(source.Width, source.Height, 128),
            () => OfficeRasterFilters.BoxBlur(source, 128));
        AssertOperationRejection(
            () => OfficeRasterFilters.EstimateGaussianSharpenAdditionalWorkingBytes(source.Width, source.Height, 42D),
            () => OfficeRasterFilters.GaussianSharpen(source, 42D));
    }

    [Fact]
    public void PlanningValidatesDimensionsAndTheSamePublicSettingsAsExecution() {
        var source = new OfficeRasterImage(1, 1);
        Assert.Throws<ArgumentException>(() => OfficeRasterFilters.EstimateGaussianBlurAdditionalWorkingBytes(0, 1));
        Assert.Throws<ArgumentException>(() => OfficeRasterFilters.EstimateBoxBlurAdditionalWorkingBytes(1, 0));
        Assert.Throws<ArgumentException>(() => OfficeRasterFilters.EstimateGaussianSharpenAdditionalWorkingBytes(int.MaxValue, 2));
        Assert.Throws<ArgumentException>(() => OfficeRasterFilters.EstimateAdaptiveThresholdAdditionalWorkingBytes(1, -1));

        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.EstimateGaussianBlurAdditionalWorkingBytes(1, 1, double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.GaussianBlur(source, double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.EstimateGaussianSharpenAdditionalWorkingBytes(1, 1, 0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.GaussianSharpen(source, 0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.EstimateBoxBlurAdditionalWorkingBytes(1, 1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.BoxBlur(source, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.EstimateAdaptiveThresholdAdditionalWorkingBytes(1, 1, contrast: double.PositiveInfinity));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.AdaptiveThreshold(source, contrast: double.PositiveInfinity));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.EstimateAdaptiveThresholdAdditionalWorkingBytes(1, 1, radius: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterFilters.AdaptiveThreshold(source, radius: 0));
    }

    [Fact]
    public void PrivateBokehKernelAndPublicConvolutionRetainTheirOwnershipContracts() {
        OfficeColor color = OfficeColor.FromRgba(60, 90, 120, 128);
        var source = new OfficeRasterImage(1, 1, color);
        Assert.Equal(color, OfficeRasterFilters.BokehBlur(source, 32, 1D).GetPixel(0, 0));

        double[] kernel = { 1D, 1D, 1D };
        Assert.Equal(color, OfficeRasterFilters.Convolve(source, kernel, 3, 1).GetPixel(0, 0));
        Assert.Equal(new double[] { 1D, 1D, 1D }, kernel);
        Assert.Equal(color, source.GetPixel(0, 0));
    }

    private static void AssertWorkingSetRejection(Func<long> estimate, Func<OfficeRasterImage> execute) {
        Assert.Contains("working set", Assert.Throws<ArgumentException>(() => estimate()).Message);
        Assert.Contains("working set", Assert.Throws<ArgumentException>(() => execute()).Message);
    }

    private static void AssertOperationRejection(Func<long> estimate, Func<OfficeRasterImage> execute) {
        Assert.Contains("operation count", Assert.Throws<ArgumentException>(() => estimate()).Message);
        Assert.Contains("operation count", Assert.Throws<ArgumentException>(() => execute()).Message);
    }
}
