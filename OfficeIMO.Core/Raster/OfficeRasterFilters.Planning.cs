using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterFilters {
    private const long FixedFilteringWorkingBytes = 64L * 1024L;

    /// <summary>Estimates peak Gaussian blur working bytes in addition to the source and final RGBA buffers.</summary>
    /// <remarks>Includes the intermediate float buffer, kernel and fixed allocation allowance. Dimensions, sigma,
    /// single-image memory and operation limits are validated without allocating image buffers. For a sequence,
    /// reserve the largest frame estimate alongside all retained source and result buffers.</remarks>
    public static long EstimateGaussianBlurAdditionalWorkingBytes(int width, int height, double sigma = 3D) =>
        CalculateGaussianAdditionalWorkingBytes(width, height, sigma, false, out _);

    /// <summary>Estimates peak box blur working bytes in addition to the source and final RGBA buffers.</summary>
    /// <remarks>Includes the intermediate float buffer, kernel and fixed allocation allowance. Dimensions, radius,
    /// single-image memory and operation limits are validated without allocating image buffers. For a sequence,
    /// reserve the largest frame estimate alongside all retained source and result buffers.</remarks>
    public static long EstimateBoxBlurAdditionalWorkingBytes(int width, int height, int radius = 7) =>
        CalculateBoxBlurAdditionalWorkingBytes(width, height, radius);

    /// <summary>Estimates peak Gaussian sharpen working bytes in addition to the source and final RGBA buffers.</summary>
    /// <remarks>Reserves the larger of the blur intermediate and retained blurred image, plus the kernel and fixed
    /// allocation allowance. Dimensions, sigma, single-image memory and operation limits are validated without
    /// allocating image buffers. For a sequence, reserve the largest frame estimate alongside all retained buffers.</remarks>
    public static long EstimateGaussianSharpenAdditionalWorkingBytes(int width, int height, double sigma = 3D) =>
        CalculateGaussianAdditionalWorkingBytes(width, height, sigma, true, out _);

    /// <summary>Estimates peak adaptive threshold working bytes in addition to the source and final RGBA buffers.</summary>
    /// <remarks>Includes both padded integral tables and a fixed allocation allowance. Dimensions, radius, contrast,
    /// single-image memory and operation limits are validated without allocating image buffers. For a sequence,
    /// reserve the largest frame estimate alongside all retained source and result buffers.</remarks>
    public static long EstimateAdaptiveThresholdAdditionalWorkingBytes(int width, int height, int radius = 15, double contrast = .15D) =>
        CalculateAdaptiveThresholdAdditionalWorkingBytes(width, height, radius, contrast, out _);

    private static long CalculateGaussianAdditionalWorkingBytes(int width, int height, double sigma, bool sharpen, out int radius) {
        ValidateAmount(sigma, nameof(sigma), .01D, 42D);
        radius = (int)Math.Ceiling(sigma * 3D);
        return CalculateSmoothingAdditionalWorkingBytes(width, height, radius * 2 + 1, sharpen);
    }

    private static long CalculateBoxBlurAdditionalWorkingBytes(int width, int height, int radius) {
        ValidateRadius(radius, 128, nameof(radius));
        return CalculateSmoothingAdditionalWorkingBytes(width, height, radius * 2 + 1, false);
    }

    private static long CalculateSmoothingAdditionalWorkingBytes(int width, int height, int kernelLength, bool sharpen) {
        long pixels = OfficeRasterGuards.EnsureOutputPixels(width, height, "Raster filtering dimensions exceed the managed image limit.");
        long intermediateBytes = checked(pixels * 4L * sizeof(float));
        long retainedBlurBytes = sharpen ? checked(pixels * 4L) : 0L;
        long additionalBytes = checked(Math.Max(intermediateBytes, retainedBlurBytes) + kernelLength * (long)sizeof(double) + FixedFilteringWorkingBytes);
        ValidateFilterLimits(pixels, additionalBytes, kernelLength * 2L + (sharpen ? 1L : 0L));
        return additionalBytes;
    }

    private static long CalculateAdaptiveThresholdAdditionalWorkingBytes(int width, int height, int radius, double contrast, out int entries) {
        ValidateRadius(radius, 4096, nameof(radius));
        ValidateAmount(contrast, nameof(contrast), maximum: 1D);
        long pixels = OfficeRasterGuards.EnsureOutputPixels(width, height, "Raster filtering dimensions exceed the managed image limit.");
        long tableEntries = checked(((long)width + 1L) * ((long)height + 1L));
        long additionalBytes = checked(tableEntries * sizeof(double) * 2L + FixedFilteringWorkingBytes);
        ValidateFilterLimits(pixels, additionalBytes, 2L);
        entries = checked((int)tableEntries);
        return additionalBytes;
    }

    private static void ValidateFilterLimits(long pixels, long additionalWorkingBytes, long operationsPerPixel) {
        if (additionalWorkingBytes < 0L || additionalWorkingBytes > OfficeRasterGuards.MaximumDecodedBytes - pixels * 8L) {
            throw new ArgumentException("Raster filtering working set exceeds the managed image limit.");
        }
        if (operationsPerPixel <= 0L || pixels > MaximumOperations / operationsPerPixel) {
            throw new ArgumentException("Raster filtering exceeds the bounded operation count.");
        }
    }
}
