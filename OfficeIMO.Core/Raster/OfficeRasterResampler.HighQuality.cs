using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterResampler {
    private const double LanczosRadius = 3D;
    private const string ScratchLimitMessage = "High-quality raster resampling scratch space exceeds the managed image limit.";

    private static OfficeRasterImage ResizeSeparable(
        OfficeRasterImage source,
        int width,
        int height,
        OfficeRasterResamplingMode mode,
        OfficeRasterResamplingColorSpace colorSpace) {
        if (!TryMeasureSeparableWorkingSet(
                source.Width,
                source.Height,
                width,
                height,
                mode,
                out int intermediateLength,
                out int horizontalContributionCount,
                out int verticalContributionCount,
                out _)) {
            throw new ArgumentException(ScratchLimitMessage, nameof(source));
        }
        long horizontalFirstLength = (long)width * source.Height * 4L;
        long verticalFirstLength = (long)source.Width * height * 4L;
        AxisContributions horizontal = CreateContributions(
            source.Width, width, mode, horizontalContributionCount);
        AxisContributions vertical = CreateContributions(
            source.Height, height, mode, verticalContributionCount);
        var result = new OfficeRasterImage(width, height);
        float[] intermediate = AllocateExactHighQualityScratch(intermediateLength);
        if (horizontalFirstLength <= verticalFirstLength) {
            ResampleHorizontal(source.PixelBuffer, source.Width, source.Height, width, horizontal, intermediate, colorSpace);
            ResampleVertical(intermediate, width, height, vertical, result.PixelBuffer, colorSpace);
        } else {
            ResampleVertical(source.PixelBuffer, source.Width, source.Height, height, vertical, intermediate, colorSpace);
            ResampleHorizontal(intermediate, source.Width, width, height, horizontal, result.PixelBuffer, colorSpace);
        }
        return result;
    }

    internal static float[] AllocateExactHighQualityScratch(int length) {
        if (length <= 0) throw new ArgumentOutOfRangeException(nameof(length));
#if NET8_0_OR_GREATER
        return GC.AllocateUninitializedArray<float>(length);
#else
        return new float[length];
#endif
    }

    internal static bool TryGetHighQualityWorkingSetBytes(
        int sourceWidth,
        int sourceHeight,
        int width,
        int height,
        OfficeRasterResamplingMode mode,
        out long workingSetBytes) =>
        TryMeasureSeparableWorkingSet(
            sourceWidth,
            sourceHeight,
            width,
            height,
            mode,
            out _,
            out _,
            out _,
            out workingSetBytes);

    private static bool TryMeasureSeparableWorkingSet(
        int sourceWidth,
        int sourceHeight,
        int width,
        int height,
        OfficeRasterResamplingMode mode,
        out int intermediateLength,
        out int horizontalContributionCount,
        out int verticalContributionCount,
        out long workingSetBytes) {
        intermediateLength = 0;
        horizontalContributionCount = 0;
        verticalContributionCount = 0;
        workingSetBytes = 0L;
        if (sourceWidth <= 0 || sourceHeight <= 0 || width <= 0 || height <= 0) return false;
        try {
            long horizontalFirstLength = checked((long)width * sourceHeight * 4L);
            long verticalFirstLength = checked((long)sourceWidth * height * 4L);
            long requiredIntermediate = Math.Min(horizontalFirstLength, verticalFirstLength);
            if (requiredIntermediate <= 0L || requiredIntermediate > int.MaxValue ||
                !TryMeasureContributions(
                    sourceWidth, width, mode, out horizontalContributionCount, out long horizontalBytes) ||
                !TryMeasureContributions(
                    sourceHeight, height, mode, out verticalContributionCount, out long verticalBytes)) {
                return false;
            }

            long sourceBytes = checked((long)sourceWidth * sourceHeight * 4L);
            long destinationBytes = checked((long)width * height * 4L);
            workingSetBytes = checked(
                sourceBytes + destinationBytes + requiredIntermediate * sizeof(float) +
                horizontalBytes + verticalBytes + 64L * 1024L);
            if (workingSetBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
            intermediateLength = (int)requiredIntermediate;
            return true;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool TryMeasureContributions(
        int sourceLength,
        int destinationLength,
        OfficeRasterResamplingMode mode,
        out int contributionCount,
        out long storageBytes) {
        contributionCount = 0;
        storageBytes = 0L;
        long metadataBytes = (long)destinationLength * 3L * sizeof(int);
        if (sourceLength <= 0 || destinationLength <= 0 ||
            metadataBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
        double scale = sourceLength / (double)destinationLength;
        long total = 0L;
        for (int destination = 0; destination < destinationLength; destination++) {
            GetContributionRange(sourceLength, destination, scale, mode, out _, out int count);
            total += count;
            if (total > int.MaxValue) return false;
        }
        storageBytes = checked(metadataBytes + total * sizeof(double));
        if (storageBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
        contributionCount = (int)total;
        return total > 0L;
    }

    private static AxisContributions CreateContributions(
        int sourceLength,
        int destinationLength,
        OfficeRasterResamplingMode mode,
        int contributionCount) {
        var starts = new int[destinationLength];
        var counts = new int[destinationLength];
        var offsets = new int[destinationLength];
        double scale = sourceLength / (double)destinationLength;
        long total = 0L;
        for (int destination = 0; destination < destinationLength; destination++) {
            GetContributionRange(
                sourceLength,
                destination,
                scale,
                mode,
                out int start,
                out int count);
            starts[destination] = start;
            counts[destination] = count;
            offsets[destination] = (int)total;
            total += count;
        }
        if (total != contributionCount) {
            throw new ArgumentException(ScratchLimitMessage);
        }
        var weights = new double[contributionCount];
        for (int destination = 0; destination < destinationLength; destination++) {
            WriteContributionWeights(
                sourceLength,
                destination,
                scale,
                mode,
                starts[destination],
                counts[destination],
                weights,
                offsets[destination]);
        }
        return new AxisContributions(starts, counts, offsets, weights);
    }

    private static void GetContributionRange(
        int sourceLength,
        int destination,
        double scale,
        OfficeRasterResamplingMode mode,
        out int start,
        out int count) {
        if (mode == OfficeRasterResamplingMode.Area && scale > 1D) {
            double sourceStart = destination * scale;
            double sourceEnd = (destination + 1D) * scale;
            start = Clamp((int)Math.Floor(sourceStart), 0, sourceLength - 1);
            int end = Clamp((int)Math.Ceiling(sourceEnd) - 1, start, sourceLength - 1);
            count = end - start + 1;
            return;
        }

        double center = ((destination + 0.5D) * scale) - 0.5D;
        if (mode == OfficeRasterResamplingMode.Area) {
            double sample = Clamp(center, 0D, sourceLength - 1D);
            start = (int)Math.Floor(sample);
            count = start < sourceLength - 1 && sample > start ? 2 : 1;
            return;
        }

        double filterScale = Math.Max(1D, scale);
        double support = LanczosRadius * filterScale;
        start = Math.Max(0, (int)Math.Ceiling(center - support));
        int last = Math.Min(sourceLength - 1, (int)Math.Floor(center + support));
        if (last < start) {
            start = Clamp((int)Math.Round(center), 0, sourceLength - 1);
            count = 1;
        } else {
            count = last - start + 1;
        }
    }

    private static void WriteContributionWeights(
        int sourceLength,
        int destination,
        double scale,
        OfficeRasterResamplingMode mode,
        int start,
        int count,
        double[] weights,
        int offset) {
        if (mode == OfficeRasterResamplingMode.Area && scale > 1D) {
            double sourceStart = destination * scale;
            double sourceEnd = (destination + 1D) * scale;
            double sum = 0D;
            for (int index = 0; index < count; index++) {
                int source = start + index;
                double weight = Math.Max(0D, Math.Min(sourceEnd, source + 1D) - Math.Max(sourceStart, source));
                weights[offset + index] = weight;
                sum += weight;
            }
            Normalize(weights, offset, count, sum);
            return;
        }

        double center = ((destination + 0.5D) * scale) - 0.5D;
        if (mode == OfficeRasterResamplingMode.Area) {
            if (count == 1) {
                weights[offset] = 1D;
            } else {
                double fraction = Clamp(center, 0D, sourceLength - 1D) - start;
                weights[offset] = 1D - fraction;
                weights[offset + 1] = fraction;
            }
            return;
        }

        double filterScale = Math.Max(1D, scale);
        double total = 0D;
        for (int index = 0; index < count; index++) {
            double weight = Lanczos((center - (start + index)) / filterScale);
            weights[offset + index] = weight;
            total += weight;
        }
        if (total <= 1E-12D) {
            for (int index = 0; index < count; index++) weights[offset + index] = 0D;
            int nearest = Clamp((int)Math.Round(center), start, start + count - 1);
            weights[offset + nearest - start] = 1D;
        } else {
            Normalize(weights, offset, count, total);
        }
    }

    private static void Normalize(double[] weights, int offset, int count, double total) {
        if (total <= 0D) throw new InvalidOperationException("Raster resampling weights are empty.");
        for (int index = 0; index < count; index++) weights[offset + index] /= total;
    }

    private static double Lanczos(double value) {
        double absolute = Math.Abs(value);
        if (absolute < 1E-12D) return 1D;
        if (absolute >= LanczosRadius) return 0D;
        double piValue = Math.PI * value;
        return (Math.Sin(piValue) / piValue) *
            (Math.Sin(piValue / LanczosRadius) / (piValue / LanczosRadius));
    }

    private static void ResampleHorizontal(
        byte[] input,
        int sourceWidth,
        int sourceHeight,
        int destinationWidth,
        AxisContributions contributions,
        float[] output,
        OfficeRasterResamplingColorSpace colorSpace) {
        for (int y = 0; y < sourceHeight; y++) {
            for (int x = 0; x < destinationWidth; x++) {
                int target = ((y * destinationWidth) + x) * 4;
                AccumulateBytes(input, (y * sourceWidth) * 4, 4, contributions, x, output, target, colorSpace);
            }
        }
    }

    private static void ResampleVertical(
        byte[] input,
        int sourceWidth,
        int sourceHeight,
        int destinationHeight,
        AxisContributions contributions,
        float[] output,
        OfficeRasterResamplingColorSpace colorSpace) {
        for (int y = 0; y < destinationHeight; y++) {
            for (int x = 0; x < sourceWidth; x++) {
                int target = ((y * sourceWidth) + x) * 4;
                AccumulateBytes(input, x * 4, sourceWidth * 4, contributions, y, output, target, colorSpace);
            }
        }
    }

    private static void ResampleHorizontal(
        float[] input,
        int sourceWidth,
        int destinationWidth,
        int height,
        AxisContributions contributions,
        byte[] output,
        OfficeRasterResamplingColorSpace colorSpace) {
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < destinationWidth; x++) {
                int target = ((y * destinationWidth) + x) * 4;
                AccumulateFloats(input, (y * sourceWidth) * 4, 4, contributions, x, output, target, colorSpace);
            }
        }
    }

    private static void ResampleVertical(
        float[] input,
        int width,
        int destinationHeight,
        AxisContributions contributions,
        byte[] output,
        OfficeRasterResamplingColorSpace colorSpace) {
        for (int y = 0; y < destinationHeight; y++) {
            for (int x = 0; x < width; x++) {
                int target = ((y * width) + x) * 4;
                AccumulateFloats(input, x * 4, width * 4, contributions, y, output, target, colorSpace);
            }
        }
    }

    private static void AccumulateBytes(
        byte[] input,
        int baseOffset,
        int stride,
        AxisContributions contributions,
        int destination,
        float[] output,
        int target,
        OfficeRasterResamplingColorSpace colorSpace) {
        int start = contributions.Starts[destination];
        int count = contributions.Counts[destination];
        int weights = contributions.Offsets[destination];
        double red = 0D;
        double green = 0D;
        double blue = 0D;
        double alpha = 0D;
        for (int index = 0; index < count; index++) {
            int source = baseOffset + ((start + index) * stride);
            double weight = contributions.Weights[weights + index];
            double sourceAlpha = input[source + 3];
            double premultiply = weight * sourceAlpha / 255D;
            red += DecodeChannel(input[source], colorSpace) * premultiply;
            green += DecodeChannel(input[source + 1], colorSpace) * premultiply;
            blue += DecodeChannel(input[source + 2], colorSpace) * premultiply;
            alpha += sourceAlpha * weight;
        }
        output[target] = (float)red;
        output[target + 1] = (float)green;
        output[target + 2] = (float)blue;
        output[target + 3] = (float)alpha;
    }

    private static void AccumulateFloats(
        float[] input,
        int baseOffset,
        int stride,
        AxisContributions contributions,
        int destination,
        byte[] output,
        int target,
        OfficeRasterResamplingColorSpace colorSpace) {
        int start = contributions.Starts[destination];
        int count = contributions.Counts[destination];
        int weights = contributions.Offsets[destination];
        double red = 0D;
        double green = 0D;
        double blue = 0D;
        double alpha = 0D;
        for (int index = 0; index < count; index++) {
            int source = baseOffset + ((start + index) * stride);
            double weight = contributions.Weights[weights + index];
            red += input[source] * weight;
            green += input[source + 1] * weight;
            blue += input[source + 2] * weight;
            alpha += input[source + 3] * weight;
        }
        WriteStraightRgba(output, target, red, green, blue, alpha, colorSpace);
    }

    private static void WriteStraightRgba(
        byte[] output,
        int target,
        double premultipliedRed,
        double premultipliedGreen,
        double premultipliedBlue,
        double alpha,
        OfficeRasterResamplingColorSpace colorSpace) {
        if (alpha <= 1E-6D) {
            output[target] = output[target + 1] = output[target + 2] = output[target + 3] = 0;
            return;
        }
        output[target] = EncodeChannel(premultipliedRed * 255D / alpha, colorSpace);
        output[target + 1] = EncodeChannel(premultipliedGreen * 255D / alpha, colorSpace);
        output[target + 2] = EncodeChannel(premultipliedBlue * 255D / alpha, colorSpace);
        output[target + 3] = ToByte(alpha);
    }

    private static byte ToByte(double value) =>
        (byte)Math.Round(Clamp(value, 0D, 255D));

    private sealed class AxisContributions {
        internal AxisContributions(int[] starts, int[] counts, int[] offsets, double[] weights) {
            Starts = starts;
            Counts = counts;
            Offsets = offsets;
            Weights = weights;
        }

        internal int[] Starts { get; }
        internal int[] Counts { get; }
        internal int[] Offsets { get; }
        internal double[] Weights { get; }
    }
}
