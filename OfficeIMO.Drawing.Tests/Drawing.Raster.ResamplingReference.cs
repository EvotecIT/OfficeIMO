using System;
using System.Collections.Generic;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterResamplingReferenceTests {
    [Theory]
    [InlineData(131, 97)]
    [InlineData(347, 31)]
    public void HighQualityResizeCancelledDuringEitherPassOrderLeavesTheSourceUntouched(int width, int height) {
        var source = new OfficeRasterImage(513, 257, OfficeColor.FromRgba(123, 45, 67, 128));
        byte[] original = source.GetPixels();
        using var cancellation = new CancellationTokenSource();
        bool started = false;

        OperationCanceledException exception = Assert.ThrowsAny<OperationCanceledException>(() =>
            OfficeRasterResampler.Resize(source, width, height, OfficeRasterResamplingMode.Lanczos3,
                OfficeRasterResamplingColorSpace.EncodedSrgb, retainedManagedBytes: 0,
                cancellationToken: cancellation.Token, resamplingWorkStarted: () => {
                    started = true;
                    cancellation.Cancel();
                }));

        Assert.True(started);
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(original, source.GetPixels());
    }

    public static IEnumerable<object[]> Cases() {
        foreach (var size in new[] { (4, 7), (29, 3), (4, 17), (41, 13) })
            foreach (OfficeRasterResamplingMode mode in new[] { OfficeRasterResamplingMode.Area, OfficeRasterResamplingMode.Lanczos3 })
                foreach (OfficeRasterResamplingColorSpace space in new[] { OfficeRasterResamplingColorSpace.EncodedSrgb, OfficeRasterResamplingColorSpace.LinearLight })
                    yield return new object[] { 17, 11, size.Item1, size.Item2, mode, space };
        foreach (var size in new[] { (131, 97), (347, 31), (400, 200) })
            foreach (OfficeRasterResamplingColorSpace space in new[] { OfficeRasterResamplingColorSpace.EncodedSrgb, OfficeRasterResamplingColorSpace.LinearLight })
                yield return new object[] { 513, 257, size.Item1, size.Item2, OfficeRasterResamplingMode.Lanczos3, space };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void HighQualityResizeMatchesIndependentCoverageAndSincReference(
        int sourceWidth, int sourceHeight, int width, int height, OfficeRasterResamplingMode mode, OfficeRasterResamplingColorSpace space) {
        var source = new OfficeRasterImage(sourceWidth, sourceHeight);
        var random = new Random(719);
        for (int y = 0; y < source.Height; y++) {
            for (int x = 0; x < source.Width; x++) {
                source.SetPixel(x, y, OfficeColor.FromRgba(
                    (byte)random.Next(256), (byte)random.Next(256), (byte)random.Next(256),
                    (byte)((x + y) % 4 == 0 ? 0 : random.Next(256))));
            }
        }

        // Full-axis mathematical reference: it does not use production contribution
        // ranges or accumulation helpers. Mixed scaling exercises both pass orders,
        // clipped edge support, hidden transparent colors, and all four vector lanes.
        byte[] expected = Reference(source, width, height, mode, space);
        byte[] actual = OfficeRasterResampler.Resize(source, width, height, mode, space).GetPixels();

        Assert.Equal(expected, actual);
        if (sourceWidth > 17)
            Assert.Equal(actual, OfficeRasterResampler.Resize(source, width, height, mode, space).GetPixels());
    }

    private static byte[] Reference(OfficeRasterImage source, int width, int height,
        OfficeRasterResamplingMode mode, OfficeRasterResamplingColorSpace space) {
        double[][] horizontal = Weights(source.Width, width, mode);
        double[][] vertical = Weights(source.Height, height, mode);
        bool horizontalFirst = (long)width * source.Height <= (long)source.Width * height;
        int intermediateWidth = horizontalFirst ? width : source.Width;
        int intermediateHeight = horizontalFirst ? source.Height : height;
        var intermediate = new float[intermediateWidth * intermediateHeight * 4];
        for (int y = 0; y < intermediateHeight; y++) {
            for (int x = 0; x < intermediateWidth; x++) {
                var sum = new double[4];
                double[] weights = horizontalFirst ? horizontal[x] : vertical[y];
                for (int i = 0; i < weights.Length; i++) {
                    OfficeColor pixel = source.GetPixel(horizontalFirst ? i : x, horizontalFirst ? y : i);
                    double premultiply = weights[i] * pixel.A / 255D;
                    sum[0] += Decode(pixel.R, space) * premultiply;
                    sum[1] += Decode(pixel.G, space) * premultiply;
                    sum[2] += Decode(pixel.B, space) * premultiply;
                    sum[3] += pixel.A * weights[i];
                }
                for (int channel = 0; channel < 4; channel++)
                    intermediate[((y * intermediateWidth + x) * 4) + channel] = (float)sum[channel];
            }
        }
        var result = new byte[width * height * 4];
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                var sum = new double[4];
                double[] weights = horizontalFirst ? vertical[y] : horizontal[x];
                for (int i = 0; i < weights.Length; i++) {
                    int index = ((horizontalFirst ? i * intermediateWidth + x : y * intermediateWidth + i) * 4);
                    for (int channel = 0; channel < 4; channel++)
                        sum[channel] += intermediate[index + channel] * weights[i];
                }
                if (sum[3] <= 1E-6D) continue;
                int target = (y * width + x) * 4;
                for (int channel = 0; channel < 3; channel++) {
                    double value = sum[channel] * 255D / sum[3];
                    if (space == OfficeRasterResamplingColorSpace.LinearLight) {
                        double linear = Math.Max(0D, Math.Min(1D, value / 255D));
                        value = 255D * (linear <= 0.0031308D ? 12.92D * linear : 1.055D * Math.Pow(linear, 1D / 2.4D) - 0.055D);
                    }
                    result[target + channel] = ToByte(value);
                }
                result[target + 3] = ToByte(sum[3]);
            }
        }
        return result;
    }

    private static double[][] Weights(int sourceLength, int destinationLength, OfficeRasterResamplingMode mode) {
        var result = new double[destinationLength][];
        double scale = sourceLength / (double)destinationLength;
        for (int destination = 0; destination < destinationLength; destination++) {
            var weights = new double[sourceLength];
            double center = (destination + 0.5D) * scale - 0.5D;
            for (int source = 0; source < sourceLength; source++) {
                if (mode == OfficeRasterResamplingMode.Area) {
                    weights[source] = scale > 1D
                        ? Math.Max(0D, Math.Min((destination + 1D) * scale, source + 1D) - Math.Max(destination * scale, source))
                        : Math.Max(0D, 1D - Math.Abs(source - Math.Max(0D, Math.Min(sourceLength - 1D, center))));
                } else {
                    double distance = (center - source) / Math.Max(1D, scale);
                    weights[source] = Math.Abs(distance) < 1E-12D ? 1D : Math.Abs(distance) >= 3D ? 0D
                        : Math.Sin(Math.PI * distance) / (Math.PI * distance) *
                          (Math.Sin(Math.PI * distance / 3D) / (Math.PI * distance / 3D));
                }
            }
            double total = 0D;
            foreach (double weight in weights) total += weight;
            for (int source = 0; source < sourceLength; source++) weights[source] /= total;
            result[destination] = weights;
        }
        return result;
    }

    private static double Decode(byte value, OfficeRasterResamplingColorSpace space) {
        if (space == OfficeRasterResamplingColorSpace.EncodedSrgb) return value;
        double encoded = value / 255D;
        return 255D * (encoded <= 0.04045D ? encoded / 12.92D : Math.Pow((encoded + 0.055D) / 1.055D, 2.4D));
    }

    private static byte ToByte(double value) => (byte)Math.Round(Math.Max(0D, Math.Min(255D, value)));
}
