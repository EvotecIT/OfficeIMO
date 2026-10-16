using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    private static void FillVp8lSourcePixels(byte[] rgba, uint[] pixels, CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null) {
        for (int index = 0; index < pixels.Length; index++) {
            if ((index & 4095) == 0) {
                checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.WebpCompressionBlock);
                cancellationToken.ThrowIfCancellationRequested();
            }
            pixels[index] = PackRgba(rgba, index * 4);
        }
    }

    /// <summary>Selects a predictor per 32-pixel block without retaining another full image.</summary>
    private static uint[] CreateVp8lBlockResiduals(int width, int height, byte[] rgba,
        uint[] residuals, CancellationToken cancellationToken, bool quick = false) {
        int columns = DivideRoundUp(width, 32);
        int rows = DivideRoundUp(height, 32);
        var modes = new uint[columns * rows];
        for (int blockY = 0; blockY < rows; blockY++) {
            for (int blockX = 0; blockX < columns; blockX++) {
                cancellationToken.ThrowIfCancellationRequested();
                int startX = blockX * 32, endX = Math.Min(startX + 32, width);
                int startY = blockY * 32, endY = Math.Min(startY + 32, height);
                int bestMode = 12;
                long bestScore = long.MaxValue;
                // Sample a fixed grid to bound mode search. Mode 12 wins ties so
                // flat areas and gradients don't acquire arbitrary mode changes.
                for (int choice = 0; choice < (quick ? 3 : 14); choice++) {
                    int mode = quick ? choice == 0 ? 12 : choice : choice == 0 ? 12 : choice <= 12 ? choice - 1 : 13;
                    long score = 0;
                    for (int y = startY; y < endY; y += quick ? 8 : 4) {
                        for (int x = startX; x < endX; x += quick ? 8 : 4) {
                            uint value = PackRgba(rgba, (y * width + x) * 4);
                            uint difference = SubtractArgb(value, PredictVp8lSource(rgba, width, x, y, mode));
                            score += Vp8lResidualScore(difference);
                        }
                    }
                    if (score < bestScore) { bestScore = score; bestMode = mode; }
                }
                modes[blockY * columns + blockX] = 0xFF000000U | (uint)(bestMode << 8);
                for (int y = startY; y < endY; y++) {
                    for (int x = startX; x < endX; x++) {
                        int position = y * width + x;
                        residuals[position] = SubtractArgb(PackRgba(rgba, position * 4),
                            PredictVp8lSource(rgba, width, x, y, bestMode));
                    }
                }
            }
        }
        return modes;
    }

    // A fixed sample bounds the Balanced decision without encoding both candidates.
    // Include every alpha sample and the format's actual edge predictor semantics.
    private static bool ShouldUseVp8lPrediction(int width, byte[] rgba, CancellationToken token) {
        long rawScore = 0, predictedScore = 0;
        int count = rgba.Length / 4;
        for (int i = 0, step = Math.Max(1, count / 4096); i < count; i += step) {
            token.ThrowIfCancellationRequested();
            uint value = PackRgba(rgba, i * 4);
            rawScore += Vp8lResidualScore(value);
            predictedScore += Vp8lResidualScore(SubtractArgb(value, PredictVp8lSource(rgba, width, i % width, i / width, 12)));
        }
        return predictedScore * 4 < rawScore * 3;
    }

    private static uint PredictVp8lSource(byte[] rgba, int width, int x, int y, int mode) {
        int offset = (y * width + x) * 4;
        if (y == 0) return x == 0 ? 0xFF000000U : PackRgba(rgba, offset - 4);
        if (x == 0) return PackRgba(rgba, offset - width * 4);
        // VP8L's right-edge TR is the leftmost sample of the current row.
        int topRight = x + 1 < width ? offset - width * 4 + 4 : y * width * 4;
        return PredictVp8l(mode, PackRgba(rgba, offset - 4), PackRgba(rgba, offset - width * 4),
            PackRgba(rgba, offset - width * 4 - 4), PackRgba(rgba, topRight));
    }

    private static int Vp8lResidualScore(uint residual) {
        int result = 0;
        for (int shift = 0; shift <= 24; shift += 8) {
            int value = (int)(residual >> shift) & 255;
            int magnitude = Math.Min(value, 256 - value);
            // Approximate signed residual bit cost, with a strong preference for zero.
            while (magnitude > 0) { result++; magnitude >>= 1; }
        }
        return result;
    }

    private static bool SelectVp8lSubtractGreen(uint[] pixels, CancellationToken cancellationToken) {
        var counts = new int[1024];
        for (int index = 0; index < pixels.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            uint color = pixels[index];
            int red = (int)(color >> 16) & 255, green = (int)(color >> 8) & 255, blue = (int)color & 255;
            counts[red]++;
            counts[256 + blue]++;
            counts[512 + ((red - green) & 255)]++;
            counts[768 + ((blue - green) & 255)]++;
        }
        double original = Vp8lEntropy(counts, 0, pixels.Length) + Vp8lEntropy(counts, 256, pixels.Length);
        double transformed = Vp8lEntropy(counts, 512, pixels.Length) + Vp8lEntropy(counts, 768, pixels.Length);
        if (transformed + 16 >= original) return false;
        ApplyVp8lSubtractGreen(pixels, cancellationToken);
        return true;
    }

    private static void ApplyVp8lSubtractGreen(uint[] pixels, CancellationToken cancellationToken) {
        for (int index = 0; index < pixels.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            uint color = pixels[index];
            uint green = (color >> 8) & 255U;
            pixels[index] = (color & 0xFF00FF00U) |
                (((color & 0x00FF00FFU) + 0x01000100U - green * 0x00010001U) & 0x00FF00FFU);
        }
    }

    /// <summary>Fits bounded signed multipliers, then verifies block entropy before applying them.</summary>
    private static uint[] SelectVp8lColorTransform(int width, int height, uint[] pixels,
        CancellationToken cancellationToken) {
        int columns = DivideRoundUp(width, 32), rows = DivideRoundUp(height, 32);
        var transforms = new uint[columns * rows];
        var counts = new int[256];
        double savings = 0;
        for (int by = 0; by < rows; by++) {
            for (int bx = 0; bx < columns; bx++) {
                cancellationToken.ThrowIfCancellationRequested();
                int startX = bx * 32, endX = Math.Min(startX + 32, width);
                int startY = by * 32, endY = Math.Min(startY + 32, height);
                int greenToRed = SelectVp8lColorMultiplier(pixels, width, startX, endX, startY, endY,
                    counts, true, 0, cancellationToken);
                int greenToBlue = SelectVp8lColorMultiplier(pixels, width, startX, endX, startY, endY,
                    counts, false, 0, cancellationToken);
                int redToBlue = SelectVp8lColorMultiplier(pixels, width, startX, endX, startY, endY,
                    counts, false, greenToBlue, cancellationToken, fitRed: true);
                double original = Vp8lColorEntropy(pixels, width, startX, endX, startY, endY, counts, true, 0, 0, 1) +
                    Vp8lColorEntropy(pixels, width, startX, endX, startY, endY, counts, false, 0, 0, 1);
                double transformed = Vp8lColorEntropy(pixels, width, startX, endX, startY, endY, counts, true, greenToRed, 0, 1) +
                    Vp8lColorEntropy(pixels, width, startX, endX, startY, endY, counts, false, greenToBlue, redToBlue, 1);
                if (original > transformed + 16) {
                    savings += original - transformed;
                    transforms[by * columns + bx] = 0xFF000000U | ((uint)(redToBlue & 255) << 16) |
                        ((uint)(greenToBlue & 255) << 8) | (uint)(greenToRed & 255);
                } else {
                    transforms[by * columns + bx] = 0xFF000000U;
                }
            }
        }
        // Include a conservative map/header charge; avoid fitting noise at the
        // cost of a larger bitstream. Final raw-vs-transformed selection is exact.
        if (savings <= transforms.Length * 32L + 128L) return Array.Empty<uint>();
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                if ((x & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                uint data = transforms[(y >> 5) * columns + (x >> 5)];
                uint color = pixels[y * width + x];
                int green = (int)(color >> 8) & 255, red = (int)(color >> 16) & 255, blue = (int)color & 255;
                int newRed = (red - ColorTransformDelta(unchecked((sbyte)data), green)) & 255;
                int newBlue = (blue - ColorTransformDelta(unchecked((sbyte)(data >> 8)), green) -
                    ColorTransformDelta(unchecked((sbyte)(data >> 16)), red)) & 255;
                pixels[y * width + x] = (color & 0xFF00FF00U) | (uint)(newRed << 16 | newBlue);
            }
        }
        return transforms;
    }

    private static int SelectVp8lColorMultiplier(uint[] pixels, int width, int startX, int endX, int startY, int endY,
        int[] counts, bool red, int greenMultiplier, CancellationToken cancellationToken, bool fitRed = false) {
        int best = 0;
        double bestScore = Vp8lColorEntropy(pixels, width, startX, endX, startY, endY,
            counts, red, greenMultiplier, 0, 4);
        // Fixed-point coefficients 0, +/-0.5, +/-1 and +/-2, searched deterministically.
        for (int choice = 0; choice < 6; choice++) {
            cancellationToken.ThrowIfCancellationRequested();
            int multiplier = (16 << (choice / 2)) * ((choice & 1) == 0 ? 1 : -1);
            double score = Vp8lColorEntropy(pixels, width, startX, endX, startY, endY,
                counts, red, fitRed ? greenMultiplier : multiplier, fitRed ? multiplier : 0, 4);
            if (score + 1 < bestScore) { bestScore = score; best = multiplier; }
        }
        return best;
    }

    private static double Vp8lColorEntropy(uint[] pixels, int width, int startX, int endX, int startY, int endY,
        int[] counts, bool redChannel, int greenMultiplier, int redMultiplier, int step) {
        Array.Clear(counts, 0, counts.Length);
        int samples = 0;
        for (int y = startY; y < endY; y += step) {
            for (int x = startX; x < endX; x += step) {
                uint color = pixels[y * width + x];
                int red = (int)(color >> 16) & 255, green = (int)(color >> 8) & 255;
                int value = redChannel ? red : (int)color & 255;
                value -= ColorTransformDelta(greenMultiplier, green);
                if (!redChannel) value -= ColorTransformDelta(redMultiplier, red);
                counts[value & 255]++;
                samples++;
            }
        }
        return Vp8lEntropy(counts, 0, samples);
    }

    private static double Vp8lEntropy(int[] counts, int offset, int total) {
        double bits = 0;
        for (int value = 0; value < 256; value++) {
            int count = counts[offset + value];
            if (count > 0) bits += count * Math.Log((double)total / count, 2);
        }
        return bits;
    }
}
