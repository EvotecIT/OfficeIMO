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

    // The decoder's neighbourhood: the first row predicts from the left, the first column from above,
    // and the last column's top-right neighbour is the first pixel of the current row.
    private static uint PredictVp8lPixel(uint[] pixels, int width, int x, int y, int mode) {
        int position = y * width + x;
        if (y == 0) return x == 0 ? 0xFF000000U : pixels[position - 1];
        if (x == 0) return pixels[position - width];
        uint topRight = x + 1 < width ? pixels[position - width + 1] : pixels[y * width];
        return PredictVp8l(mode, pixels[position - 1], pixels[position - width], pixels[position - width - 1], topRight);
    }

    private static uint SubtractGreenFromResidual(uint value) {
        uint green = (value >> 8) & 255U;
        return (value & 0xFF00FF00U) | (((value & 0x00FF00FFU) + 0x01000100U - green * 0x00010001U) & 0x00FF00FFU);
    }

    /// <summary>
    /// Decides subtract-green from gradient residuals of a row sample: it pays when red and blue
    /// follow green, which holds for most photographs and many rendered images.
    /// </summary>
    private static bool ShouldSubtractVp8lGreen(uint[] pixels, int width, int height, CancellationToken cancellationToken) {
        var counts = new int[4 * 256];
        int step = Math.Max(1, height / 64), samples = 0;
        for (int y = 1; y < height; y += step) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 1; x < width; x++) {
                uint residual = SubtractArgb(pixels[y * width + x], PredictVp8lPixel(pixels, width, x, y, 11));
                int red = (int)(residual >> 16) & 255, green = (int)(residual >> 8) & 255, blue = (int)residual & 255;
                counts[red]++;
                counts[256 + blue]++;
                counts[512 + ((red - green) & 255)]++;
                counts[768 + ((blue - green) & 255)]++;
                samples++;
            }
        }
        if (samples == 0) return false;
        return Vp8lShannonBits(counts, 512, 256) + Vp8lShannonBits(counts, 768, 256) + 16 <
            Vp8lShannonBits(counts, 0, 256) + Vp8lShannonBits(counts, 256, 256);
    }

    /// <summary>
    /// Chooses a predictor per tile by the bits its residuals would take under a code fitted to the
    /// residuals chosen so far, then writes the residuals (green already subtracted when requested).
    /// </summary>
    private static uint[] SelectVp8lPredictors(uint[] pixels, uint[] residuals, int width, int height, int bits,
        int[] modes, bool subtractGreen, CancellationToken cancellationToken) {
        int size = 1 << bits;
        int columns = DivideRoundUp(width, size), rows = DivideRoundUp(height, size);
        var map = new uint[columns * rows];
        var accumulated = new int[4 * 256];
        var costs = new float[4 * 256];
        int accumulatedCount = 0;
        UpdateVp8lResidualCosts(accumulated, accumulatedCount, costs);
        var scores = new float[14];
        for (int tileY = 0; tileY < rows; tileY++) {
            cancellationToken.ThrowIfCancellationRequested();
            int startY = tileY * size, endY = Math.Min(startY + size, height);
            for (int tileX = 0; tileX < columns; tileX++) {
                int startX = tileX * size, endX = Math.Min(startX + size, width);
                foreach (int mode in modes) {
                    scores[mode] = ScoreVp8lPredictor(pixels, width, startX, endX, startY, endY, mode, subtractGreen, costs);
                }
                int best = modes[0];
                // Keeping the neighbours' mode makes the predictor map cheaper; it wins small ties.
                uint leftMode = tileX > 0 ? (map[tileY * columns + tileX - 1] >> 8) & 255U : uint.MaxValue;
                uint topMode = tileY > 0 ? (map[(tileY - 1) * columns + tileX] >> 8) & 255U : uint.MaxValue;
                float bestScore = float.MaxValue;
                foreach (int mode in modes) {
                    float score = scores[mode] * (mode == leftMode || mode == topMode ? 0.99f : 1f);
                    if (score < bestScore) {
                        bestScore = score;
                        best = mode;
                    }
                }
                map[tileY * columns + tileX] = 0xFF000000U | (uint)best << 8;
                for (int y = startY; y < endY; y++) {
                    for (int x = startX; x < endX; x++) {
                        int position = y * width + x;
                        uint residual = SubtractArgb(pixels[position], PredictVp8lPixel(pixels, width, x, y, best));
                        if (subtractGreen) residual = SubtractGreenFromResidual(residual);
                        residuals[position] = residual;
                        accumulated[residual & 255]++;
                        accumulated[256 + ((residual >> 8) & 255)]++;
                        accumulated[512 + ((residual >> 16) & 255)]++;
                        accumulated[768 + (residual >> 24)]++;
                        accumulatedCount++;
                    }
                }
            }
            UpdateVp8lResidualCosts(accumulated, accumulatedCount, costs);
        }
        return map;
    }

    // Bits of one tile's residuals under one mode. Rows and columns zero are predicted the same way by
    // every mode, so they are left out; the mode is fixed for the whole loop, so its branch is predictable.
    private static float ScoreVp8lPredictor(uint[] pixels, int width, int startX, int endX, int startY, int endY, int mode,
        bool subtractGreen, float[] costs) {
        float score = 0;
        for (int y = Math.Max(1, startY); y < endY; y++) {
            int row = y * width;
            for (int x = Math.Max(1, startX); x < endX; x++) {
                int position = row + x;
                uint topRight = x + 1 < width ? pixels[position - width + 1] : pixels[row];
                uint residual = SubtractArgb(pixels[position],
                    PredictVp8l(mode, pixels[position - 1], pixels[position - width], pixels[position - width - 1], topRight));
                if (subtractGreen) residual = SubtractGreenFromResidual(residual);
                score += costs[residual & 255] + costs[256 + ((residual >> 8) & 255)] +
                    costs[512 + ((residual >> 16) & 255)] + costs[768 + (residual >> 24)];
            }
        }
        return score;
    }

    /// <summary>Predicts every pixel with one mode and returns the matching (uniform) predictor map.</summary>
    private static uint[] ApplyVp8lUniformPredictor(uint[] pixels, uint[] residuals, int width, int height, int bits,
        int mode, bool subtractGreen, CancellationToken cancellationToken) {
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                int position = y * width + x;
                uint residual = SubtractArgb(pixels[position], PredictVp8lPixel(pixels, width, x, y, mode));
                residuals[position] = subtractGreen ? SubtractGreenFromResidual(residual) : residual;
            }
        }
        var map = new uint[DivideRoundUp(width, 1 << bits) * DivideRoundUp(height, 1 << bits)];
        for (int index = 0; index < map.Length; index++) map[index] = 0xFF000000U | (uint)mode << 8;
        return map;
    }

    // Bits per residual value under the residuals so far; the prior keeps unseen values plausible
    // and favours small magnitudes before anything has been seen.
    private static void UpdateVp8lResidualCosts(int[] accumulated, int count, float[] costs) {
        for (int channel = 0; channel < 4; channel++) {
            double total = count + 256.0 * 0.5;
            for (int value = 0; value < 256; value++) {
                int magnitude = Math.Min(value, 256 - value);
                double prior = 0.5 * 4.0 / (4.0 + magnitude);
                costs[channel * 256 + value] = (float)Log2((total + 256) / (accumulated[channel * 256 + value] + prior + 0.05));
            }
        }
    }

    /// <summary>
    /// Fits the three cross-colour multipliers per tile, searching coarse to fine, and applies them.
    /// Returns an empty map when the saving would not pay for the map.
    /// </summary>
    private static uint[] SelectVp8lColorTransform(uint[] residuals, int width, int height, int bits, int sampleStep,
        CancellationToken cancellationToken) {
        int size = 1 << bits;
        int columns = DivideRoundUp(width, size), rows = DivideRoundUp(height, size);
        var map = new uint[columns * rows];
        var counts = new int[256];
        var samples = new uint[size * size];
        double savings = 0;
        for (int tileY = 0; tileY < rows; tileY++) {
            cancellationToken.ThrowIfCancellationRequested();
            int startY = tileY * size, endY = Math.Min(startY + size, height);
            for (int tileX = 0; tileX < columns; tileX++) {
                int startX = tileX * size, endX = Math.Min(startX + size, width);
                int sampleCount = 0;
                for (int y = startY; y < endY; y += sampleStep) {
                    for (int x = startX; x < endX; x += sampleStep) samples[sampleCount++] = residuals[y * width + x];
                }
                uint left = tileX > 0 ? map[tileY * columns + tileX - 1] : 0U;
                uint top = tileY > 0 ? map[(tileY - 1) * columns + tileX] : 0U;
                int greenToRed = SearchVp8lMultiplier(samples, sampleCount, counts, Vp8lColorTarget.Red, 0,
                    (sbyte)left, (sbyte)top);
                int greenToBlue = SearchVp8lMultiplier(samples, sampleCount, counts, Vp8lColorTarget.BlueFromGreen, 0,
                    (sbyte)(left >> 8), (sbyte)(top >> 8));
                int redToBlue = SearchVp8lMultiplier(samples, sampleCount, counts, Vp8lColorTarget.BlueFromRed, greenToBlue,
                    (sbyte)(left >> 16), (sbyte)(top >> 16));
                greenToBlue = SearchVp8lMultiplier(samples, sampleCount, counts, Vp8lColorTarget.BlueFromGreen, redToBlue,
                    greenToBlue, greenToBlue);
                double original = Vp8lColorCost(samples, sampleCount, counts, Vp8lColorTarget.Red, 0, 0) +
                    Vp8lColorCost(samples, sampleCount, counts, Vp8lColorTarget.BlueFromGreen, 0, 0);
                double transformed = Vp8lColorCost(samples, sampleCount, counts, Vp8lColorTarget.Red, greenToRed, 0) +
                    Vp8lColorCost(samples, sampleCount, counts, Vp8lColorTarget.BlueFromGreen, greenToBlue, redToBlue);
                savings += (original - transformed) * sampleStep * sampleStep;
                map[tileY * columns + tileX] = 0xFF000000U | (uint)(redToBlue & 255) << 16 |
                    (uint)(greenToBlue & 255) << 8 | (uint)(greenToRed & 255);
            }
        }
        if (savings <= map.Length * 12L + 256L) return Array.Empty<uint>();
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            int rowStart = y * width;
            for (int x = 0; x < width; x++) {
                uint data = map[(y >> bits) * columns + (x >> bits)];
                residuals[rowStart + x] = ApplyVp8lColorTransform(residuals[rowStart + x], data);
            }
        }
        return map;
    }

    private static uint ApplyVp8lColorTransform(uint color, uint data) {
        int green = (int)(color >> 8) & 255, red = (int)(color >> 16) & 255, blue = (int)color & 255;
        int newRed = (red - ColorTransformDelta(unchecked((sbyte)data), green)) & 255;
        int newBlue = (blue - ColorTransformDelta(unchecked((sbyte)(data >> 8)), green) -
            ColorTransformDelta(unchecked((sbyte)(data >> 16)), red)) & 255;
        return (color & 0xFF00FF00U) | (uint)(newRed << 16 | newBlue);
    }

    private enum Vp8lColorTarget {
        Red,
        BlueFromGreen,
        BlueFromRed
    }

    private static int SearchVp8lMultiplier(uint[] samples, int count, int[] counts, Vp8lColorTarget target,
        int other, int left, int top) {
        int best = 0;
        double bestCost = Cost(0);
        // Keeping a neighbour's value keeps the map cheap; a new value must clearly win.
        Consider(left, 0.5);
        Consider(top, 0.5);
        for (int value = -128; value < 128; value += 32) Consider(value, 2);
        for (int step = 16; step >= 1; step >>= 1) {
            int center = best;
            Consider(center - step, 2);
            Consider(center + step, 2);
        }
        return best;

        void Consider(int value, double penalty) {
            if (value < -128 || value > 127 || value == best) return;
            double cost = Cost(value) + (value == left || value == top ? 0 : penalty);
            if (cost < bestCost) {
                bestCost = cost;
                best = value;
            }
        }

        double Cost(int value) => target switch {
            Vp8lColorTarget.Red => Vp8lColorCost(samples, count, counts, target, value, 0),
            Vp8lColorTarget.BlueFromGreen => Vp8lColorCost(samples, count, counts, target, value, other),
            _ => Vp8lColorCost(samples, count, counts, target, other, value)
        };
    }

    // Shannon bits of the transformed red or blue channel over the tile sample.
    private static double Vp8lColorCost(uint[] samples, int count, int[] counts, Vp8lColorTarget target,
        int greenMultiplier, int redMultiplier) {
        Array.Clear(counts, 0, counts.Length);
        if (target == Vp8lColorTarget.Red) {
            for (int index = 0; index < count; index++) {
                uint color = samples[index];
                int red = (int)(color >> 16) & 255, green = (int)(color >> 8) & 255;
                counts[(red - ColorTransformDelta(greenMultiplier, green)) & 255]++;
            }
        } else {
            for (int index = 0; index < count; index++) {
                uint color = samples[index];
                int red = (int)(color >> 16) & 255, green = (int)(color >> 8) & 255, blue = (int)color & 255;
                counts[(blue - ColorTransformDelta(greenMultiplier, green) - ColorTransformDelta(redMultiplier, red)) & 255]++;
            }
        }
        return Vp8lShannonBits(counts, 0, 256);
    }

    private static uint PackRgba(byte[] rgba, int offset) =>
        (uint)(rgba[offset + 3] << 24 | rgba[offset] << 16 | rgba[offset + 1] << 8 | rgba[offset + 2]);

    private static uint SubtractArgb(uint value, uint predictor) {
        // Byte-wise subtraction modulo 256: borrow-free on the even and odd lanes separately.
        uint evenLanes = (value & 0x00FF00FFU) + 0x01000100U - (predictor & 0x00FF00FFU);
        uint oddLanes = ((value >> 8) & 0x00FF00FFU) + 0x01000100U - ((predictor >> 8) & 0x00FF00FFU);
        return (evenLanes & 0x00FF00FFU) | (oddLanes & 0x00FF00FFU) << 8;
    }
}
