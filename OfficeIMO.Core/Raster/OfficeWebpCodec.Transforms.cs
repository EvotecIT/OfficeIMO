using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    private static bool TryReadVp8lTransforms(
        LsbBitReader reader,
        int width,
        int height,
        Vp8lAllocationBudget allocationBudget,
        CancellationToken cancellationToken,
        out int encodedWidth,
        out List<Vp8lTransform> transforms) {
        if (!allocationBudget.TryReserveBytes(256L)) {
            transforms = new List<Vp8lTransform>();
            encodedWidth = width;
            return false;
        }
        transforms = new List<Vp8lTransform>(4);
        encodedWidth = width;
        int seen = 0;
        while (reader.ReadBits(1) != 0) {
            int type = (int)reader.ReadBits(2);
            int mask = 1 << type;
            if ((seen & mask) != 0) return false;
            seen |= mask;
            if (type == 2) {
                transforms.Add(new Vp8lTransform(type, encodedWidth, encodedWidth, 0, Array.Empty<uint>(), 0));
                continue;
            }
            if (type == 3) {
                int tableSize = (int)reader.ReadBits(8) + 1;
                if (!TryDecodeVp8lImageData(reader, tableSize, 1, false, 1,
                        allocationBudget, cancellationToken, out uint[] palette)) return false;
                uint previous = 0;
                for (int index = 0; index < palette.Length; index++) {
                    previous = AddArgb(previous, palette[index]);
                    palette[index] = previous;
                }
                int widthBits = tableSize <= 2 ? 3 : tableSize <= 4 ? 2 : tableSize <= 16 ? 1 : 0;
                int packedWidth = DivideRoundUp(encodedWidth, 1 << widthBits);
                transforms.Add(new Vp8lTransform(type, encodedWidth, packedWidth, widthBits, palette, tableSize));
                encodedWidth = packedWidth;
                continue;
            }

            int sizeBits = (int)reader.ReadBits(3) + 2;
            int transformWidth = DivideRoundUp(encodedWidth, 1 << sizeBits);
            int transformHeight = DivideRoundUp(height, 1 << sizeBits);
            if (!TryDecodeVp8lImageData(reader, transformWidth, transformHeight, false, 1,
                    allocationBudget, cancellationToken, out uint[] data)) return false;
            transforms.Add(new Vp8lTransform(type, encodedWidth, encodedWidth, sizeBits, data, transformWidth));
        }
        return true;
    }

    private static bool TryApplyVp8lTransforms(
        uint[] encoded,
        int encodedWidth,
        int height,
        int finalWidth,
        List<Vp8lTransform> transforms,
        Vp8lAllocationBudget allocationBudget,
        CancellationToken cancellationToken,
        out uint[] result) {
        result = encoded;
        int width = encodedWidth;
        for (int index = transforms.Count - 1; index >= 0; index--) {
            Vp8lTransform transform = transforms[index];
            if (transform.OutputWidth != width) return false;
            switch (transform.Type) {
                case 0:
                    if (!TryApplyPredictorTransform(result, width, height, transform,
                            allocationBudget, cancellationToken, out result)) return false;
                    break;
                case 1:
                    if (!TryApplyColorTransform(result, width, height, transform, cancellationToken)) return false;
                    break;
                case 2:
                    for (int pixel = 0; pixel < result.Length; pixel++) {
                        if ((pixel & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                        uint color = result[pixel];
                        int green = (int)(color >> 8) & 255;
                        int red = ((int)(color >> 16) + green) & 255;
                        int blue = ((int)color + green) & 255;
                        result[pixel] = (color & 0xFF00FF00U) | (uint)(red << 16 | blue);
                    }
                    break;
                case 3:
                    if (!TryApplyPaletteTransform(result, width, height, transform,
                            allocationBudget, cancellationToken, out result)) return false;
                    width = transform.InputWidth;
                    break;
                default:
                    return false;
            }
        }
        return width == finalWidth && result.Length == checked(finalWidth * height);
    }

    private static bool TryApplyPaletteTransform(
        uint[] packed,
        int packedWidth,
        int height,
        Vp8lTransform transform,
        Vp8lAllocationBudget allocationBudget,
        CancellationToken cancellationToken,
        out uint[] output) {
        int outputLength = checked(transform.InputWidth * height);
        if (!allocationBudget.TryReserveArray(outputLength, sizeof(uint))) {
            output = Array.Empty<uint>();
            return false;
        }
        output = new uint[outputLength];
        int pixelsPerPacked = 1 << transform.SizeBits;
        int bitsPerIndex = 8 >> transform.SizeBits;
        int mask = (1 << bitsPerIndex) - 1;
        for (int y = 0; y < height; y++) {
            if ((y & 31) == 0) cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < transform.InputWidth; x++) {
                int packedX = x / pixelsPerPacked;
                int shift = (x & (pixelsPerPacked - 1)) * bitsPerIndex;
                int paletteIndex = (((int)(packed[y * packedWidth + packedX] >> 8)) >> shift) & mask;
                output[y * transform.InputWidth + x] = paletteIndex < transform.Value
                    ? transform.Data[paletteIndex]
                    : 0U;
            }
        }
        return true;
    }

    private static bool TryApplyColorTransform(
        uint[] pixels,
        int width,
        int height,
        Vp8lTransform transform,
        CancellationToken cancellationToken) {
        int blockWidth = transform.Value;
        if (blockWidth < 1) return false;
        for (int y = 0; y < height; y++) {
            if ((y & 31) == 0) cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                uint color = pixels[y * width + x];
                uint data = transform.Data[(y >> transform.SizeBits) * blockWidth + (x >> transform.SizeBits)];
                int green = (int)(color >> 8) & 255;
                int red = (int)(color >> 16) & 255;
                int blue = (int)color & 255;
                int greenToRed = unchecked((sbyte)data);
                int greenToBlue = unchecked((sbyte)(data >> 8));
                int redToBlue = unchecked((sbyte)(data >> 16));
                int restoredRed = (red + ColorTransformDelta(greenToRed, green)) & 255;
                int restoredBlue = (blue + ColorTransformDelta(greenToBlue, green) +
                    ColorTransformDelta(redToBlue, restoredRed)) & 255;
                pixels[y * width + x] = (color & 0xFF00FF00U) | (uint)(restoredRed << 16 | restoredBlue);
            }
        }
        return true;
    }

    private static bool TryApplyPredictorTransform(
        uint[] residuals,
        int width,
        int height,
        Vp8lTransform transform,
        Vp8lAllocationBudget allocationBudget,
        CancellationToken cancellationToken,
        out uint[] output) {
        if (!allocationBudget.TryReserveArray(residuals.Length, sizeof(uint))) {
            output = Array.Empty<uint>();
            return false;
        }
        output = new uint[residuals.Length];
        int blockWidth = transform.Value;
        if (blockWidth < 1) return false;
        for (int y = 0; y < height; y++) {
            if ((y & 31) == 0) cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                uint prediction;
                if (x == 0 && y == 0) {
                    prediction = 0xFF000000U;
                } else if (y == 0) {
                    prediction = output[x - 1];
                } else if (x == 0) {
                    prediction = output[(y - 1) * width];
                } else {
                    uint left = output[y * width + x - 1];
                    uint top = output[(y - 1) * width + x];
                    uint topLeft = output[(y - 1) * width + x - 1];
                    uint topRight = x + 1 < width
                        ? output[(y - 1) * width + x + 1]
                        : output[y * width];
                    int mode = (int)(transform.Data[(y >> transform.SizeBits) * blockWidth + (x >> transform.SizeBits)] >> 8) & 255;
                    if (mode > 13) return false;
                    prediction = PredictVp8l(mode, left, top, topLeft, topRight);
                }
                output[y * width + x] = AddArgb(residuals[y * width + x], prediction);
            }
        }
        return true;
    }

    internal static uint PredictVp8l(int mode, uint left, uint top, uint topLeft, uint topRight) {
        switch (mode) {
            case 0: return 0xFF000000U;
            case 1: return left;
            case 2: return top;
            case 3: return topRight;
            case 4: return topLeft;
            case 5: return AverageArgb(AverageArgb(left, topRight), top);
            case 6: return AverageArgb(left, topLeft);
            case 7: return AverageArgb(left, top);
            case 8: return AverageArgb(topLeft, top);
            case 9: return AverageArgb(top, topRight);
            case 10: return AverageArgb(AverageArgb(left, topLeft), AverageArgb(top, topRight));
            case 11: return SelectArgb(left, top, topLeft);
            case 12: return ComponentOperation(left, top, topLeft, 0);
            case 13: return ComponentOperation(AverageArgb(left, top), 0U, topLeft, 1);
            default: return 0U;
        }
    }

    private static uint AddArgb(uint first, uint second) {
        uint result = 0;
        for (int shift = 0; shift <= 24; shift += 8) {
            result |= (uint)((((first >> shift) & 255U) + ((second >> shift) & 255U)) & 255U) << shift;
        }
        return result;
    }

    private static uint AverageArgb(uint first, uint second) {
        uint result = 0;
        for (int shift = 0; shift <= 24; shift += 8) {
            result |= ((((first >> shift) & 255U) + ((second >> shift) & 255U)) >> 1) << shift;
        }
        return result;
    }

    private static uint SelectArgb(uint left, uint top, uint topLeft) {
        int leftDistance = 0;
        int topDistance = 0;
        for (int shift = 0; shift <= 24; shift += 8) {
            int l = (int)(left >> shift) & 255;
            int t = (int)(top >> shift) & 255;
            int estimate = l + t - ((int)(topLeft >> shift) & 255);
            leftDistance += Math.Abs(estimate - l);
            topDistance += Math.Abs(estimate - t);
        }
        return leftDistance < topDistance ? left : top;
    }

    private static uint ComponentOperation(uint first, uint second, uint subtract, int half) {
        uint result = 0;
        for (int shift = 0; shift <= 24; shift += 8) {
            int a = (int)(first >> shift) & 255;
            int b = (int)(second >> shift) & 255;
            int c = (int)(subtract >> shift) & 255;
            int value = half == 0 ? a + b - c : a + ((a - c) >> 1);
            value = Math.Max(0, Math.Min(255, value));
            result |= (uint)value << shift;
        }
        return result;
    }

    private static int ColorTransformDelta(int transform, int color) =>
        (transform * unchecked((sbyte)color)) >> 5;

    private sealed class Vp8lTransform {
        internal Vp8lTransform(int type, int inputWidth, int outputWidth, int sizeBits,
            uint[] data, int value) {
            Type = type;
            InputWidth = inputWidth;
            OutputWidth = outputWidth;
            SizeBits = sizeBits;
            Data = data;
            Value = value;
        }
        internal int Type { get; }
        internal int InputWidth { get; }
        internal int OutputWidth { get; }
        internal int SizeBits { get; }
        internal uint[] Data { get; }
        internal int Value { get; }
    }
}
