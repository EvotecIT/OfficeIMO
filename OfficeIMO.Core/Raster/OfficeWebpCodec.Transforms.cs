#if NET8_0_OR_GREATER
using Vp8lPixels = System.Span<uint>;
#else
using Vp8lPixels = uint[];
#endif

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

    /// <summary>
    /// Inverts the transforms row by row so each row stays in cache across every stage. Stages that
    /// follow the predictor run one row behind it, because the predictor reads the previous row in
    /// its own output state. When <paramref name="rgba"/> is supplied, the last stage writes RGBA bytes,
    /// and <paramref name="encoded"/> may be the tail of that same buffer: output row y ends where
    /// encoded row y + 1 starts, so writing a row never overwrites encoded rows still to be restored.
    /// </summary>
    private static bool TryApplyVp8lTransforms(
        Vp8lPixels encoded,
        uint[]? encodedArray,
        int encodedWidth,
        int height,
        int finalWidth,
        List<Vp8lTransform> transforms,
        Vp8lAllocationBudget allocationBudget,
        CancellationToken cancellationToken,
        byte[]? rgba,
        out uint[] result) {
        result = encodedArray ?? Array.Empty<uint>();
        int stageCount = transforms.Count;
        // A null buffer is the encoded pixels; only a palette that is not the last stage widens into its own.
        var buffers = new uint[]?[stageCount + 1];
        var widths = new int[stageCount + 1];
        var lags = new int[stageCount + 1];
        buffers[0] = encodedArray;
        widths[0] = encodedWidth;
        int lag = 0;
        bool writesRgba = false;
        for (int stage = 0; stage < stageCount; stage++) {
            Vp8lTransform transform = transforms[stageCount - 1 - stage];
            int width = widths[stage];
            if (transform.OutputWidth != width) return false;
            lags[stage] = lag;
            buffers[stage + 1] = buffers[stage];
            widths[stage + 1] = width;
            switch (transform.Type) {
                case 0:
                    if (transform.Value < 1 || !HasValidVp8lPredictorModes(transform, width, height)) return false;
                    lag++;
                    break;
                case 1:
                    if (transform.Value < 1) return false;
                    break;
                case 2:
                    break;
                case 3:
                    widths[stage + 1] = transform.InputWidth;
                    if (rgba != null && stage == stageCount - 1) {
                        // The last stage can map indexes straight to RGBA bytes.
                        writesRgba = true;
                        break;
                    }
                    int outputLength = checked(transform.InputWidth * height);
                    if (!allocationBudget.TryReserveArray(outputLength, sizeof(uint))) return false;
                    cancellationToken.ThrowIfCancellationRequested();
                    buffers[stage + 1] = AllocateVp8lPixels(outputLength);
                    break;
                default:
                    return false;
            }
        }
        lags[stageCount] = lag;
        long finalPixels = (long)finalWidth * height;
        uint[]? finalBuffer = buffers[stageCount];
        if (widths[stageCount] != finalWidth ||
            !writesRgba && (finalBuffer != null ? finalBuffer.Length : encoded.Length) != finalPixels ||
            rgba != null && rgba.Length != finalPixels * 4L ||
            rgba == null && finalBuffer == null) return false;
        cancellationToken.ThrowIfCancellationRequested();

        uint[] palette = Array.Empty<uint>();
#if NET8_0_OR_GREATER
        uint[]? paletteRow = null;
#endif
        try {
            for (int row = 0; row < height + lag; row++) {
                if ((row & 15) == 0) cancellationToken.ThrowIfCancellationRequested();
                for (int stage = 0; stage < stageCount; stage++) {
                    int y = row - lags[stage];
                    if ((uint)y >= (uint)height) continue;
                    Vp8lTransform transform = transforms[stageCount - 1 - stage];
                    Vp8lPixels pixels = buffers[stage] ?? encoded;
                    int width = widths[stage];
                    switch (transform.Type) {
                        case 0:
                            ApplyVp8lPredictorRow(pixels, width, y, transform);
                            break;
                        case 1:
                            ApplyVp8lColorTransformRow(pixels, width, y, transform);
                            break;
                        case 2:
                            AddVp8lGreenRow(pixels, y * width, width);
                            break;
                        default:
                            if (palette.Length == 0) palette = CreateVp8lPalette(transform, writesRgba && stage == stageCount - 1);
                            if (writesRgba && stage == stageCount - 1) {
#if NET8_0_OR_GREATER
                                // Packed indexes can share memory with the output row they expand into.
                                paletteRow ??= System.Buffers.ArrayPool<uint>.Shared.Rent(width);
                                CopyVp8lPixels(pixels, y * width, paletteRow, 0, width);
                                ApplyVp8lPaletteRowRgba(paletteRow, 0, rgba!, widths[stage + 1], y, transform, palette);
#else
                                ApplyVp8lPaletteRowRgba(pixels, y * width, rgba!, widths[stage + 1], y, transform, palette);
#endif
                            } else {
                                ApplyVp8lPaletteRow(pixels, width, buffers[stage + 1]!, widths[stage + 1], y, transform, palette);
                            }
                            break;
                    }
                }
                if (rgba != null && !writesRgba) {
                    int y = row - lags[stageCount];
                    if ((uint)y < (uint)height) {
                        // Without a palette last, the final words sit at the same offsets as their RGBA bytes.
                        WriteVp8lRgbaRow(finalBuffer ?? encoded, y * finalWidth, finalWidth, rgba);
                    }
                }
            }
        } finally {
#if NET8_0_OR_GREATER
            if (paletteRow != null) System.Buffers.ArrayPool<uint>.Shared.Return(paletteRow);
#endif
        }
        result = finalBuffer ?? Array.Empty<uint>();
        return true;
    }

    private static uint[] AllocateVp8lPixels(int length) {
#if NET8_0_OR_GREATER
        return GC.AllocateUninitializedArray<uint>(length);
#else
        return new uint[length];
#endif
    }

    /// <summary>Rejects undefined predictor modes in blocks that cover at least one predicted pixel.</summary>
    private static bool HasValidVp8lPredictorModes(Vp8lTransform transform, int width, int height) {
        // Width or height one predicts only from the first row or column, which ignore the modes;
        // otherwise every block contains a pixel past the first row and column.
        if (width < 2 || height < 2) return true;
        int blocksHigh = DivideRoundUp(height, 1 << transform.SizeBits);
        int needed = checked(blocksHigh * transform.Value);
        if (transform.Data.Length < needed) return false;
        for (int index = 0; index < needed; index++) {
            if (((transform.Data[index] >> 8) & 255U) > 13U) return false;
        }
        return true;
    }

    private static uint[] CreateVp8lPalette(Vp8lTransform transform, bool rgbaOrder) {
        // Indexes beyond the color table decode as transparent black.
        var palette = new uint[256];
        Array.Copy(transform.Data, palette, Math.Min(256, Math.Min(transform.Value, transform.Data.Length)));
        if (rgbaOrder) {
            // Little-endian words whose bytes are R, G, B, A.
            for (int index = 0; index < palette.Length; index++) {
                uint color = palette[index];
                palette[index] = (color & 0xFF00FF00U) | ((color >> 16) & 255U) | ((color & 255U) << 16);
            }
        }
        return palette;
    }

    private static void ApplyVp8lPaletteRowRgba(Vp8lPixels packed, int source, byte[] rgba, int outputWidth,
        int y, Vp8lTransform transform, uint[] palette) {
        int target = y * outputWidth * 4;
        int pixelsPerPacked = 1 << transform.SizeBits;
        int bitsPerIndex = 8 >> transform.SizeBits;
        uint mask = (1U << bitsPerIndex) - 1U;
        int x = 0;
        for (int packedX = 0; x < outputWidth; packedX++) {
            uint indexes = packed[source + packedX] >> 8;
            int count = Math.Min(pixelsPerPacked, outputWidth - x);
            for (int index = 0; index < count; index++, x++) {
                WriteRgbaWord(rgba, target + x * 4, palette[indexes & mask]);
                indexes >>= bitsPerIndex;
            }
        }
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static void WriteRgbaWord(byte[] rgba, int offset, uint value) {
#if NET8_0_OR_GREATER
        if (BitConverter.IsLittleEndian) {
            System.Runtime.CompilerServices.Unsafe.WriteUnaligned(ref rgba[offset], value);
            return;
        }
#endif
        rgba[offset] = (byte)value;
        rgba[offset + 1] = (byte)(value >> 8);
        rgba[offset + 2] = (byte)(value >> 16);
        rgba[offset + 3] = (byte)(value >> 24);
    }

    private static void ApplyVp8lPaletteRow(Vp8lPixels packed, int packedWidth, uint[] output, int outputWidth,
        int y, Vp8lTransform transform, uint[] palette) {
        int source = y * packedWidth;
        int target = y * outputWidth;
        if (transform.SizeBits == 0) {
            for (int x = 0; x < outputWidth; x++) output[target + x] = palette[(packed[source + x] >> 8) & 255U];
            return;
        }
        int pixelsPerPacked = 1 << transform.SizeBits;
        int bitsPerIndex = 8 >> transform.SizeBits;
        uint mask = (1U << bitsPerIndex) - 1U;
        int x2 = 0;
        for (int packedX = 0; x2 < outputWidth; packedX++) {
            uint indexes = packed[source + packedX] >> 8;
            int count = Math.Min(pixelsPerPacked, outputWidth - x2);
            for (int index = 0; index < count; index++) {
                output[target + x2++] = palette[indexes & mask];
                indexes >>= bitsPerIndex;
            }
        }
    }

    private static void ApplyVp8lColorTransformRow(Vp8lPixels pixels, int width, int y, Vp8lTransform transform) {
        int sizeBits = transform.SizeBits;
        int dataRow = (y >> sizeBits) * transform.Value;
        int row = y * width;
        for (int x = 0; x < width;) {
            int end = Math.Min(width, ((x >> sizeBits) + 1) << sizeBits);
            uint data = transform.Data[dataRow + (x >> sizeBits)];
            // VP8L stores red-to-blue in R, green-to-blue in G, and green-to-red in B.
            int greenToRed = unchecked((sbyte)data);
            int greenToBlue = unchecked((sbyte)(data >> 8));
            int redToBlue = unchecked((sbyte)(data >> 16));
            int index = row + x;
#if NET8_0_OR_GREATER
            index = ApplyVp8lColorTransformVector(pixels, index, row + end, greenToRed, greenToBlue, redToBlue);
#endif
            // Only the low byte of each sum matters, so the unmasked high bits never change a result.
            for (; index < row + end; index++) {
                uint color = pixels[index];
                int green = unchecked((sbyte)(color >> 8));
                int red = (int)(color >> 16) + ((greenToRed * green) >> 5);
                int blue = (int)color + ((greenToBlue * green) >> 5) + ((redToBlue * unchecked((sbyte)red)) >> 5);
                pixels[index] = (color & 0xFF00FF00U) | ((uint)red & 255U) << 16 | ((uint)blue & 255U);
            }
            x = end;
        }
    }

    private static void AddVp8lGreenRow(Vp8lPixels pixels, int start, int count) {
        int index = start;
        int end = start + count;
#if NET8_0_OR_GREATER
        index = AddVp8lGreenVector(pixels, start, end);
#endif
        for (; index < end; index++) {
            uint color = pixels[index];
            uint green = (color >> 8) & 255U;
            uint redBlue = ((color & 0x00FF00FFU) + green * 0x00010001U) & 0x00FF00FFU;
            pixels[index] = (color & 0xFF00FF00U) | redBlue;
        }
    }

    private static void WriteVp8lRgbaRow(Vp8lPixels argb, int start, int count, byte[] rgba) {
        int index = start;
        int end = start + count;
#if NET8_0_OR_GREATER
        index = WriteVp8lRgbaVector(argb, start, end, rgba);
#endif
        for (; index < end; index++) {
            uint color = argb[index];
            int offset = index * 4;
            rgba[offset] = (byte)(color >> 16);
            rgba[offset + 1] = (byte)(color >> 8);
            rgba[offset + 2] = (byte)color;
            rgba[offset + 3] = (byte)(color >> 24);
        }
    }

    private static void ApplyVp8lPredictorRow(Vp8lPixels pixels, int width, int y, Vp8lTransform transform) {
        int row = y * width;
        if (y == 0) {
            uint left = AddArgb(pixels[0], 0xFF000000U);
            pixels[0] = left;
            for (int index = 1; index < width; index++) {
                left = AddArgb(pixels[index], left);
                pixels[index] = left;
            }
            return;
        }
        pixels[row] = AddArgb(pixels[row], pixels[row - width]);
        int sizeBits = transform.SizeBits;
        int dataRow = (y >> sizeBits) * transform.Value;
        for (int x = 1; x < width;) {
            int end = Math.Min(width, ((x >> sizeBits) + 1) << sizeBits);
            int mode = (int)(transform.Data[dataRow + (x >> sizeBits)] >> 8) & 255;
            PredictVp8lRun(pixels, mode, row + x, row + end, width);
            x = end;
        }
    }

    /// <summary>
    /// Restores one run of pixels that share a predictor mode. The top-right neighbor of the last
    /// column is the first pixel of the current row, which is the next word after the previous row.
    /// </summary>
    private static void PredictVp8lRun(Vp8lPixels p, int mode, int start, int end, int width) {
        int index = start;
        switch (mode) {
            case 0:
            case 2:
            case 3:
            case 4:
            case 8:
            case 9:
#if NET8_0_OR_GREATER
                index = PredictVp8lTopRunVector(p, mode, start, end, width);
#endif
                for (; index < end; index++) {
                    uint top = p[index - width];
                    uint prediction = mode switch {
                        0 => 0xFF000000U,
                        2 => top,
                        3 => p[index - width + 1],
                        4 => p[index - width - 1],
                        8 => AverageArgb(p[index - width - 1], top),
                        _ => AverageArgb(top, p[index - width + 1])
                    };
                    p[index] = AddArgb(p[index], prediction);
                }
                return;
        }
        uint left = p[start - 1];
        switch (mode) {
            case 1:
                for (; index < end; index++) p[index] = left = AddArgb(p[index], left);
                return;
            case 5:
                for (; index < end; index++) {
                    uint prediction = AverageArgb(AverageArgb(left, p[index - width + 1]), p[index - width]);
                    p[index] = left = AddArgb(p[index], prediction);
                }
                return;
            case 6:
                for (; index < end; index++) p[index] = left = AddArgb(p[index], AverageArgb(left, p[index - width - 1]));
                return;
            case 7:
                for (; index < end; index++) p[index] = left = AddArgb(p[index], AverageArgb(left, p[index - width]));
                return;
            case 10:
                for (; index < end; index++) {
                    uint prediction = AverageArgb(AverageArgb(left, p[index - width - 1]),
                        AverageArgb(p[index - width], p[index - width + 1]));
                    p[index] = left = AddArgb(p[index], prediction);
                }
                return;
            case 11:
                for (; index < end; index++) {
                    p[index] = left = AddArgb(p[index], SelectArgb(left, p[index - width], p[index - width - 1]));
                }
                return;
            case 12: {
                // Keep the left neighbor split into even and odd channel lanes so the serial chain
                // through each pixel is only the clamp and the residual add; the top-row part of
                // the sum does not depend on the previous pixel.
                const uint Lanes = 0x00FF00FFU;
                uint leftEven = left & Lanes;
                uint leftOdd = (left >> 8) & Lanes;
                for (; index < end; index++) {
                    uint top = p[index - width];
                    uint topLeft = p[index - width - 1];
                    uint residual = p[index];
                    uint even = (top & Lanes) + 0x01000100U - (topLeft & Lanes);
                    uint odd = ((top >> 8) & Lanes) + 0x01000100U - ((topLeft >> 8) & Lanes);
                    leftEven = (ClampBiasedVp8lLanes(even + leftEven) + (residual & Lanes)) & Lanes;
                    leftOdd = (ClampBiasedVp8lLanes(odd + leftOdd) + ((residual >> 8) & Lanes)) & Lanes;
                    p[index] = leftEven | leftOdd << 8;
                }
                return;
            }
            default:
                for (; index < end; index++) {
                    p[index] = left = AddArgb(p[index], ClampAddSubtractHalf(AverageArgb(left, p[index - width]), p[index - width - 1]));
                }
                return;
        }
    }

    /// <summary>
    /// Clamps first + second - subtract per channel without branches. Even and odd channels sit in
    /// separate 16-bit lanes biased by 256, so every lane stays in 1..766: bit 9 marks an overflow
    /// past 255 and a clear bit 8 marks a negative sum.
    /// </summary>
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static uint ClampAddSubtractFull(uint first, uint second, uint subtract) {
        const uint Lanes = 0x00FF00FFU;
        uint even = (first & Lanes) + (second & Lanes) + 0x01000100U - (subtract & Lanes);
        uint odd = ((first >> 8) & Lanes) + ((second >> 8) & Lanes) + 0x01000100U - ((subtract >> 8) & Lanes);
        return ClampBiasedVp8lLanes(even) | ClampBiasedVp8lLanes(odd) << 8;
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static uint ClampBiasedVp8lLanes(uint lanes) {
        uint inRange = (lanes >> 8) & 0x00010001U;
        uint overflow = (lanes >> 9) & 0x00010001U;
        return (lanes & ((inRange << 8) - inRange)) | ((overflow << 8) - overflow);
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static uint ClampAddSubtractHalf(uint average, uint subtract) {
        uint result = 0;
        for (int shift = 0; shift <= 24; shift += 8) {
            int a = (int)((average >> shift) & 255U);
            int value = a + (a - (int)((subtract >> shift) & 255U)) / 2;
            result |= (uint)ClampVp8lChannel(value) << shift;
        }
        return result;
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static int ClampVp8lChannel(int value) => (value & ~255) == 0 ? value : (~value >> 31) & 255;

    internal static uint ApplyVp8lInverseColorTransform(uint color, uint transformData) {
        int green = (int)(color >> 8) & 255;
        int red = (int)(color >> 16) & 255;
        int blue = (int)color & 255;
        // VP8L stores red-to-blue in R, green-to-blue in G, and green-to-red in B.
        int greenToRed = unchecked((sbyte)transformData);
        int greenToBlue = unchecked((sbyte)(transformData >> 8));
        int redToBlue = unchecked((sbyte)(transformData >> 16));
        int restoredRed = (red + ColorTransformDelta(greenToRed, green)) & 255;
        int restoredBlue = (blue + ColorTransformDelta(greenToBlue, green) +
            ColorTransformDelta(redToBlue, restoredRed)) & 255;
        return (color & 0xFF00FF00U) | (uint)(restoredRed << 16 | restoredBlue);
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
        // Independent 16-bit lanes keep carries out of the neighboring color channel.
        uint redBlue = ((first & 0x00FF00FFU) + (second & 0x00FF00FFU)) & 0x00FF00FFU;
        uint alphaGreen = (((first >> 8) & 0x00FF00FFU) + ((second >> 8) & 0x00FF00FFU)) & 0x00FF00FFU;
        return redBlue | (alphaGreen << 8);
    }

    private static uint AverageArgb(uint first, uint second) {
        return (first & second) + (((first ^ second) & 0xFEFEFEFEU) >> 1);
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static uint SelectArgb(uint left, uint top, uint topLeft) {
#if NET8_0_OR_GREATER
        if (System.Runtime.Intrinsics.X86.Sse2.IsSupported)
            return SelectVp8lNeighborVector(left, top, topLeft);
#endif
        int leftDistance = ChannelDistance(top, topLeft);
        int topDistance = ChannelDistance(left, topLeft);
        return leftDistance < topDistance ? left : top;
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static int ChannelDistance(uint first, uint second) =>
        Math.Abs((int)(first & 255U) - (int)(second & 255U)) +
        Math.Abs((int)((first >> 8) & 255U) - (int)((second >> 8) & 255U)) +
        Math.Abs((int)((first >> 16) & 255U) - (int)((second >> 16) & 255U)) +
        Math.Abs((int)(first >> 24) - (int)(second >> 24));

    private static uint ComponentOperation(uint first, uint second, uint subtract, int half) {
        uint result = 0;
        for (int shift = 0; shift <= 24; shift += 8) {
            int a = (int)(first >> shift) & 255;
            int b = (int)(second >> shift) & 255;
            int c = (int)(subtract >> shift) & 255;
            int value = half == 0 ? a + b - c : a + (a - c) / 2;
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
