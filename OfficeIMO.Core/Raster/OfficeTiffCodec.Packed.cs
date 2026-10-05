using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    // Expand packed gray samples to device bytes, but retain palette indices.
    // Each stored row starts on a byte boundary, including partial edge tiles.
    private static bool TryDecodePackedSegments(byte[] bytes, IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian, int width, int height, int bits, int photometric, int compression,
        OfficeRasterDecodeOptions options, TiffValidationBudget? budget, bool retainPixels, out byte[] source) {
        source = Array.Empty<byte>();
        bool strips = entries.ContainsKey(273) || entries.ContainsKey(279);
        bool tiles = entries.ContainsKey(324) || entries.ContainsKey(325) || entries.ContainsKey(322) || entries.ContainsKey(323);
        if (strips == tiles) return false;
        int segmentWidth = width, segmentHeight;
        if (strips) {
            if (!TryReadScalarOrDefault(bytes, entries, 278, littleEndian, height, out segmentHeight) || segmentHeight < 1) return false;
            segmentHeight = Math.Min(segmentHeight, height);
        } else if (!TryReadScalar(bytes, entries, 322, littleEndian, out segmentWidth) ||
            !TryReadScalar(bytes, entries, 323, littleEndian, out segmentHeight) ||
            segmentWidth < 1 || segmentHeight < 1) return false;

        int across = checked((int)(((long)width + segmentWidth - 1) / segmentWidth));
        int down = checked((int)(((long)height + segmentHeight - 1) / segmentHeight));
        int count = checked(across * down);
        int rowBytes = checked((int)(((long)segmentWidth * bits + 7) / 8));
        int scratchLength = OfficeRasterGuards.EnsureByteCount((long)rowBytes * segmentHeight,
            "TIFF packed segment exceeds the managed limit.");
        int sourceLength = OfficeRasterGuards.EnsureByteCount((long)width * height,
            "TIFF unpacked samples exceed the managed limit.");
        int rgbaLength = retainPixels ? OfficeRasterGuards.EnsureByteCount((long)width * height * 4,
            "TIFF RGBA output exceeds the managed limit.") : 0;
        long metadataBytes = (long)count * 2 * sizeof(int);
        bool WithinLimit(int compressed) => IsTiffDecodeWorkingSetWithinLimit(bytes.LongLength,
            sourceLength, scratchLength, rgbaLength, metadataBytes, retainPixels, compression,
            compressed, scratchLength, options.RetainedManagedBytes);
        if (!WithinLimit(0) || !TryReadValues(bytes, entries, strips ? 273 : 324, littleEndian, count,
            options.CancellationToken, out int[] offsets) || !TryReadValues(bytes, entries, strips ? 279 : 325,
            littleEndian, count, options.CancellationToken, out int[] lengths)) return false;
        int maximumCompressed = 0;
        for (int segment = 0; segment < count; segment++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!HasSegment(bytes, offsets[segment], lengths[segment])) return false;
            int rows = strips ? Math.Min(segmentHeight, height - segment * segmentHeight) : segmentHeight;
            int expected = checked(rows * rowBytes);
            if (budget != null && !budget.TryReserve(lengths[segment], expected)) return false;
            maximumCompressed = Math.Max(maximumCompressed, lengths[segment]);
        }
        if (!WithinLimit(maximumCompressed)) return false;
        var packed = new byte[scratchLength];
        if (retainPixels) source = new byte[sourceLength];
        int mask = (1 << bits) - 1;
        for (int segment = 0; segment < count; segment++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            int left = checked(segment % across * segmentWidth);
            int top = checked(segment / across * segmentHeight);
            int rows = Math.Min(segmentHeight, height - top);
            int columns = Math.Min(segmentWidth, width - left);
            int expected = checked(rowBytes * (strips ? rows : segmentHeight));
            if (!TryDecodeStrip(bytes, offsets[segment], lengths[segment], compression,
                packed, 0, expected, options.CancellationToken)) return false;
            if (!retainPixels) continue;
            for (int y = 0; y < rows; y++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                for (int x = 0; x < columns; x++) {
                    if ((x & 4095) == 0) options.CancellationToken.ThrowIfCancellationRequested();
                    int bit = checked(x * bits);
                    int value = (packed[y * rowBytes + bit / 8] >> (8 - bits - bit % 8)) & mask;
                    source[(top + y) * width + left + x] = (byte)(photometric == 3 ? value : value * 255 / mask);
                }
            }
        }
        return true;
    }
}
