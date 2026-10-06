using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    // TIFF packs twelve-bit samples most-significant bit first, independently of
    // IFD byte order. Rows restart on byte boundaries; expanded words retain the
    // native 0..4095 range until alpha and color conversion.
    private static bool TryDecodePacked12Segments(byte[] bytes, IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian, int width, int height, int samples, int compression, int planar,
        OfficeRasterDecodeOptions options, TiffValidationBudget? budget, bool retainPixels, out byte[] source) {
        source = Array.Empty<byte>();
        bool strips = entries.ContainsKey(273) || entries.ContainsKey(279);
        bool tiles = entries.ContainsKey(324) || entries.ContainsKey(325) || entries.ContainsKey(322) || entries.ContainsKey(323);
        if (strips == tiles) return false;
        int sw = width, sh;
        if (strips) {
            if (!TryReadRowsPerStrip(bytes, entries, littleEndian, height, out sh)) return false;
            sh = Math.Min(sh, height);
        } else if (!TryReadScalar(bytes, entries, 322, littleEndian, out sw) ||
            !TryReadScalar(bytes, entries, 323, littleEndian, out sh) || sw < 1 || sh < 1) return false;
        int across = checked((int)(((long)width + sw - 1) / sw));
        int down = checked((int)(((long)height + sh - 1) / sh));
        int perPlane = checked(across * down), count = checked(perPlane * (planar == 2 ? samples : 1));
        int channels = planar == 2 ? 1 : samples;
        int rowBytes = checked((int)(((long)sw * channels * 12 + 7) / 8));
        int scratchLength = OfficeRasterGuards.EnsureByteCount((long)rowBytes * sh, "TIFF packed segment exceeds the managed limit.");
        int sourceLength = OfficeRasterGuards.EnsureByteCount((long)width * height * samples * 2, "TIFF unpacked samples exceed the managed limit.");
        int rgbaLength = retainPixels ? OfficeRasterGuards.EnsureByteCount((long)width * height * 4, "TIFF RGBA output exceeds the managed limit.") : 0;
        bool WithinLimit(int compressed) => IsTiffDecodeWorkingSetWithinLimit(bytes.LongLength,
            sourceLength, scratchLength, rgbaLength, (long)count * 8, retainPixels, compression,
            compressed, scratchLength, options.RetainedManagedBytes);
        if (!WithinLimit(0) || !TryReadValues(bytes, entries, strips ? 273 : 324, littleEndian, count,
            options.CancellationToken, out int[] offsets) || !TryReadValues(bytes, entries, strips ? 279 : 325,
            littleEndian, count, options.CancellationToken, out int[] lengths)) return false;
        int maximumCompressed = 0;
        for (int segment = 0; segment < count; segment++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!HasSegment(bytes, offsets[segment], lengths[segment])) return false;
            int top = segment % perPlane / across * sh;
            int expected = checked(rowBytes * (strips ? Math.Min(sh, height - top) : sh));
            if (budget != null && !budget.TryReserve(lengths[segment], expected)) return false;
            maximumCompressed = Math.Max(maximumCompressed, lengths[segment]);
        }
        if (!WithinLimit(maximumCompressed)) return false;
        byte[] packed = new byte[scratchLength];
        if (retainPixels) source = new byte[sourceLength];
        for (int segment = 0; segment < count; segment++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            int tile = segment % perPlane, plane = segment / perPlane;
            int left = tile % across * sw, top = tile / across * sh;
            int rows = Math.Min(sh, height - top), columns = Math.Min(sw, width - left);
            int expected = checked(rowBytes * (strips ? rows : sh));
            if (!TryDecodeStrip(bytes, offsets[segment], lengths[segment], compression,
                packed, 0, expected, options.CancellationToken)) return false;
            if (!retainPixels) continue;
            for (int y = 0; y < rows; y++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                for (int x = 0; x < columns; x++) {
                    if ((x & 4095) == 0) options.CancellationToken.ThrowIfCancellationRequested();
                    for (int c = 0; c < channels; c++) {
                        long bit = ((long)x * channels + c) * 12;
                        int at = checked(y * rowBytes + (int)(bit / 8));
                        int value = (bit & 7) == 0 ? (packed[at] << 4) | (packed[at + 1] >> 4)
                            : ((packed[at] & 15) << 8) | packed[at + 1];
                        int sample = checked(((top + y) * width + left + x) * samples + (planar == 2 ? plane : c));
                        WriteTiffJpegSample(source, sample, 2, littleEndian, value);
                    }
                }
            }
        }
        return true;
    }
}
