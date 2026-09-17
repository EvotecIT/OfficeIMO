using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeImageMetadataInspector {
    private static bool IsCanonicalSrgbTiffColorimetry(
        byte[] data,
        bool little,
        int bitsPerSampleEntry,
        int photometricInterpretation,
        int samplesPerPixel,
        int transferFunctionEntry,
        int whitePointEntry,
        int primaryChromaticitiesEntry,
        CancellationToken cancellationToken) {
        if (photometricInterpretation != 2 || samplesPerPixel < 3 || bitsPerSampleEntry < 0 ||
            transferFunctionEntry < 0 || whitePointEntry < 0 || primaryChromaticitiesEntry < 0 ||
            !HasEightBitRgbSamples(data, bitsPerSampleEntry, little, data.Length)) {
            return false;
        }

        double[] expectedWhitePoint = { 0.3127D, 0.3290D };
        double[] expectedPrimaries = { 0.6400D, 0.3300D, 0.3000D, 0.6000D, 0.1500D, 0.0600D };
        if (!HasCanonicalTiffRationals(
                data, whitePointEntry, little, data.Length, expectedWhitePoint) ||
            !HasCanonicalTiffRationals(
                data, primaryChromaticitiesEntry, little, data.Length, expectedPrimaries)) {
            return false;
        }

        if (!TryGetTiffValueRange(
                data,
                transferFunctionEntry,
                little,
                data.Length,
                expectedType: 3,
                itemSize: 2,
                out int transferOffset,
                out uint transferCount) ||
            transferCount != 256U && transferCount != 768U) {
            return false;
        }

        int tableCount = transferCount == 768U ? 3 : 1;
        for (int table = 0; table < tableCount; table++) {
            for (int index = 0; index < 256; index++) {
                if ((index & 63) == 0) cancellationToken.ThrowIfCancellationRequested();
                int actual = ReadUInt16(data, transferOffset + (table * 256 + index) * 2, little);
                double encoded = index / 255D;
                double linear = encoded <= 0.04045D
                    ? encoded / 12.92D
                    : Math.Pow((encoded + 0.055D) / 1.055D, 2.4D);
                int expected = (int)Math.Round(linear * ushort.MaxValue, MidpointRounding.AwayFromZero);
                if (Math.Abs(actual - expected) > 1) return false;
            }
        }
        return true;
    }

    private static bool HasEightBitRgbSamples(
        byte[] data,
        int entry,
        bool little,
        int viewEnd) {
        if (!TryGetTiffValueRange(
                data, entry, little, viewEnd, expectedType: 3, itemSize: 2,
                out int offset, out uint count) || count != 1U && count < 3U) {
            return false;
        }
        int samplesToCheck = count == 1U ? 1 : 3;
        for (int index = 0; index < samplesToCheck; index++) {
            if (ReadUInt16(data, offset + index * 2, little) != 8) return false;
        }
        return true;
    }

    private static bool HasCanonicalTiffRationals(
        byte[] data,
        int entry,
        bool little,
        int viewEnd,
        double[] expected) {
        if (!TryGetTiffValueRange(
                data, entry, little, viewEnd, expectedType: 5, itemSize: 8,
                out int offset, out uint count) || count != (uint)expected.Length) {
            return false;
        }
        for (int index = 0; index < expected.Length; index++) {
            uint numerator = ReadUInt32Unsigned(data, offset + index * 8, little);
            uint denominator = ReadUInt32Unsigned(data, offset + index * 8 + 4, little);
            if (denominator == 0U ||
                Math.Abs(numerator / (double)denominator - expected[index]) > 0.000001D) {
                return false;
            }
        }
        return true;
    }

    private static bool TryGetTiffValueRange(
        byte[] data,
        int entry,
        bool little,
        int viewEnd,
        int expectedType,
        int itemSize,
        out int offset,
        out uint count) {
        offset = 0;
        count = 0U;
        if (viewEnd < 0 || viewEnd > data.Length || entry < 0 || entry > viewEnd - 12 ||
            ReadUInt16(data, entry + 2, little) != expectedType || itemSize <= 0) {
            return false;
        }
        count = ReadUInt32Unsigned(data, entry + 4, little);
        ulong byteCount = (ulong)count * (uint)itemSize;
        if (count == 0U || byteCount > int.MaxValue) return false;
        if (byteCount <= 4U) {
            offset = entry + 8;
        } else {
            uint relativeOffset = ReadUInt32Unsigned(data, entry + 8, little);
            if (relativeOffset > int.MaxValue) return false;
            offset = (int)relativeOffset;
        }
        return offset >= 0 && offset <= viewEnd && byteCount <= (ulong)(viewEnd - offset);
    }
}
